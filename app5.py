#!/usr/bin/env python3
"""
Email Analytics Hub - Multi-Account Email Analytics Dashboard
Enhanced with: Domain analytics, subject analytics, sentiment CSV download,
and Google Drive / Google Sheets integration.
"""

import os
import io
import csv
import json
import imaplib
import email
import sqlite3
import hashlib
import threading
import re
from datetime import datetime, timedelta
from email.header import decode_header
from email.utils import parsedate_to_datetime

from flask import (
    Flask, render_template_string, request,
    jsonify, redirect, url_for, session, Response
)
from textblob import TextBlob

# ── Google API imports (optional – gracefully disabled if not installed) ──
try:
    from google.oauth2.credentials import Credentials
    from google.oauth2 import service_account
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request as GRequest
    from googleapiclient.discovery import build
    GOOGLE_AVAILABLE = True
except ImportError:
    GOOGLE_AVAILABLE = False

# ── Resolve template directory relative to this file ──
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
os.makedirs(TEMPLATE_DIR, exist_ok=True)

app = Flask(__name__, template_folder=TEMPLATE_DIR)
app.secret_key = os.urandom(24)
DB_PATH = os.path.join(BASE_DIR, "email_analytics.db")

# Google OAuth scopes
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]
GOOGLE_TOKEN_PATH = os.path.join(BASE_DIR, "google_token.json")
GOOGLE_CREDS_PATH = os.path.join(BASE_DIR, "google_credentials.json")


def load_template(name):
    path = os.path.join(TEMPLATE_DIR, name)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    raise FileNotFoundError(
        f"\n\n  Template not found: {path}\n"
        f"Please make sure '{name}' is inside a 'templates' folder\n"
        f"next to app.py:\n\n"
        f"  {BASE_DIR}/\n"
        f"  app.py\n"
        f"  templates/\n"
        f"      {name}\n"
    )


# ─────────────────────────────────────────────
# DATABASE SETUP
# ─────────────────────────────────────────────
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS email_accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            app_password TEXT NOT NULL,
            imap_server TEXT NOT NULL,
            display_name TEXT,
            added_at TEXT,
            last_synced TEXT,
            total_emails INTEGER DEFAULT 0,
            status TEXT DEFAULT 'active'
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS emails (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            account_id INTEGER NOT NULL,
            message_id TEXT,
            subject TEXT,
            sender TEXT,
            sender_email TEXT,
            sender_name TEXT,
            sender_domain TEXT,
            recipient TEXT,
            date TEXT,
            body TEXT,
            snippet TEXT,
            folder TEXT DEFAULT 'INBOX',
            is_read INTEGER DEFAULT 0,
            sentiment TEXT DEFAULT 'neutral',
            sentiment_score REAL DEFAULT 0.0,
            sentiment_positive REAL DEFAULT 0.0,
            sentiment_negative REAL DEFAULT 0.0,
            word_count INTEGER DEFAULT 0,
            phone_numbers TEXT,
            campaign TEXT,
            journey TEXT,
            job_title TEXT,
            company_name TEXT,
            reply_to TEXT,
            FOREIGN KEY(account_id) REFERENCES email_accounts(id),
            UNIQUE(account_id, message_id)
        )
    """)
    # Migrate existing DB: add new columns if missing
    existing = [row[1] for row in c.execute("PRAGMA table_info(emails)").fetchall()]
    new_cols = {
        "sender_name": "TEXT",
        "sender_domain": "TEXT",
        "phone_numbers": "TEXT",
        "campaign": "TEXT",
        "journey": "TEXT",
        "job_title": "TEXT",
        "company_name": "TEXT",
        "reply_to": "TEXT",
    }
    for col, ctype in new_cols.items():
        if col not in existing:
            c.execute(f"ALTER TABLE emails ADD COLUMN {col} {ctype}")
    conn.commit()
    conn.close()


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


# ─────────────────────────────────────────────
# IMAP HELPERS
# ─────────────────────────────────────────────
IMAP_SERVERS = {
    "gmail.com": "imap.gmail.com",
    "yahoo.com": "imap.mail.yahoo.com",
    "outlook.com": "imap-mail.outlook.com",
    "hotmail.com": "imap-mail.outlook.com",
    "live.com": "imap-mail.outlook.com",
    "icloud.com": "imap.mail.me.com",
    "protonmail.com": "imap.protonmail.ch",
    "zoho.com": "imap.zoho.com",
    "aol.com": "imap.aol.com",
}


def detect_imap_server(email_addr):
    domain = email_addr.split("@")[-1].lower()
    return IMAP_SERVERS.get(domain, f"imap.{domain}")


def decode_str(s):
    if not s:
        return ""
    parts = decode_header(s)
    result = []
    for part, charset in parts:
        if isinstance(part, bytes):
            try:
                result.append(part.decode(charset or "utf-8", errors="replace"))
            except Exception:
                result.append(part.decode("utf-8", errors="replace"))
        else:
            result.append(str(part))
    return " ".join(result)


def get_email_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                try:
                    charset = part.get_content_charset() or "utf-8"
                    body = part.get_payload(decode=True).decode(charset, errors="replace")
                    break
                except Exception:
                    pass
    else:
        try:
            charset = msg.get_content_charset() or "utf-8"
            body = msg.get_payload(decode=True).decode(charset, errors="replace")
        except Exception:
            body = str(msg.get_payload())
    return body.strip()


def extract_phone_numbers(text):
    if not text:
        return ""
    pattern = r'(\+?\d[\d\s\-().]{7,}\d)'
    phones = re.findall(pattern, text)
    cleaned = []
    for p in phones:
        digits = re.sub(r'\D', '', p)
        if 7 <= len(digits) <= 15:
            cleaned.append(p.strip())
    return "; ".join(cleaned[:5])


def extract_metadata_from_body(body):
    meta = {"job_title": "", "company_name": "", "campaign": "", "journey": ""}
    if not body:
        return meta
    lower = body.lower()
    camp_m = re.search(r'campaign[:\s]+([^\n\r|]+)', lower)
    if camp_m:
        meta["campaign"] = camp_m.group(1).strip()[:100]
    journey_m = re.search(r'journey[:\s]+([^\n\r|]+)', lower)
    if journey_m:
        meta["journey"] = journey_m.group(1).strip()[:100]
    title_m = re.search(r'(?:title|position|role)[:\s]+([^\n\r|,]+)', lower)
    if title_m:
        meta["job_title"] = title_m.group(1).strip()[:100]
    company_m = re.search(r'(?:company|organisation|organization|corp|inc|ltd)[:\s]+([^\n\r|,]+)', lower)
    if company_m:
        meta["company_name"] = company_m.group(1).strip()[:100]
    return meta


def analyze_sentiment(text):
    if not text or len(text.strip()) < 5:
        return "neutral", 0.0, 0.0, 0.0
    clean = re.sub(r'[^\w\s.,!?]', '', text[:2000])
    blob = TextBlob(clean)
    polarity = blob.sentiment.polarity
    positive_score = max(0, polarity)
    negative_score = max(0, -polarity)
    positive_words = [
        "thank", "great", "excellent", "good", "happy", "pleased", "wonderful",
        "appreciate", "love", "amazing", "perfect", "best", "awesome", "fantastic",
        "congratulations", "success", "well done", "impressive", "outstanding",
        "interested", "excited", "looking forward", "keen",
    ]
    negative_words = [
        "problem", "issue", "bad", "terrible", "awful", "disappointed", "unhappy",
        "error", "fail", "wrong", "poor", "unacceptable", "hate", "worst", "complaint",
        "refund", "cancel", "urgent", "critical", "broken", "bug", "not interested",
        "unsubscribe", "remove me", "stop emailing",
    ]
    lower_text = text.lower()
    pos_hits = sum(1 for w in positive_words if w in lower_text)
    neg_hits = sum(1 for w in negative_words if w in lower_text)
    boost = (pos_hits - neg_hits) * 0.05
    polarity = max(-1, min(1, polarity + boost))
    positive_score = min(1, positive_score + pos_hits * 0.05)
    negative_score = min(1, negative_score + neg_hits * 0.05)
    if polarity > 0.05:
        label = "positive"
    elif polarity < -0.05:
        label = "negative"
    else:
        label = "neutral"
    return label, round(polarity, 4), round(positive_score, 4), round(negative_score, 4)


def fetch_emails_from_imap(account_id, email_addr, app_password, imap_server, folder="INBOX", limit=200):
    conn = get_db()
    synced = 0
    errors = []
    try:
        mail = imaplib.IMAP4_SSL(imap_server, 993, timeout=30)
        mail.login(email_addr, app_password)
        mail.select(folder)
        _, data = mail.search(None, "ALL")
        message_ids = data[0].split()
        message_ids = message_ids[-limit:]
        for mid in reversed(message_ids):
            try:
                _, msg_data = mail.fetch(mid, "(RFC822)")
                raw = msg_data[0][1]
                msg = email.message_from_bytes(raw)
                msg_id = msg.get("Message-ID", "").strip()
                if not msg_id:
                    msg_id = hashlib.md5(raw[:200]).hexdigest()
                subject = decode_str(msg.get("Subject", "(No Subject)"))
                sender = decode_str(msg.get("From", ""))
                recipient = decode_str(msg.get("To", ""))
                reply_to = decode_str(msg.get("Reply-To", ""))
                sender_email = ""
                sender_name = sender
                m = re.search(r'<([^>]+)>', sender)
                if m:
                    sender_email = m.group(1).strip()
                    sender_name = sender[:sender.index("<")].strip().strip('"')
                elif "@" in sender:
                    sender_email = sender.strip()
                    sender_name = sender.split("@")[0].strip()
                sender_domain = sender_email.split("@")[-1].lower() if "@" in sender_email else ""
                date_str = msg.get("Date", "")
                try:
                    dt = parsedate_to_datetime(date_str)
                    date_iso = dt.strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    date_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                body = get_email_body(msg)
                snippet = body[:200].replace("\n", " ").replace("\r", " ").strip()
                word_count = len(body.split()) if body else 0
                sentiment, score, pos, neg = analyze_sentiment(body or subject)
                phones = extract_phone_numbers(body)
                meta = extract_metadata_from_body(body)
                c = conn.cursor()
                try:
                    c.execute("""
                        INSERT OR IGNORE INTO emails
                        (account_id, message_id, subject, sender, sender_email,
                         sender_name, sender_domain, recipient, date, body, snippet,
                         folder, sentiment, sentiment_score, sentiment_positive,
                         sentiment_negative, word_count, phone_numbers, campaign,
                         journey, job_title, company_name, reply_to)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (
                        account_id, msg_id, subject, sender, sender_email,
                        sender_name, sender_domain, recipient, date_iso, body, snippet,
                        folder, sentiment, score, pos, neg, word_count,
                        phones, meta["campaign"], meta["journey"],
                        meta["job_title"], meta["company_name"], reply_to
                    ))
                    conn.commit()
                    if c.rowcount > 0:
                        synced += 1
                except Exception:
                    pass
            except Exception as e:
                errors.append(str(e))
                continue
        mail.logout()
        c = conn.cursor()
        c.execute("""
            UPDATE email_accounts
            SET last_synced=?, total_emails=(SELECT COUNT(*) FROM emails WHERE account_id=?)
            WHERE id=?
        """, (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), account_id, account_id))
        conn.commit()
    except imaplib.IMAP4.error as e:
        raise Exception(f"IMAP Auth Failed: {str(e)}")
    except Exception as e:
        raise Exception(f"Sync Error: {str(e)}")
    finally:
        conn.close()
    return synced


# ─────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────
@app.route("/")
def index():
    return redirect(url_for("dashboard"))


@app.route("/dashboard")
def dashboard():
    return render_template_string(load_template("dashboard.html"))


@app.route("/api/accounts", methods=["GET"])
def get_accounts():
    conn = get_db()
    accounts = conn.execute(
        "SELECT id, email, display_name, imap_server, added_at, last_synced, total_emails, status FROM email_accounts"
    ).fetchall()
    conn.close()
    return jsonify([dict(a) for a in accounts])


@app.route("/api/accounts", methods=["POST"])
def add_account():
    data = request.json
    email_addr = data.get("email", "").strip().lower()
    app_password = data.get("app_password", "").strip()
    display_name = data.get("display_name", "").strip() or email_addr.split("@")[0]
    imap_server = data.get("imap_server", "").strip() or detect_imap_server(email_addr)
    if not email_addr or not app_password:
        return jsonify({"error": "Email and app password are required"}), 400
    try:
        mail = imaplib.IMAP4_SSL(imap_server, 993, timeout=15)
        mail.login(email_addr, app_password)
        mail.logout()
    except imaplib.IMAP4.error as e:
        return jsonify({"error": f"Authentication failed. Check your app password. ({str(e)})"}), 401
    except Exception as e:
        return jsonify({"error": f"Connection failed to {imap_server}: {str(e)}"}), 500
    conn = get_db()
    try:
        conn.execute("""
            INSERT INTO email_accounts (email, app_password, imap_server, display_name, added_at)
            VALUES (?, ?, ?, ?, ?)
        """, (email_addr, app_password, imap_server, display_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        account = conn.execute("SELECT * FROM email_accounts WHERE email=?", (email_addr,)).fetchone()
        conn.close()
        return jsonify({"success": True, "account": dict(account), "message": "Account added successfully!"})
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({"error": "This email account is already added"}), 409


@app.route("/api/accounts/<int:account_id>", methods=["DELETE"])
def delete_account(account_id):
    conn = get_db()
    conn.execute("DELETE FROM emails WHERE account_id=?", (account_id,))
    conn.execute("DELETE FROM email_accounts WHERE id=?", (account_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})


@app.route("/api/sync/<int:account_id>", methods=["POST"])
def sync_account(account_id):
    conn = get_db()
    account = conn.execute("SELECT * FROM email_accounts WHERE id=?", (account_id,)).fetchone()
    conn.close()
    if not account:
        return jsonify({"error": "Account not found"}), 404

    def run_sync():
        try:
            fetch_emails_from_imap(
                account_id, account["email"], account["app_password"],
                account["imap_server"], "INBOX", 300
            )
        except Exception as e:
            print(f"Sync error for {account['email']}: {e}")

    t = threading.Thread(target=run_sync)
    t.daemon = True
    t.start()
    return jsonify({"success": True, "message": f"Syncing {account['email']} in background..."})


@app.route("/api/sync/all", methods=["POST"])
def sync_all():
    conn = get_db()
    accounts = conn.execute("SELECT * FROM email_accounts WHERE status='active'").fetchall()
    conn.close()

    def run_all():
        for account in accounts:
            try:
                fetch_emails_from_imap(
                    account["id"], account["email"], account["app_password"],
                    account["imap_server"], "INBOX", 300
                )
            except Exception as e:
                print(f"Sync error: {e}")

    t = threading.Thread(target=run_all)
    t.daemon = True
    t.start()
    return jsonify({"success": True, "message": f"Syncing {len(accounts)} account(s) in background..."})


@app.route("/api/sync/status/<int:account_id>")
def sync_status(account_id):
    conn = get_db()
    account = conn.execute(
        "SELECT id, email, last_synced, total_emails FROM email_accounts WHERE id=?", (account_id,)
    ).fetchone()
    conn.close()
    if not account:
        return jsonify({"error": "Not found"}), 404
    return jsonify(dict(account))


# ─────────────────────────────────────────────
# STATS API (enhanced)
# ─────────────────────────────────────────────
@app.route("/api/stats")
def get_stats():
    account_ids = request.args.getlist("accounts[]") or request.args.getlist("accounts")
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")

    conn = get_db()
    where, params = [], []

    if account_ids:
        placeholders = ",".join("?" * len(account_ids))
        where.append(f"account_id IN ({placeholders})")
        params.extend(account_ids)
    if date_from:
        where.append("date >= ?")
        params.append(date_from + " 00:00:00")
    if date_to:
        where.append("date <= ?")
        params.append(date_to + " 23:59:59")

    clause = "WHERE " + " AND ".join(where) if where else ""
    and_kw = "AND" if clause else "WHERE"

    total    = conn.execute(f"SELECT COUNT(*) FROM emails {clause}", params).fetchone()[0]
    positive = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {and_kw} sentiment='positive'", params).fetchone()[0]
    negative = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {and_kw} sentiment='negative'", params).fetchone()[0]
    neutral  = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {and_kw} sentiment='neutral'",  params).fetchone()[0]
    avg_score = conn.execute(f"SELECT AVG(sentiment_score) FROM emails {clause}", params).fetchone()[0] or 0

    # Top senders
    top_senders = conn.execute(f"""
        SELECT sender_email, sender_name, COUNT(*) as cnt FROM emails {clause}
        GROUP BY sender_email ORDER BY cnt DESC LIMIT 10
    """, params).fetchall()

    # Top reply subject lines
    top_subjects = conn.execute(f"""
        SELECT subject, COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN sentiment='negative' THEN 1 ELSE 0 END) as negative,
               SUM(CASE WHEN sentiment='neutral'  THEN 1 ELSE 0 END) as neutral
        FROM emails {clause}
        {and_kw if not clause else 'AND'} subject IS NOT NULL AND subject != '' AND subject != '(No Subject)'
        GROUP BY subject ORDER BY cnt DESC LIMIT 15
    """, params).fetchall()

    # Domain analytics
    top_domains = conn.execute(f"""
        SELECT sender_domain,
               COUNT(*) as total,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN sentiment='negative' THEN 1 ELSE 0 END) as negative,
               SUM(CASE WHEN sentiment='neutral'  THEN 1 ELSE 0 END) as neutral,
               COUNT(DISTINCT sender_email) as unique_senders
        FROM emails {clause}
        {and_kw if not clause else 'AND'} sender_domain IS NOT NULL AND sender_domain != ''
        GROUP BY sender_domain ORDER BY total DESC LIMIT 20
    """, params).fetchall()

    # Daily trend
    trend = conn.execute(f"""
        SELECT DATE(date) as day, COUNT(*) as total,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN sentiment='negative' THEN 1 ELSE 0 END) as negative,
               SUM(CASE WHEN sentiment='neutral'  THEN 1 ELSE 0 END) as neutral
        FROM emails {clause}
        GROUP BY DATE(date) ORDER BY day ASC LIMIT 60
    """, params).fetchall()

    # Sentiment by account (needs e. prefix)
    e_where = []
    e_params = []
    if account_ids:
        placeholders = ",".join("?" * len(account_ids))
        e_where.append(f"e.account_id IN ({placeholders})")
        e_params.extend(account_ids)
    if date_from:
        e_where.append("e.date >= ?")
        e_params.append(date_from + " 00:00:00")
    if date_to:
        e_where.append("e.date <= ?")
        e_params.append(date_to + " 23:59:59")
    e_clause = "WHERE " + " AND ".join(e_where) if e_where else ""

    by_account = conn.execute(f"""
        SELECT ea.email, ea.display_name,
               COUNT(*) as total,
               SUM(CASE WHEN e.sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN e.sentiment='negative' THEN 1 ELSE 0 END) as negative,
               SUM(CASE WHEN e.sentiment='neutral'  THEN 1 ELSE 0 END) as neutral,
               AVG(e.sentiment_score) as avg_score
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id
        {e_clause}
        GROUP BY e.account_id
    """, e_params).fetchall()

    # Hour-of-day distribution
    hourly = conn.execute(f"""
        SELECT CAST(strftime('%H', date) AS INTEGER) as hour, COUNT(*) as cnt
        FROM emails {clause}
        GROUP BY hour ORDER BY hour
    """, params).fetchall()

    # Weekday distribution
    weekday = conn.execute(f"""
        SELECT CASE strftime('%w', date)
            WHEN '0' THEN 'Sun' WHEN '1' THEN 'Mon' WHEN '2' THEN 'Tue'
            WHEN '3' THEN 'Wed' WHEN '4' THEN 'Thu' WHEN '5' THEN 'Fri'
            ELSE 'Sat' END as day_name,
            strftime('%w', date) as day_num,
            COUNT(*) as cnt
        FROM emails {clause}
        GROUP BY day_num ORDER BY day_num
    """, params).fetchall()

    # Response length distribution
    length_dist = conn.execute(f"""
        SELECT
            SUM(CASE WHEN word_count < 20  THEN 1 ELSE 0 END) as very_short,
            SUM(CASE WHEN word_count BETWEEN 20 AND 99  THEN 1 ELSE 0 END) as short,
            SUM(CASE WHEN word_count BETWEEN 100 AND 299 THEN 1 ELSE 0 END) as medium,
            SUM(CASE WHEN word_count >= 300 THEN 1 ELSE 0 END) as long
        FROM emails {clause}
    """, params).fetchone()

    conn.close()

    return jsonify({
        "total": total,
        "positive": positive,
        "negative": negative,
        "neutral": neutral,
        "avg_score": round(avg_score, 3),
        "top_senders":  [dict(s) for s in top_senders],
        "top_subjects":  [dict(s) for s in top_subjects],
        "top_domains":   [dict(d) for d in top_domains],
        "trend":         [dict(t) for t in trend],
        "by_account":    [dict(a) for a in by_account],
        "hourly":        [dict(h) for h in hourly],
        "weekday":       [dict(w) for w in weekday],
        "length_dist":   dict(length_dist) if length_dist else {},
    })


# ─────────────────────────────────────────────
# EMAILS LIST API
# ─────────────────────────────────────────────
@app.route("/api/emails")
def get_emails():
    account_ids = request.args.getlist("accounts[]") or request.args.getlist("accounts")
    date_from   = request.args.get("date_from", "")
    date_to     = request.args.get("date_to", "")
    sentiment   = request.args.get("sentiment", "")
    search      = request.args.get("search", "")
    folder      = request.args.get("folder", "INBOX")
    page        = int(request.args.get("page", 1))
    per_page    = int(request.args.get("per_page", 50))

    conn = get_db()
    where, params = [], []

    if account_ids:
        placeholders = ",".join("?" * len(account_ids))
        where.append(f"e.account_id IN ({placeholders})")
        params.extend(account_ids)
    if date_from:
        where.append("e.date >= ?")
        params.append(date_from + " 00:00:00")
    if date_to:
        where.append("e.date <= ?")
        params.append(date_to + " 23:59:59")
    if sentiment:
        where.append("e.sentiment = ?")
        params.append(sentiment)
    if search:
        where.append("(e.subject LIKE ? OR e.sender LIKE ? OR e.snippet LIKE ?)")
        params.extend([f"%{search}%", f"%{search}%", f"%{search}%"])
    if folder:
        where.append("e.folder = ?")
        params.append(folder)

    clause = "WHERE " + " AND ".join(where) if where else ""
    total  = conn.execute(f"SELECT COUNT(*) FROM emails e {clause}", params).fetchone()[0]
    offset = (page - 1) * per_page
    emails = conn.execute(f"""
        SELECT e.*, ea.email as account_email, ea.display_name as account_name
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id
        {clause}
        ORDER BY e.date DESC
        LIMIT ? OFFSET ?
    """, params + [per_page, offset]).fetchall()
    conn.close()
    return jsonify({
        "total": total,
        "page": page,
        "per_page": per_page,
        "pages": (total + per_page - 1) // per_page,
        "emails": [dict(e) for e in emails],
    })


@app.route("/api/emails/<int:email_id>")
def get_email(email_id):
    conn = get_db()
    e = conn.execute("""
        SELECT e.*, ea.email as account_email, ea.display_name as account_name
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id
        WHERE e.id=?
    """, (email_id,)).fetchone()
    conn.close()
    if not e:
        return jsonify({"error": "Not found"}), 404
    return jsonify(dict(e))


# ─────────────────────────────────────────────
# CSV DOWNLOAD (date-wise, by sentiment)
# ─────────────────────────────────────────────
EXPORT_COLUMNS = [
    "date", "sender_name", "campaign", "journey", "phone_numbers",
    "job_title", "company_name", "sender_email", "sender_domain",
    "subject", "snippet", "reply_to", "sentiment", "sentiment_score",
    "account_email",
]

EXPORT_HEADERS = [
    "Date", "Sender Name", "Campaign", "Journey", "Phone Numbers",
    "Job Title", "Company Name", "Email ID", "Domain",
    "Subject", "Reply Message", "Reply To", "Sentiment", "Sentiment Score",
    "Account Email",
]


@app.route("/api/download/csv")
def download_csv():
    """
    Download filtered emails as CSV.
    Query params: sentiment (positive|negative|neutral|all), date_from, date_to, accounts[]
    """
    account_ids = request.args.getlist("accounts[]") or request.args.getlist("accounts")
    date_from   = request.args.get("date_from", "")
    date_to     = request.args.get("date_to", "")
    sentiment   = request.args.get("sentiment", "")

    conn = get_db()
    where, params = [], []

    if account_ids:
        placeholders = ",".join("?" * len(account_ids))
        where.append(f"e.account_id IN ({placeholders})")
        params.extend(account_ids)
    if date_from:
        where.append("e.date >= ?")
        params.append(date_from + " 00:00:00")
    if date_to:
        where.append("e.date <= ?")
        params.append(date_to + " 23:59:59")
    if sentiment and sentiment != "all":
        where.append("e.sentiment = ?")
        params.append(sentiment)

    clause = "WHERE " + " AND ".join(where) if where else ""

    rows = conn.execute(f"""
        SELECT e.date, e.sender_name, e.campaign, e.journey, e.phone_numbers,
               e.job_title, e.company_name, e.sender_email, e.sender_domain,
               e.subject, e.snippet, e.reply_to, e.sentiment, e.sentiment_score,
               ea.email as account_email
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id
        {clause}
        ORDER BY e.date DESC
    """, params).fetchall()
    conn.close()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(EXPORT_HEADERS)
    for row in rows:
        writer.writerow([row[col] or "" for col in EXPORT_COLUMNS])

    filename_parts = ["emails"]
    if sentiment and sentiment != "all":
        filename_parts.append(sentiment)
    if date_from:
        filename_parts.append(date_from)
    if date_to:
        filename_parts.append(f"to_{date_to}")
    filename = "_".join(filename_parts) + ".csv"

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )




# ─────────────────────────────────────────────
# AI SUGGESTIONS API
# ─────────────────────────────────────────────
@app.route("/api/ai-suggestions")
def ai_suggestions():
    account_ids = request.args.getlist("accounts[]") or request.args.getlist("accounts")
    date_from = request.args.get("date_from", "")
    date_to   = request.args.get("date_to", "")

    conn = get_db()
    where, params = [], []
    if account_ids:
        placeholders = ",".join("?" * len(account_ids))
        where.append(f"account_id IN ({placeholders})")
        params.extend(account_ids)
    if date_from:
        where.append("date >= ?")
        params.append(date_from + " 00:00:00")
    if date_to:
        where.append("date <= ?")
        params.append(date_to + " 23:59:59")
    clause = "WHERE " + " AND ".join(where) if where else ""
    and_kw = "AND" if clause else "WHERE"

    # ── Best hour & day (by positive reply volume) ──
    hourly = conn.execute(f"""
        SELECT CAST(strftime('%H', date) AS INTEGER) as hour, COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as pos
        FROM emails {clause}
        GROUP BY hour ORDER BY pos DESC, cnt DESC LIMIT 1
    """, params).fetchone()

    weekday_best = conn.execute(f"""
        SELECT CASE strftime('%w', date)
            WHEN '0' THEN 'Sunday' WHEN '1' THEN 'Monday' WHEN '2' THEN 'Tuesday'
            WHEN '3' THEN 'Wednesday' WHEN '4' THEN 'Thursday' WHEN '5' THEN 'Friday'
            ELSE 'Saturday' END as day_name,
            COUNT(*) as cnt,
            SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as pos
        FROM emails {clause}
        GROUP BY day_name ORDER BY pos DESC, cnt DESC LIMIT 1
    """, params).fetchone()

    avg_score = conn.execute(f"SELECT AVG(sentiment_score) FROM emails {clause}", params).fetchone()[0] or 0
    total     = conn.execute(f"SELECT COUNT(*) FROM emails {clause}", params).fetchone()[0]
    positive  = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {and_kw} sentiment='positive'", params).fetchone()[0]
    negative  = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {and_kw} sentiment='negative'", params).fetchone()[0]
    neutral   = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {and_kw} sentiment='neutral'",  params).fetchone()[0]

    # ── Top subject lines with reply rates ──
    top_subjects = conn.execute(f"""
        SELECT subject,
               COUNT(*) as total,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN sentiment='negative' THEN 1 ELSE 0 END) as negative,
               SUM(CASE WHEN sentiment='neutral'  THEN 1 ELSE 0 END) as neutral
        FROM emails {clause}
        {and_kw if not clause else "AND"} subject IS NOT NULL AND subject != '' AND subject != '(No Subject)'
        GROUP BY subject ORDER BY total DESC LIMIT 10
    """, params).fetchall()

    # ── Phrase extraction from positive / negative bodies ──
    pos_bodies = conn.execute(f"""
        SELECT body FROM emails {clause} {and_kw} sentiment='positive' AND body IS NOT NULL LIMIT 200
    """, params).fetchall()
    neg_bodies = conn.execute(f"""
        SELECT body FROM emails {clause} {and_kw} sentiment='negative' AND body IS NOT NULL LIMIT 200
    """, params).fetchall()
    conn.close()

    def extract_phrases(rows, min_len=3, top_n=10):
        stop = {"the","a","an","and","or","but","in","on","at","to","for","of","is","it",
                "i","we","you","my","your","our","this","that","with","have","be","are",
                "was","will","would","can","could","do","did","not","so","if","as","from",
                "by","about","up","also","just","more","me","us","he","she","they","them",
                "how","what","when","where","which","who","all","any","been","had","has",
                "hi","hello","thanks","thank","please","dear","best","regards","sincerely"}
        bigram_counts = {}
        for row in rows:
            text = (row["body"] or "").lower()
            text = re.sub(r"[^a-z0-9\s]", " ", text)
            words = [w for w in text.split() if w not in stop and len(w) >= min_len]
            for j in range(len(words) - 1):
                bg = words[j] + " " + words[j+1]
                bigram_counts[bg] = bigram_counts.get(bg, 0) + 1
        sorted_bg = sorted(bigram_counts.items(), key=lambda x: -x[1])
        return [{"phrase": k, "count": v} for k, v in sorted_bg[:top_n] if v >= 2]

    positive_phrases = extract_phrases(pos_bodies)
    negative_phrases = extract_phrases(neg_bodies)

    # ── Subject performance with reply-rate calc ──
    subject_perf = []
    for s in top_subjects:
        subject_perf.append({
            "subject": s["subject"],
            "total": s["total"],
            "positive": s["positive"],
            "negative": s["negative"],
            "neutral": s["neutral"],
            "reply_rate": round(s["positive"] / s["total"] * 100, 1) if s["total"] else 0,
        })
    subject_perf.sort(key=lambda x: -x["reply_rate"])

    # ── Smart suggestions generated from real data ──
    suggestions = []
    if hourly and hourly["pos"] > 0:
        h = hourly["hour"]
        hour_str = f"{h}:00" if h >= 10 else f"0{h}:00"
        suggestions.append({
            "icon": "⏰",
            "title": f"Schedule sends 1–2 hours before {hour_str}",
            "text": f"Replies peak at {hour_str} UTC — sending earlier increases the chance of landing in an active inbox."
        })
    if weekday_best:
        suggestions.append({
            "icon": "📅",
            "title": f"{weekday_best['day_name']} drives the most positive replies",
            "text": f"Prioritize outreach on {weekday_best['day_name']}s — your data shows {weekday_best['pos']} positive replies that day."
        })
    if subject_perf:
        best = subject_perf[0]
        subj_short = best["subject"][:50] + ("..." if len(best["subject"]) > 50 else "")
        suggestions.append({
            "icon": "📧",
            "title": f"Replicate the hook in \"{subj_short}\"",
            "text": f"This subject line achieves {best['reply_rate']}% positive reply rate — study its framing and use it across other sequences."
        })
    if positive_phrases:
        phrases_str = ", ".join([f'"{p["phrase"]}"' for p in positive_phrases[:3]])
        suggestions.append({
            "icon": "💬",
            "title": "Listen for buying signals in replies",
            "text": f"Positive replies often include phrases like {phrases_str} — flag threads containing these for immediate follow-up."
        })
    if total > 0 and positive / total < 0.2:
        suggestions.append({
            "icon": "🎯",
            "title": "Positive reply rate is below 20% — review your CTA",
            "text": "Try a softer call-to-action, personalise the opening line, or A/B test two subject-line variants this week."
        })

    # ── Improvement tips from sentiment breakdown ──
    improvement_tips = []
    if total > 0:
        neg_pct = round(negative / total * 100)
        if neg_pct > 20:
            improvement_tips.append({
                "icon": "⚠️",
                "title": f"{neg_pct}% of replies are negative — reduce send frequency",
                "text": "High negative sentiment often signals over-sending or irrelevant targeting. Consider pausing non-responders after 3 touches."
            })
        if neutral > positive:
            improvement_tips.append({
                "icon": "💡",
                "title": "Most replies are neutral — add a clearer value proposition",
                "text": "Neutral responses suggest leads read your emails but aren't compelled. Lead with the specific outcome you deliver, not features."
            })
        if avg_score < 0:
            improvement_tips.append({
                "icon": "📝",
                "title": "Overall tone is negative — rewrite sequence openers",
                "text": "Use warmer, curiosity-driven openers. Avoid words like 'urgent', 'final notice', or 'last chance' in cold outreach."
            })
        elif avg_score > 0.15:
            improvement_tips.append({
                "icon": "🚀",
                "title": "Positive sentiment is strong — scale this sequence",
                "text": "Your messaging is landing well. Consider expanding your outreach volume or testing new ICPs with a similar approach."
            })
    if negative_phrases:
        top_neg = negative_phrases[0]["phrase"]
        improvement_tips.append({
            "icon": "🛑",
            "title": f"Address the objection around \"{top_neg}\"",
            "text": "This phrase appears frequently in negative replies. Add a proactive objection-handling line to your follow-up sequences."
        })
    if not improvement_tips:
        improvement_tips.append({
            "icon": "✅",
            "title": "Your sequences look healthy!",
            "text": "Sentiment is balanced. Keep testing subject lines and refining targeting to maintain positive reply rates."
        })

    return jsonify({
        "best_hour": hourly["hour"] if hourly else None,
        "best_hour_replies": hourly["pos"] if hourly else 0,
        "best_day": weekday_best["day_name"] if weekday_best else None,
        "best_day_replies": weekday_best["pos"] if weekday_best else 0,
        "avg_score": round(float(avg_score), 3),
        "total": total,
        "positive": positive,
        "negative": negative,
        "neutral": neutral,
        "positive_phrases": positive_phrases,
        "negative_phrases": negative_phrases,
        "subject_performance": subject_perf,
        "suggestions": suggestions,
        "improvement_tips": improvement_tips,
    })

# ─────────────────────────────────────────────
# GOOGLE DRIVE / SHEETS INTEGRATION
# ─────────────────────────────────────────────

def get_google_creds():
    if not GOOGLE_AVAILABLE:
        return None, "Google API libraries not installed. Run: pip install google-auth google-auth-oauthlib google-api-python-client"
    creds = None
    if os.path.exists(GOOGLE_TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(GRequest())
            except Exception as e:
                return None, f"Token refresh failed: {e}"
        else:
            if not os.path.exists(GOOGLE_CREDS_PATH):
                return None, (
                    "Google credentials file not found. "
                    "Please place your OAuth2 credentials JSON at: " + GOOGLE_CREDS_PATH
                )
            flow = InstalledAppFlow.from_client_secrets_file(GOOGLE_CREDS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(GOOGLE_TOKEN_PATH, "w") as f:
            f.write(creds.to_json())
    return creds, None


@app.route("/api/google/status")
def google_status():
    if not GOOGLE_AVAILABLE:
        return jsonify({
            "available": False,
            "message": "Google API libraries not installed.",
            "install": "pip install google-auth google-auth-oauthlib google-api-python-client",
        })
    token_exists = os.path.exists(GOOGLE_TOKEN_PATH)
    creds_exists = os.path.exists(GOOGLE_CREDS_PATH)
    return jsonify({
        "available": True,
        "authenticated": token_exists,
        "credentials_file": creds_exists,
        "credentials_path": GOOGLE_CREDS_PATH,
    })


def _fetch_export_rows(account_ids, date_from, date_to, sentiment, limit=5000):
    conn = get_db()
    where, params = [], []
    if account_ids:
        placeholders = ",".join("?" * len(account_ids))
        where.append(f"e.account_id IN ({placeholders})")
        params.extend(account_ids)
    if date_from:
        where.append("e.date >= ?")
        params.append(date_from + " 00:00:00")
    if date_to:
        where.append("e.date <= ?")
        params.append(date_to + " 23:59:59")
    if sentiment and sentiment != "all":
        where.append("e.sentiment = ?")
        params.append(sentiment)
    clause = "WHERE " + " AND ".join(where) if where else ""
    rows = conn.execute(f"""
        SELECT e.date, e.sender_name, e.campaign, e.journey, e.phone_numbers,
               e.job_title, e.company_name, e.sender_email, e.sender_domain,
               e.subject, e.snippet, e.reply_to, e.sentiment, e.sentiment_score,
               ea.email as account_email
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id
        {clause}
        ORDER BY e.date DESC
        LIMIT {limit}
    """, params).fetchall()
    conn.close()
    return rows


@app.route("/api/google/export", methods=["POST"])
def google_export():
    """Create a new Google Sheet and populate it with filtered email data."""
    data        = request.json or {}
    account_ids = data.get("accounts", [])
    date_from   = data.get("date_from", "")
    date_to     = data.get("date_to", "")
    sentiment   = data.get("sentiment", "")
    sheet_name  = data.get("sheet_name", "") or f"Email Analytics {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    creds, err = get_google_creds()
    if err:
        return jsonify({"error": err}), 500

    rows = _fetch_export_rows(account_ids, date_from, date_to, sentiment)

    try:
        sheets_service = build("sheets", "v4", credentials=creds)

        spreadsheet = sheets_service.spreadsheets().create(body={
            "properties": {"title": sheet_name},
            "sheets": [{"properties": {"title": "Email Data"}}],
        }).execute()

        spreadsheet_id = spreadsheet["spreadsheetId"]
        spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"

        values = [EXPORT_HEADERS]
        for row in rows:
            values.append([str(row[col] or "") for col in EXPORT_COLUMNS])

        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range="Email Data!A1",
            valueInputOption="RAW",
            body={"values": values},
        ).execute()

        # Format: bold header, freeze row, auto-resize columns
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [
                {
                    "repeatCell": {
                        "range": {"sheetId": 0, "startRowIndex": 0, "endRowIndex": 1},
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 0.2, "green": 0.4, "blue": 0.8},
                                "textFormat": {
                                    "bold": True,
                                    "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                                },
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat)",
                    }
                },
                {
                    "autoResizeDimensions": {
                        "dimensions": {
                            "sheetId": 0,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": len(EXPORT_HEADERS),
                        }
                    }
                },
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": 0,
                            "gridProperties": {"frozenRowCount": 1},
                        },
                        "fields": "gridProperties.frozenRowCount",
                    }
                },
            ]},
        ).execute()

        return jsonify({
            "success": True,
            "sheet_url": spreadsheet_url,
            "spreadsheet_id": spreadsheet_id,
            "rows_exported": len(rows),
            "sheet_name": sheet_name,
        })

    except Exception as e:
        return jsonify({"error": f"Google Sheets export failed: {str(e)}"}), 500


@app.route("/api/google/append", methods=["POST"])
def google_append():
    """Append filtered email data to an existing Google Sheet."""
    data           = request.json or {}
    spreadsheet_id = data.get("spreadsheet_id", "")
    account_ids    = data.get("accounts", [])
    date_from      = data.get("date_from", "")
    date_to        = data.get("date_to", "")
    sentiment      = data.get("sentiment", "")

    if not spreadsheet_id:
        return jsonify({"error": "spreadsheet_id is required"}), 400

    creds, err = get_google_creds()
    if err:
        return jsonify({"error": err}), 500

    rows = _fetch_export_rows(account_ids, date_from, date_to, sentiment)

    try:
        sheets_service = build("sheets", "v4", credentials=creds)
        values = [[str(row[col] or "") for col in EXPORT_COLUMNS] for row in rows]
        sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range="A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": values},
        ).execute()
        return jsonify({
            "success": True,
            "rows_appended": len(rows),
            "sheet_url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}",
        })
    except Exception as e:
        return jsonify({"error": f"Append failed: {str(e)}"}), 500


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    init_db()
    print("=" * 60)
    print("  Email Analytics Hub  (Enhanced)")
    print("  Open: http://localhost:5001/dashboard")
    print("=" * 60)
    if not GOOGLE_AVAILABLE:
        print("  Google Drive disabled.")
        print("  To enable: pip install google-auth google-auth-oauthlib google-api-python-client")
    app.run(debug=True, port=5001, threaded=True)
