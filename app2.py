#!/usr/bin/env python3
"""
INBOX.IQ — Email Analytics Hub
Multi-Account Email Analytics Dashboard with AI Insights
"""

import os
import json
import csv
import io
import re
import imaplib
import email
import sqlite3
import hashlib
import threading
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime
from email.header import decode_header
from email.utils import parsedate_to_datetime
from flask import Flask, render_template_string, request, jsonify, redirect, url_for, make_response
from textblob import TextBlob

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
os.makedirs(TEMPLATE_DIR, exist_ok=True)

app = Flask(__name__, template_folder=TEMPLATE_DIR)
app.secret_key = os.urandom(24)
DB_PATH = os.path.join(BASE_DIR, "email_analytics.db")

def load_template(name):
    path = os.path.join(TEMPLATE_DIR, name)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    raise FileNotFoundError(f"\n\n❌ Template not found: {path}\n")

# ─────────────────────────────────────────────
# DATABASE
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
            sender_domain TEXT DEFAULT '',
            sender_name TEXT DEFAULT '',
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
            phone_numbers TEXT DEFAULT '',
            job_title TEXT DEFAULT '',
            company_name TEXT DEFAULT '',
            campaign TEXT DEFAULT '',
            journey TEXT DEFAULT '',
            FOREIGN KEY(account_id) REFERENCES email_accounts(id),
            UNIQUE(account_id, message_id)
        )
    """)
    existing_cols = [row[1] for row in c.execute("PRAGMA table_info(emails)").fetchall()]
    for col, typ in [
        ("sender_domain", "TEXT DEFAULT ''"),
        ("sender_name", "TEXT DEFAULT ''"),
        ("phone_numbers", "TEXT DEFAULT ''"),
        ("job_title", "TEXT DEFAULT ''"),
        ("company_name", "TEXT DEFAULT ''"),
        ("campaign", "TEXT DEFAULT ''"),
        ("journey", "TEXT DEFAULT ''"),
    ]:
        if col not in existing_cols:
            c.execute(f"ALTER TABLE emails ADD COLUMN {col} {typ}")
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
    pattern = r'(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}'
    phones = re.findall(pattern, text)
    return ", ".join(list(set(phones))[:3]) if phones else ""

def extract_sender_name(sender_str):
    m = re.match(r'^"?([^"<]+)"?\s*<', sender_str)
    if m:
        return m.group(1).strip()
    if "@" not in sender_str:
        return sender_str.strip()
    return ""

def extract_domain(email_addr):
    if "@" in email_addr:
        return email_addr.split("@")[-1].lower()
    return ""

def analyze_sentiment(text):
    if not text or len(text.strip()) < 5:
        return "neutral", 0.0, 0.0, 0.0
    clean = re.sub(r'[^\w\s.,!?]', '', text[:2000])
    blob = TextBlob(clean)
    polarity = blob.sentiment.polarity
    positive_score = max(0, polarity)
    negative_score = max(0, -polarity)
    positive_words = ["thank", "great", "excellent", "good", "happy", "pleased", "wonderful",
                     "appreciate", "love", "amazing", "perfect", "best", "awesome", "fantastic",
                     "congratulations", "success", "well done", "impressive", "outstanding",
                     "interested", "let's connect", "sounds good", "would love to", "yes please"]
    negative_words = ["problem", "issue", "bad", "terrible", "awful", "disappointed", "unhappy",
                     "error", "fail", "wrong", "poor", "unacceptable", "hate", "worst", "complaint",
                     "refund", "cancel", "urgent", "critical", "broken", "bug", "not interested",
                     "unsubscribe", "stop emailing", "remove me"]
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
                sender_email = ""
                m = re.search(r'<([^>]+)>', sender)
                if m:
                    sender_email = m.group(1)
                elif "@" in sender:
                    sender_email = sender.strip()
                sender_name = extract_sender_name(sender)
                sender_domain = extract_domain(sender_email)
                date_str = msg.get("Date", "")
                try:
                    dt = parsedate_to_datetime(date_str)
                    date_iso = dt.strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    date_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                body = get_email_body(msg)
                snippet = body[:200].replace("\n", " ").replace("\r", " ").strip()
                word_count = len(body.split()) if body else 0
                phone_numbers = extract_phone_numbers(body + " " + sender)
                sentiment, score, pos, neg = analyze_sentiment(body or subject)
                c = conn.cursor()
                try:
                    c.execute("""
                        INSERT OR IGNORE INTO emails
                        (account_id, message_id, subject, sender, sender_email, sender_name,
                         sender_domain, recipient, date, body, snippet, folder, sentiment,
                         sentiment_score, sentiment_positive, sentiment_negative, word_count, phone_numbers)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (account_id, msg_id, subject, sender, sender_email, sender_name,
                          sender_domain, recipient, date_iso, body, snippet, folder, sentiment,
                          score, pos, neg, word_count, phone_numbers))
                    conn.commit()
                    if c.rowcount > 0:
                        synced += 1
                except Exception:
                    pass
            except Exception:
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
    accounts = conn.execute("SELECT id, email, display_name, imap_server, added_at, last_synced, total_emails, status FROM email_accounts").fetchall()
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
        return jsonify({"error": f"Authentication failed. ({str(e)})"}), 401
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
        return jsonify({"success": True, "account": dict(account)})
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({"error": "Account already added"}), 409

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
            fetch_emails_from_imap(account_id, account["email"], account["app_password"], account["imap_server"], "INBOX", 300)
        except Exception as e:
            print(f"Sync error: {e}")
    threading.Thread(target=run_sync, daemon=True).start()
    return jsonify({"success": True, "message": f"Syncing {account['email']}..."})

@app.route("/api/sync/all", methods=["POST"])
def sync_all():
    conn = get_db()
    accounts = conn.execute("SELECT * FROM email_accounts WHERE status='active'").fetchall()
    conn.close()
    def run_all():
        for account in accounts:
            try:
                fetch_emails_from_imap(account["id"], account["email"], account["app_password"], account["imap_server"], "INBOX", 300)
            except Exception as e:
                print(f"Sync error: {e}")
    threading.Thread(target=run_all, daemon=True).start()
    return jsonify({"success": True, "message": f"Syncing {len(accounts)} accounts..."})

@app.route("/api/sync/status/<int:account_id>")
def sync_status(account_id):
    conn = get_db()
    account = conn.execute("SELECT id, email, last_synced, total_emails FROM email_accounts WHERE id=?", (account_id,)).fetchone()
    conn.close()
    if not account:
        return jsonify({"error": "Not found"}), 404
    return jsonify(dict(account))

@app.route("/api/stats")
def get_stats():
    account_ids = request.args.getlist("accounts[]") or request.args.getlist("accounts")
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")
    conn = get_db()
    where = []
    params = []
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

    total = conn.execute(f"SELECT COUNT(*) FROM emails {clause}", params).fetchone()[0]
    positive = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {'AND' if clause else 'WHERE'} sentiment='positive'", params).fetchone()[0]
    negative = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {'AND' if clause else 'WHERE'} sentiment='negative'", params).fetchone()[0]
    neutral = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {'AND' if clause else 'WHERE'} sentiment='neutral'", params).fetchone()[0]
    avg_score = conn.execute(f"SELECT AVG(sentiment_score) FROM emails {clause}", params).fetchone()[0] or 0

    top_senders = conn.execute(f"""
        SELECT sender_email, sender_name, sender_domain, COUNT(*) as cnt FROM emails {clause}
        GROUP BY sender_email ORDER BY cnt DESC LIMIT 10
    """, params).fetchall()

    trend = conn.execute(f"""
        SELECT DATE(date) as day, COUNT(*) as total,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN sentiment='negative' THEN 1 ELSE 0 END) as negative,
               SUM(CASE WHEN sentiment='neutral' THEN 1 ELSE 0 END) as neutral
        FROM emails {clause}
        GROUP BY DATE(date) ORDER BY day ASC LIMIT 30
    """, params).fetchall()

    by_account = conn.execute(f"""
        SELECT ea.email, ea.display_name,
               COUNT(*) as total,
               SUM(CASE WHEN e.sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN e.sentiment='negative' THEN 1 ELSE 0 END) as negative,
               SUM(CASE WHEN e.sentiment='neutral' THEN 1 ELSE 0 END) as neutral,
               AVG(e.sentiment_score) as avg_score
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id
        {clause}
        GROUP BY e.account_id
    """, params).fetchall()

    # NEW: Top Subject Lines with reply rate
    subject_reply_rate = conn.execute(f"""
        SELECT subject,
               COUNT(*) as total,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN sentiment='negative' THEN 1 ELSE 0 END) as negative,
               ROUND(100.0 * SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) / COUNT(*), 1) as reply_rate,
               AVG(sentiment_score) as avg_score
        FROM emails {clause}
        WHERE subject IS NOT NULL AND subject != '' AND subject != '(No Subject)'
        GROUP BY subject HAVING total >= 1
        ORDER BY reply_rate DESC, total DESC LIMIT 15
    """, params).fetchall()

    # NEW: Top Domains
    top_domains = conn.execute(f"""
        SELECT sender_domain, COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN sentiment='negative' THEN 1 ELSE 0 END) as negative,
               ROUND(100.0 * SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) / COUNT(*), 1) as reply_rate,
               AVG(sentiment_score) as avg_score
        FROM emails {clause}
        WHERE sender_domain IS NOT NULL AND sender_domain != ''
        GROUP BY sender_domain ORDER BY cnt DESC LIMIT 15
    """, params).fetchall()

    # NEW: Best hour/day
    best_hours = conn.execute(f"""
        SELECT strftime('%H', date) as hour, COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive
        FROM emails {clause}
        GROUP BY hour ORDER BY positive DESC LIMIT 24
    """, params).fetchall()

    best_days = conn.execute(f"""
        SELECT strftime('%w', date) as dow, COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive
        FROM emails {clause}
        GROUP BY dow ORDER BY positive DESC
    """, params).fetchall()

    # NEW: Most replied subject overall
    most_replied_subject = conn.execute(f"""
        SELECT subject, COUNT(*) as cnt, AVG(sentiment_score) as avg_score
        FROM emails {clause} {'AND' if clause else 'WHERE'} sentiment='positive'
        GROUP BY subject ORDER BY cnt DESC LIMIT 1
    """, params).fetchone()

    conn.close()

    day_names = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    best_day_data = [{"day": day_names[int(d["dow"])], "cnt": d["cnt"], "positive": d["positive"]} for d in best_days if d["dow"] is not None]

    return jsonify({
        "total": total,
        "positive": positive,
        "negative": negative,
        "neutral": neutral,
        "avg_score": round(avg_score, 3),
        "top_senders": [dict(s) for s in top_senders],
        "trend": [dict(t) for t in trend],
        "by_account": [dict(a) for a in by_account],
        "subject_reply_rate": [dict(s) for s in subject_reply_rate],
        "top_domains": [dict(d) for d in top_domains],
        "most_replied_subject": dict(most_replied_subject) if most_replied_subject else None,
        "best_hours": [dict(h) for h in best_hours],
        "best_days": best_day_data,
    })

@app.route("/api/emails")
def get_emails():
    account_ids = request.args.getlist("accounts[]") or request.args.getlist("accounts")
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")
    sentiment = request.args.get("sentiment", "")
    search = request.args.get("search", "")
    folder = request.args.get("folder", "INBOX")
    page = int(request.args.get("page", 1))
    per_page = int(request.args.get("per_page", 50))
    conn = get_db()
    where = []
    params = []
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
    total = conn.execute(f"SELECT COUNT(*) FROM emails e {clause}", params).fetchone()[0]
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
        "total": total, "page": page, "per_page": per_page,
        "pages": (total + per_page - 1) // per_page,
        "emails": [dict(e) for e in emails]
    })

@app.route("/api/emails/<int:email_id>")
def get_email(email_id):
    conn = get_db()
    e = conn.execute("""
        SELECT e.*, ea.email as account_email, ea.display_name as account_name
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id WHERE e.id=?
    """, (email_id,)).fetchone()
    conn.close()
    if not e:
        return jsonify({"error": "Not found"}), 404
    return jsonify(dict(e))

# ─────────────────────────────────────────────
# NEW: CSV DOWNLOAD BY SENTIMENT
# ─────────────────────────────────────────────
@app.route("/api/download/csv")
def download_csv():
    sentiment = request.args.get("sentiment", "all")
    account_ids = request.args.getlist("accounts[]") or request.args.getlist("accounts")
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")

    conn = get_db()
    where = []
    params = []
    if account_ids:
        placeholders = ",".join("?" * len(account_ids))
        where.append(f"e.account_id IN ({placeholders})")
        params.extend(account_ids)
    if sentiment and sentiment != "all":
        where.append("e.sentiment = ?")
        params.append(sentiment)
    if date_from:
        where.append("e.date >= ?")
        params.append(date_from + " 00:00:00")
    if date_to:
        where.append("e.date <= ?")
        params.append(date_to + " 23:59:59")
    clause = "WHERE " + " AND ".join(where) if where else ""

    emails = conn.execute(f"""
        SELECT
            e.date as "Date",
            COALESCE(e.sender_name, '') as "Sender Name",
            COALESCE(ea.display_name, '') as "Campaign",
            COALESCE(e.journey, '') as "Journey",
            COALESCE(e.phone_numbers, '') as "Phone Numbers",
            COALESCE(e.job_title, '') as "Job Title",
            COALESCE(e.company_name, '') as "Company Name",
            COALESCE(e.sender_email, '') as "Email ID",
            COALESCE(e.sender_domain, '') as "Domain",
            COALESCE(e.subject, '') as "Subject",
            COALESCE(e.snippet, '') as "Reply Message",
            COALESCE(e.recipient, '') as "Reply To",
            e.sentiment as "Sentiment",
            e.sentiment_score as "Sentiment Score"
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id
        {clause}
        ORDER BY e.date DESC
    """, params).fetchall()
    conn.close()

    output = io.StringIO()
    fieldnames = ["Date","Sender Name","Campaign","Journey","Phone Numbers","Job Title",
                  "Company Name","Email ID","Domain","Subject","Reply Message","Reply To",
                  "Sentiment","Sentiment Score"]
    writer = csv.DictWriter(output, fieldnames=fieldnames)
    writer.writeheader()
    for row in emails:
        writer.writerow(dict(row))

    output.seek(0)
    label = sentiment if sentiment and sentiment != "all" else "all"
    filename = f"inbox_iq_{label}_replies_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    response = make_response(output.getvalue())
    response.headers["Content-Type"] = "text/csv; charset=utf-8"
    response.headers["Content-Disposition"] = f"attachment; filename={filename}"
    return response

# ─────────────────────────────────────────────
# NEW: GOOGLE DRIVE EXPORT
# ─────────────────────────────────────────────
@app.route("/api/export/gdrive", methods=["POST"])
def export_to_gdrive():
    data = request.json or {}
    access_token = data.get("access_token", "")
    sheet_title = data.get("sheet_title", f"INBOX.IQ Export {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    sentiment = data.get("sentiment", "all")
    account_ids = data.get("account_ids", [])
    date_from = data.get("date_from", "")
    date_to = data.get("date_to", "")

    if not access_token:
        return jsonify({"error": "Google access_token required. Please connect Google Drive first."}), 400

    conn = get_db()
    where = []
    params = []
    if account_ids:
        placeholders = ",".join("?" * len(account_ids))
        where.append(f"e.account_id IN ({placeholders})")
        params.extend(account_ids)
    if sentiment and sentiment != "all":
        where.append("e.sentiment = ?")
        params.append(sentiment)
    if date_from:
        where.append("e.date >= ?")
        params.append(date_from + " 00:00:00")
    if date_to:
        where.append("e.date <= ?")
        params.append(date_to + " 23:59:59")
    clause = "WHERE " + " AND ".join(where) if where else ""

    emails = conn.execute(f"""
        SELECT e.date, COALESCE(e.sender_name,'') as sender_name,
               COALESCE(ea.display_name,'') as campaign,
               COALESCE(e.journey,'') as journey,
               COALESCE(e.phone_numbers,'') as phone_numbers,
               COALESCE(e.job_title,'') as job_title,
               COALESCE(e.company_name,'') as company_name,
               COALESCE(e.sender_email,'') as sender_email,
               COALESCE(e.sender_domain,'') as sender_domain,
               COALESCE(e.subject,'') as subject,
               COALESCE(e.snippet,'') as snippet,
               COALESCE(e.recipient,'') as recipient,
               e.sentiment, e.sentiment_score
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id
        {clause} ORDER BY e.date DESC LIMIT 5000
    """, params).fetchall()
    conn.close()

    headers_row = ["Date","Sender Name","Campaign","Journey","Phone Numbers","Job Title",
                   "Company Name","Email ID","Domain","Subject","Reply Message","Reply To",
                   "Sentiment","Sentiment Score"]
    rows = [headers_row]
    for e in emails:
        rows.append([e["date"] or "", e["sender_name"] or "", e["campaign"] or "",
                     e["journey"] or "", e["phone_numbers"] or "", e["job_title"] or "",
                     e["company_name"] or "", e["sender_email"] or "", e["sender_domain"] or "",
                     e["subject"] or "", e["snippet"] or "", e["recipient"] or "",
                     e["sentiment"] or "", str(e["sentiment_score"] or "")])

    def gapi(url, method, body=None):
        payload = json.dumps(body).encode() if body else None
        req = urllib.request.Request(url, data=payload, headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }, method=method)
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                return json.loads(resp.read()), None
        except urllib.error.HTTPError as e:
            return None, f"HTTP {e.code}: {e.read().decode()[:300]}"
        except Exception as e:
            return None, str(e)

    result, err = gapi("https://sheets.googleapis.com/v4/spreadsheets", "POST", {
        "properties": {"title": sheet_title},
        "sheets": [{"properties": {"title": "Email Replies"}}]
    })
    if err:
        return jsonify({"error": f"Could not create Google Sheet: {err}"}), 502

    spreadsheet_id = result["spreadsheetId"]
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"

    _, err2 = gapi(
        f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/Email%20Replies!A1?valueInputOption=USER_ENTERED",
        "PUT", {"values": rows}
    )
    if err2:
        return jsonify({"error": f"Sheet created but data write failed: {err2}", "sheet_url": sheet_url}), 502

    return jsonify({
        "success": True,
        "sheet_url": sheet_url,
        "spreadsheet_id": spreadsheet_id,
        "rows_written": len(rows) - 1,
        "message": f"✅ {len(rows)-1} emails exported to Google Sheets!"
    })

if __name__ == "__main__":
    init_db()
    print("=" * 60)
    print("  📧 INBOX.IQ — Email Analytics Hub")
    print("  Open: http://localhost:5001/dashboard")
    print("=" * 60)
    app.run(debug=True, port=5001, threaded=True)
