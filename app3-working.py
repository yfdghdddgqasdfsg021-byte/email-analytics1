#!/usr/bin/env python3
"""
INBOX.IQ — Email Analytics Hub (upgraded)
"""

import os, json, imaplib, email, sqlite3, hashlib, threading, csv, io, re
from datetime import datetime, timedelta
from email.header import decode_header
from email.utils import parsedate_to_datetime
from collections import Counter
from flask import Flask, render_template_string, request, jsonify, redirect, Response
from textblob import TextBlob

# ── Paths ──────────────────────────────────────────────────────────────
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
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
    raise FileNotFoundError(f"Template not found: {path}\nPut {name} inside a 'templates/' folder next to app.py")

# ── Database ───────────────────────────────────────────────────────────
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS email_accounts (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            email         TEXT UNIQUE NOT NULL,
            app_password  TEXT NOT NULL,
            imap_server   TEXT NOT NULL,
            display_name  TEXT,
            added_at      TEXT,
            last_synced   TEXT,
            total_emails  INTEGER DEFAULT 0,
            status        TEXT DEFAULT 'active'
        )""")
    c.execute("""
        CREATE TABLE IF NOT EXISTS emails (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            account_id         INTEGER NOT NULL,
            message_id         TEXT,
            subject            TEXT,
            subject_clean      TEXT,
            sender             TEXT,
            sender_email       TEXT,
            sender_domain      TEXT,
            recipient          TEXT,
            date               TEXT,
            hour_of_day        INTEGER DEFAULT 0,
            day_of_week        INTEGER DEFAULT 0,
            body               TEXT,
            snippet            TEXT,
            folder             TEXT DEFAULT 'INBOX',
            is_read            INTEGER DEFAULT 0,
            sentiment          TEXT DEFAULT 'neutral',
            sentiment_score    REAL DEFAULT 0.0,
            sentiment_positive REAL DEFAULT 0.0,
            sentiment_negative REAL DEFAULT 0.0,
            word_count         INTEGER DEFAULT 0,
            FOREIGN KEY(account_id) REFERENCES email_accounts(id),
            UNIQUE(account_id, message_id)
        )""")
    # Safe migration for existing DBs
    for col, defn in [("subject_clean","TEXT"), ("sender_domain","TEXT"),
                      ("hour_of_day","INTEGER DEFAULT 0"), ("day_of_week","INTEGER DEFAULT 0")]:
        try:
            c.execute(f"ALTER TABLE emails ADD COLUMN {col} {defn}")
        except Exception:
            pass
    conn.commit()
    conn.close()

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

# ── IMAP helpers ───────────────────────────────────────────────────────
IMAP_SERVERS = {
    "gmail.com":"imap.gmail.com","yahoo.com":"imap.mail.yahoo.com",
    "outlook.com":"imap-mail.outlook.com","hotmail.com":"imap-mail.outlook.com",
    "live.com":"imap-mail.outlook.com","icloud.com":"imap.mail.me.com",
    "protonmail.com":"imap.protonmail.ch","zoho.com":"imap.zoho.com","aol.com":"imap.aol.com",
}

def detect_imap_server(addr):
    return IMAP_SERVERS.get(addr.split("@")[-1].lower(), f"imap.{addr.split('@')[-1]}")

def decode_str(s):
    if not s: return ""
    parts = decode_header(s)
    result = []
    for part, charset in parts:
        if isinstance(part, bytes):
            try: result.append(part.decode(charset or "utf-8", errors="replace"))
            except: result.append(part.decode("utf-8", errors="replace"))
        else:
            result.append(str(part))
    return " ".join(result)

def clean_subject(subject):
    s = re.sub(r'^(Re|Fwd?|AW|WG|SV|回复|答复)(\[\d+\])?:\s*', '', subject, flags=re.IGNORECASE)
    s = re.sub(r'\[.*?\]', '', s)
    return s.strip() or subject.strip()

def extract_domain(email_str):
    m = re.search(r'@([\w.\-]+)', email_str or "")
    return m.group(1).lower() if m else ""

def get_email_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and "attachment" not in str(part.get("Content-Disposition","")):
                try:
                    charset = part.get_content_charset() or "utf-8"
                    body = part.get_payload(decode=True).decode(charset, errors="replace")
                    break
                except: pass
    else:
        try:
            charset = msg.get_content_charset() or "utf-8"
            body = msg.get_payload(decode=True).decode(charset, errors="replace")
        except:
            body = str(msg.get_payload())
    return body.strip()

POSITIVE_KW = ["thank","great","excellent","good","happy","pleased","wonderful","appreciate",
               "love","amazing","perfect","best","awesome","fantastic","congratulations",
               "success","impressive","outstanding","interested","sounds good","let's connect",
               "looking forward","well done","brilliant","excited","delighted"]
NEGATIVE_KW = ["problem","issue","bad","terrible","awful","disappointed","unhappy","error",
               "fail","wrong","poor","unacceptable","hate","worst","complaint","refund",
               "cancel","urgent","critical","broken","bug","not working","unsubscribe",
               "stop sending","spam","irrelevant","not interested","confused","frustrated"]

def analyze_sentiment(text):
    if not text or len(text.strip()) < 5:
        return "neutral", 0.0, 0.0, 0.0
    clean  = re.sub(r'[^\w\s.,!?]', ' ', text[:2000])
    blob   = TextBlob(clean)
    pol    = blob.sentiment.polarity
    pos_s  = max(0, pol)
    neg_s  = max(0, -pol)
    lower  = text.lower()
    ph     = sum(1 for w in POSITIVE_KW if w in lower)
    nh     = sum(1 for w in NEGATIVE_KW if w in lower)
    pol    = max(-1, min(1, pol + (ph - nh) * 0.05))
    pos_s  = min(1, pos_s + ph * 0.05)
    neg_s  = min(1, neg_s + nh * 0.05)
    label  = "positive" if pol > 0.05 else ("negative" if pol < -0.05 else "neutral")
    return label, round(pol,4), round(pos_s,4), round(neg_s,4)

STOP = {"the","and","for","are","but","not","you","all","can","had","her","was","one","our",
        "out","day","get","has","him","his","how","man","new","now","old","see","two","way",
        "who","did","its","let","put","say","she","too","use","with","this","that","have",
        "from","they","will","been","were","your","said","each","which","their","time",
        "about","than","then","them","these","some","into","just","like","more","also",
        "over","such","even","most","made","after","back","only","come","could","would",
        "should","there","email","emails","please","thank","thanks","dear","hello",
        "regards","best","sincerely","hi","hey","hope","well"}

def extract_phrases(text, top_n=15):
    words    = re.findall(r'\b[a-z]{3,}\b', text.lower())
    filtered = [w for w in words if w not in STOP]
    bigrams  = [f"{filtered[i]} {filtered[i+1]}" for i in range(len(filtered)-1)]
    return Counter(bigrams).most_common(top_n)

def build_where(account_ids, date_from, date_to, alias=""):
    pf = f"{alias}." if alias else ""
    where, params = [], []
    if account_ids:
        where.append(f"{pf}account_id IN ({','.join('?'*len(account_ids))})")
        params.extend(account_ids)
    if date_from: where.append(f"{pf}date >= ?"); params.append(date_from+" 00:00:00")
    if date_to:   where.append(f"{pf}date <= ?"); params.append(date_to+" 23:59:59")
    return ("WHERE "+" AND ".join(where) if where else ""), params

# ── IMAP sync ──────────────────────────────────────────────────────────
def fetch_emails_from_imap(account_id, email_addr, app_password, imap_server, folder="INBOX", limit=300):
    conn = get_db(); synced = 0
    try:
        mail = imaplib.IMAP4_SSL(imap_server, 993, timeout=30)
        mail.login(email_addr, app_password)
        mail.select(folder)
        _, data = mail.search(None, "ALL")
        mids = data[0].split()[-limit:]
        for mid in reversed(mids):
            try:
                _, md = mail.fetch(mid, "(RFC822)")
                raw = md[0][1]
                msg = email.message_from_bytes(raw)
                msg_id        = msg.get("Message-ID","").strip() or hashlib.md5(raw[:200]).hexdigest()
                subject       = decode_str(msg.get("Subject","(No Subject)"))
                subject_clean = clean_subject(subject)
                sender        = decode_str(msg.get("From",""))
                recipient     = decode_str(msg.get("To",""))
                m             = re.search(r'<([^>]+)>', sender)
                sender_email  = m.group(1) if m else (sender.strip() if "@" in sender else "")
                sender_domain = extract_domain(sender_email)
                try:
                    dt = parsedate_to_datetime(msg.get("Date",""))
                    date_iso, hour_of_day, day_of_week = dt.strftime("%Y-%m-%d %H:%M:%S"), dt.hour, dt.weekday()
                except:
                    date_iso, hour_of_day, day_of_week = datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 0, 0
                body       = get_email_body(msg)
                snippet    = body[:200].replace("\n"," ").replace("\r"," ").strip()
                word_count = len(body.split()) if body else 0
                sentiment, score, pos, neg = analyze_sentiment(body or subject)
                c = conn.cursor()
                c.execute("""
                    INSERT OR IGNORE INTO emails
                    (account_id,message_id,subject,subject_clean,sender,sender_email,
                     sender_domain,recipient,date,hour_of_day,day_of_week,body,snippet,
                     folder,sentiment,sentiment_score,sentiment_positive,sentiment_negative,word_count)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,(account_id,msg_id,subject,subject_clean,sender,sender_email,sender_domain,
                     recipient,date_iso,hour_of_day,day_of_week,body,snippet,folder,sentiment,score,pos,neg,word_count))
                conn.commit()
                if c.rowcount > 0: synced += 1
            except: continue
        mail.logout()
        c = conn.cursor()
        c.execute("UPDATE email_accounts SET last_synced=?, total_emails=(SELECT COUNT(*) FROM emails WHERE account_id=?) WHERE id=?",
                  (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), account_id, account_id))
        conn.commit()
    except imaplib.IMAP4.error as e: raise Exception(f"IMAP Auth Failed: {e}")
    except Exception as e: raise Exception(f"Sync Error: {e}")
    finally: conn.close()
    return synced

# ── Routes ─────────────────────────────────────────────────────────────
@app.route("/")
def index(): return redirect("/dashboard")

@app.route("/dashboard")
def dashboard(): return render_template_string(load_template("dashboard.html"))

@app.route("/api/accounts", methods=["GET"])
def get_accounts():
    conn = get_db()
    rows = conn.execute("SELECT id,email,display_name,imap_server,added_at,last_synced,total_emails,status FROM email_accounts").fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route("/api/accounts", methods=["POST"])
def add_account():
    data         = request.json
    email_addr   = data.get("email","").strip().lower()
    app_password = data.get("app_password","").strip()
    display_name = data.get("display_name","").strip() or email_addr.split("@")[0]
    imap_server  = data.get("imap_server","").strip() or detect_imap_server(email_addr)
    if not email_addr or not app_password:
        return jsonify({"error":"Email and app password are required"}), 400
    try:
        mail = imaplib.IMAP4_SSL(imap_server, 993, timeout=15)
        mail.login(email_addr, app_password); mail.logout()
    except imaplib.IMAP4.error as e:
        return jsonify({"error":f"Authentication failed — check your app password. ({e})"}), 401
    except Exception as e:
        return jsonify({"error":f"Connection failed to {imap_server}: {e}"}), 500
    conn = get_db()
    try:
        conn.execute("INSERT INTO email_accounts (email,app_password,imap_server,display_name,added_at) VALUES (?,?,?,?,?)",
                     (email_addr,app_password,imap_server,display_name,datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        account = conn.execute("SELECT * FROM email_accounts WHERE email=?", (email_addr,)).fetchone()
        conn.close()
        return jsonify({"success":True,"account":dict(account),"message":"Account added!"})
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({"error":"Email account already added"}), 409

@app.route("/api/accounts/<int:account_id>", methods=["DELETE"])
def delete_account(account_id):
    conn = get_db()
    conn.execute("DELETE FROM emails WHERE account_id=?", (account_id,))
    conn.execute("DELETE FROM email_accounts WHERE id=?", (account_id,))
    conn.commit(); conn.close()
    return jsonify({"success":True})

@app.route("/api/sync/<int:account_id>", methods=["POST"])
def sync_account(account_id):
    conn = get_db()
    account = conn.execute("SELECT * FROM email_accounts WHERE id=?", (account_id,)).fetchone()
    conn.close()
    if not account: return jsonify({"error":"Account not found"}), 404
    def run():
        try: fetch_emails_from_imap(account_id, account["email"], account["app_password"], account["imap_server"])
        except Exception as e: print(f"Sync error [{account['email']}]: {e}")
    threading.Thread(target=run, daemon=True).start()
    return jsonify({"success":True,"message":f"Syncing {account['email']}…"})

@app.route("/api/sync/all", methods=["POST"])
def sync_all():
    conn = get_db()
    accounts = conn.execute("SELECT * FROM email_accounts WHERE status='active'").fetchall()
    conn.close()
    def run():
        for a in accounts:
            try: fetch_emails_from_imap(a["id"],a["email"],a["app_password"],a["imap_server"])
            except Exception as e: print(f"Sync error: {e}")
    threading.Thread(target=run, daemon=True).start()
    return jsonify({"success":True,"message":f"Syncing {len(accounts)} account(s)…"})

@app.route("/api/sync/status/<int:account_id>")
def sync_status(account_id):
    conn = get_db()
    row  = conn.execute("SELECT id,email,last_synced,total_emails FROM email_accounts WHERE id=?", (account_id,)).fetchone()
    conn.close()
    return jsonify(dict(row)) if row else (jsonify({"error":"Not found"}), 404)

@app.route("/api/stats")
def get_stats():
    account_ids = request.args.getlist("accounts[]") or request.args.getlist("accounts")
    date_from   = request.args.get("date_from","")
    date_to     = request.args.get("date_to","")
    conn        = get_db()
    clause, params = build_where(account_ids, date_from, date_to)
    and_or = "AND" if clause else "WHERE"

    total    = conn.execute(f"SELECT COUNT(*) FROM emails {clause}", params).fetchone()[0]
    positive = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {and_or} sentiment='positive'", params).fetchone()[0]
    negative = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {and_or} sentiment='negative'", params).fetchone()[0]
    neutral  = conn.execute(f"SELECT COUNT(*) FROM emails {clause} {and_or} sentiment='neutral'",  params).fetchone()[0]
    avg_sc   = conn.execute(f"SELECT AVG(sentiment_score) FROM emails {clause}", params).fetchone()[0] or 0

    top_senders = conn.execute(f"""
        SELECT sender_email, COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive
        FROM emails {clause} GROUP BY sender_email ORDER BY cnt DESC LIMIT 10
    """, params).fetchall()

    trend = conn.execute(f"""
        SELECT DATE(date) as day, COUNT(*) as total,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN sentiment='negative' THEN 1 ELSE 0 END) as negative,
               SUM(CASE WHEN sentiment='neutral'  THEN 1 ELSE 0 END) as neutral
        FROM emails {clause} GROUP BY DATE(date) ORDER BY day ASC LIMIT 30
    """, params).fetchall()

    by_account = conn.execute(f"""
        SELECT ea.email, ea.display_name, COUNT(*) as total,
               SUM(CASE WHEN e.sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN e.sentiment='negative' THEN 1 ELSE 0 END) as negative,
               SUM(CASE WHEN e.sentiment='neutral'  THEN 1 ELSE 0 END) as neutral,
               AVG(e.sentiment_score) as avg_score
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id
        {clause} GROUP BY e.account_id
    """, params).fetchall()

    # Top subject lines
    top_subjects = conn.execute(f"""
        SELECT COALESCE(NULLIF(subject_clean,''), subject) as subj,
               COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN sentiment='negative' THEN 1 ELSE 0 END) as negative,
               ROUND(CAST(SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) AS REAL)/COUNT(*)*100,1) as reply_rate,
               AVG(sentiment_score) as avg_score
        FROM emails {clause} GROUP BY subj ORDER BY cnt DESC LIMIT 12
    """, params).fetchall()

    # Top domains
    dom_clause = clause + (f" AND " if clause else " WHERE ") + "sender_domain != '' AND sender_domain IS NOT NULL"
    top_domains = conn.execute(f"""
        SELECT sender_domain as domain, COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as positive,
               SUM(CASE WHEN sentiment='negative' THEN 1 ELSE 0 END) as negative,
               ROUND(AVG(sentiment_score),3) as avg_score
        FROM emails {dom_clause}
        GROUP BY domain ORDER BY cnt DESC LIMIT 12
    """, params).fetchall()

    conn.close()
    return jsonify({
        "total":total,"positive":positive,"negative":negative,
        "neutral":neutral,"avg_score":round(avg_sc,3),
        "top_senders": [dict(s) for s in top_senders],
        "trend":        [dict(t) for t in trend],
        "by_account":   [dict(a) for a in by_account],
        "top_subjects": [dict(s) for s in top_subjects],
        "top_domains":  [dict(d) for d in top_domains],
    })

@app.route("/api/intelligence")
def get_intelligence():
    account_ids = request.args.getlist("accounts[]") or request.args.getlist("accounts")
    date_from   = request.args.get("date_from","")
    date_to     = request.args.get("date_to","")
    conn        = get_db()
    clause, params = build_where(account_ids, date_from, date_to)

    best_hour = conn.execute(f"""
        SELECT hour_of_day, COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as pos
        FROM emails {clause} GROUP BY hour_of_day ORDER BY pos DESC, cnt DESC LIMIT 1
    """, params).fetchone()

    best_day = conn.execute(f"""
        SELECT day_of_week, COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as pos
        FROM emails {clause} GROUP BY day_of_week ORDER BY pos DESC, cnt DESC LIMIT 1
    """, params).fetchone()

    hour_dist = conn.execute(f"""
        SELECT hour_of_day as hour, COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as pos
        FROM emails {clause} GROUP BY hour_of_day ORDER BY hour_of_day
    """, params).fetchall()

    day_dist = conn.execute(f"""
        SELECT day_of_week as dow, COUNT(*) as cnt,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as pos
        FROM emails {clause} GROUP BY day_of_week ORDER BY day_of_week
    """, params).fetchall()

    # Reply rate by subject
    subject_reply = conn.execute(f"""
        SELECT COALESCE(NULLIF(subject_clean,''),subject) as subj,
               COUNT(*) as total,
               SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) as pos,
               ROUND(CAST(SUM(CASE WHEN sentiment='positive' THEN 1 ELSE 0 END) AS REAL)/COUNT(*)*100,1) as reply_rate
        FROM emails {clause}
        GROUP BY subj HAVING total >= 1
        ORDER BY reply_rate DESC, pos DESC LIMIT 8
    """, params).fetchall()

    # Positive phrases
    pos_clause = clause + (" AND " if clause else " WHERE ") + "sentiment='positive'"
    pos_bodies = conn.execute(f"SELECT body FROM emails {pos_clause} LIMIT 300", params).fetchall()
    pos_text   = " ".join(r["body"] or "" for r in pos_bodies)
    phrases    = extract_phrases(pos_text, 15)

    # Negative phrases
    neg_clause = clause + (" AND " if clause else " WHERE ") + "sentiment='negative'"
    neg_bodies = conn.execute(f"SELECT body FROM emails {neg_clause} LIMIT 200", params).fetchall()
    neg_text   = " ".join(r["body"] or "" for r in neg_bodies)
    neg_phrases = extract_phrases(neg_text, 8)

    DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    best_hour_str  = f"{best_hour['hour_of_day']:02d}:00" if best_hour else "N/A"
    best_day_name  = DAYS[best_day["day_of_week"]] if best_day else "N/A"

    # Build suggestions
    suggestions = []
    if best_hour and best_hour["pos"] > 0:
        suggestions.append(f"Replies peak at {best_hour_str} UTC — schedule sends 1–2 hours earlier.")
    if best_day and best_day["pos"] > 0:
        suggestions.append(f"{best_day_name} drives the most replies — prioritize that day for outreach.")
    if subject_reply:
        s = subject_reply[0]
        subj_short = (s["subj"] or "")[:55]
        suggestions.append(f'Subject "{subj_short}" performs {s["reply_rate"]}% reply rate — replicate its hook.')
    if phrases:
        suggestions.append(f'Positive replies often include "{phrases[0][0]}" — listen for it in incoming threads.')
    if neg_phrases:
        suggestions.append(f'Negative replies often contain "{neg_phrases[0][0]}" — flag these for immediate follow-up.')

    conn.close()
    return jsonify({
        "best_hour": best_hour_str,
        "best_hour_replies": best_hour["pos"] if best_hour else 0,
        "best_day": best_day_name,
        "best_day_replies": best_day["pos"] if best_day else 0,
        "hour_dist": [dict(h) for h in hour_dist],
        "day_dist":  [{"day":DAYS[d["dow"]],"dow":d["dow"],"cnt":d["cnt"],"pos":d["pos"]} for d in day_dist],
        "subject_reply_rate": [dict(s) for s in subject_reply],
        "positive_phrases": [{"phrase":p,"count":c} for p,c in phrases],
        "negative_phrases": [{"phrase":p,"count":c} for p,c in neg_phrases],
        "suggestions": suggestions,
    })

@app.route("/api/emails")
def get_emails():
    account_ids = request.args.getlist("accounts[]") or request.args.getlist("accounts")
    date_from   = request.args.get("date_from","")
    date_to     = request.args.get("date_to","")
    sentiment   = request.args.get("sentiment","")
    search      = request.args.get("search","")
    folder      = request.args.get("folder","INBOX")
    page        = int(request.args.get("page",1))
    per_page    = int(request.args.get("per_page",30))
    conn        = get_db()
    where, params = [], []
    if account_ids: where.append(f"e.account_id IN ({','.join('?'*len(account_ids))})"); params.extend(account_ids)
    if date_from:   where.append("e.date >= ?"); params.append(date_from+" 00:00:00")
    if date_to:     where.append("e.date <= ?"); params.append(date_to+" 23:59:59")
    if sentiment:   where.append("e.sentiment = ?"); params.append(sentiment)
    if search:      where.append("(e.subject LIKE ? OR e.sender LIKE ? OR e.snippet LIKE ?)"); params.extend([f"%{search}%"]*3)
    if folder:      where.append("e.folder = ?"); params.append(folder)
    clause = "WHERE "+" AND ".join(where) if where else ""
    total  = conn.execute(f"SELECT COUNT(*) FROM emails e {clause}", params).fetchone()[0]
    rows   = conn.execute(f"""
        SELECT e.*, ea.email as account_email, ea.display_name as account_name
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id
        {clause} ORDER BY e.date DESC LIMIT ? OFFSET ?
    """, params+[per_page,(page-1)*per_page]).fetchall()
    conn.close()
    return jsonify({"total":total,"page":page,"per_page":per_page,
                    "pages":(total+per_page-1)//per_page,"emails":[dict(e) for e in rows]})

@app.route("/api/emails/<int:email_id>")
def get_email(email_id):
    conn = get_db()
    e    = conn.execute("""
        SELECT e.*, ea.email as account_email, ea.display_name as account_name
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id WHERE e.id=?
    """, (email_id,)).fetchone()
    conn.close()
    return jsonify(dict(e)) if e else (jsonify({"error":"Not found"}), 404)

@app.route("/api/export/csv")
def export_csv():
    account_ids = request.args.getlist("accounts[]") or request.args.getlist("accounts")
    date_from   = request.args.get("date_from","")
    date_to     = request.args.get("date_to","")
    sentiment   = request.args.get("sentiment","")
    conn        = get_db()
    where, params = [], []
    if account_ids: where.append(f"e.account_id IN ({','.join('?'*len(account_ids))})"); params.extend(account_ids)
    if date_from:   where.append("e.date >= ?"); params.append(date_from+" 00:00:00")
    if date_to:     where.append("e.date <= ?"); params.append(date_to+" 23:59:59")
    if sentiment:   where.append("e.sentiment = ?"); params.append(sentiment)
    clause = "WHERE "+" AND ".join(where) if where else ""
    rows   = conn.execute(f"""
        SELECT e.date, ea.email as account, e.subject, e.sender_email, e.sender_domain,
               e.sentiment, e.sentiment_score, e.sentiment_positive, e.sentiment_negative,
               e.word_count, e.snippet
        FROM emails e JOIN email_accounts ea ON e.account_id=ea.id
        {clause} ORDER BY e.date DESC LIMIT 5000
    """, params).fetchall()
    conn.close()
    out = io.StringIO()
    w   = csv.writer(out)
    w.writerow(["Date","Account","Subject","Sender Email","Sender Domain","Sentiment",
                "Score","Positive","Negative","Word Count","Snippet"])
    for r in rows: w.writerow(list(r))
    out.seek(0)
    fname = f"inbox_iq_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    return Response(out.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": f"attachment; filename={fname}"})

if __name__ == "__main__":
    init_db()
    print("=" * 60)
    print("  📧 INBOX.IQ — Email Analytics Hub")
    print("  Open: http://localhost:5001/dashboard")
    print("=" * 60)
    app.run(debug=True, port=5001, threaded=True)
