# 📧 Email Analytics Hub

A powerful multi-account email analytics dashboard with sentiment analysis.

## Features
- ✅ Multi-email account support (Gmail, Yahoo, Outlook, iCloud, Zoho, AOL, etc.)
- ✅ App Password authentication (secure, no OAuth needed)
- ✅ Per-account sentiment analysis (Positive / Negative / Neutral)
- ✅ Date-range filtering
- ✅ Inbox with search, filter by sentiment, pagination
- ✅ Email detail panel with sentiment bars
- ✅ Per-account breakdown charts
- ✅ Top senders analytics
- ✅ Sync button (per account + sync all)
- ✅ Background sync (non-blocking)

---

## Quick Start

### 1. Install Dependencies
```bash
pip install flask textblob
python -m textblob.download_corpora
```

### 2. Run the App
```bash
python app.py
```

### 3. Open Browser
```
http://localhost:5001/dashboard
```

---

## How to Get App Passwords

### Gmail
1. Enable 2-Step Verification: myaccount.google.com/security
2. Go to: myaccount.google.com/apppasswords
3. Select "Mail" → Generate
4. Use the 16-character password (spaces are OK)

### Yahoo Mail
1. Go to Account Security settings
2. Enable 2-Step Verification
3. Generate an app password for "Mail"

### Outlook/Hotmail
1. Enable 2-Factor Authentication at account.microsoft.com
2. Go to Security → Advanced security options
3. Create an app password

### iCloud (Apple Mail)
1. Go to appleid.apple.com
2. Sign In → Security → App-Specific Passwords
3. Generate one for "Email Analytics"

---

## IMAP Servers (Auto-Detected)
| Provider | IMAP Server |
|----------|------------|
| Gmail | imap.gmail.com |
| Yahoo | imap.mail.yahoo.com |
| Outlook | imap-mail.outlook.com |
| iCloud | imap.mail.me.com |
| Zoho | imap.zoho.com |
| AOL | imap.aol.com |

---

## Project Structure
```
email_analytics/
├── app.py              # Flask backend
├── requirements.txt    # Dependencies
├── README.md
├── email_analytics.db  # SQLite DB (auto-created)
└── templates/
    └── dashboard.html  # Full dashboard UI
```

---

## Security Notes
- App passwords are stored locally in SQLite (email_analytics.db)
- Never share your app password
- The app runs locally — your emails never leave your machine
- Delete accounts anytime from the sidebar
