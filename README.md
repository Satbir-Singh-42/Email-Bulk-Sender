# MailBlast Pro — Mass Email Sender

A modern, dark-themed desktop app for bulk email sending.

## Setup

1. Install Python 3.8+
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Run:
   ```
   python mass_email_sender.py
   ```

## Gmail App Password Setup (REQUIRED for Gmail)

1. Go to myaccount.google.com
2. Security → 2-Step Verification → Enable it
3. Security → App Passwords → Create one for "Mail"
4. Use that 16-character password in the app (not your real Gmail password)

## Excel File Format

Your Excel file should have these columns (names are configurable):
| Email         | Name     | Path (optional)       | Event       | Date       | Role      |
|---------------|----------|-----------------------|-------------|------------|-----------|
| user@mail.com | John Doe | C:\certs\john.pdf     | Workshop 1  | 2024-01-15 | Speaker   |

## Navigation

- **SMTP Setup** → Enter server, port, email, password, test connection
- **Recipients** → Load Excel file, select sheets, preview data
- **Compose** → Write subject, CC/BCC, email body with {variables}
- **Attachments** → Add per-recipient or same-for-all attachments
- **Send** → Configure delay, simulate/send, track progress
- **Logs** → Full activity log, export to file

## Template Variables

Use `{ColumnName}` in your email body to personalize:
- `{Name}` → recipient name
- `{Event}` → event name
- `{Date}` → date
- `{Role}` → their role
- Any column from your Excel file works!
