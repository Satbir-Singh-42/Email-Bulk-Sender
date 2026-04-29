# 📧 Email Bulk Sender

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)

A powerful, modern desktop application designed for high-performance bulk email delivery. Built with a focus on simplicity, personalization, and reliability.

---

## ✨ Key Features

- **🚀 High Performance:** Send hundreds of emails with configurable delays to avoid spam filters.
- **🎨 Dynamic Personalization:** Use `{Variable}` tags in your subjects and bodies to automatically pull data from your Excel/CSV columns.
- **📁 Smart Attachments:** Send common attachments to everyone or unique files per recipient based on file paths in your spreadsheet.
- **🔍 Live Preview:** Verify your recipient data and sheet mappings before hitting send.
- **🛠️ Provider Presets:** Instant configuration for Gmail, Outlook, Yahoo, and Office365.
- **📊 Detailed Logs:** Track every success and failure with real-time logging and export capabilities.

---

## 🛠️ Installation & Setup

### 1. Prerequisites
- [Python 3.8 or higher](https://www.python.org/downloads/) installed on your system.

### 2. Install Dependencies
Clone the repository and install the required Python packages:
```bash
git clone https://github.com/Satbir-Singh-42/Email-Bulk-Sender.git
cd Email-Bulk-Sender
pip install -r requirements.txt
```

### 3. Launch the Application
```bash
python mass_email_sender.py
```

---

## 🔒 Security: Gmail App Password (Required)

If you are using Gmail, you **must** use an App Password instead of your primary password:

1. Go to [Google My Account](https://myaccount.google.com/).
2. Navigate to **Security** → **2-Step Verification** (Ensure it's ON).
3. Scroll to **App Passwords** at the bottom.
4. Select **Mail** and your device, then click **Create**.
5. Copy the 16-character code and use it in the application's password field.

---

## 📂 Spreadsheet Configuration

Your Excel or CSV file should contain a header row. You can map any column name to the application's fields.

| Email | Name | Attachment_Path (Optional) | Event |
| :--- | :--- | :--- | :--- |
| user@example.com | John Doe | `C:\certs\cert1.pdf` | Tech Summit |
| jane@example.com | Jane Smith | `C:\certs\cert2.pdf` | AI Workshop |

### Using Variables
In your email body, use curly braces to reference any column:
> *"Hi {Name}, thank you for attending the {Event}!"*

---

## ⚖️ License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

---

## 🤝 Contributing

Contributions are welcome! Feel free to open an issue or submit a pull request to improve the project.
