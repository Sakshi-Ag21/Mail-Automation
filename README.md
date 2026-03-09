# Mail Automation Dashboard (Streamlit)

Bulk email automation dashboard built with Python + Streamlit.

## Features

- Upload CSV/Excel with `Name` and `Email` columns
- Write email subject and body (supports `{name}` personalization)
- Upload attachments
- Send a test email to a specified address
- Send bulk emails one by one using Gmail SMTP
- Add a delay between sends
- Live status + logs on the dashboard
- Credentials loaded from environment variables (optionally from a local `.env`)

## Prerequisites (Gmail)

Gmail SMTP requires either:

- An **App Password** (recommended) with 2FA enabled on the account, or
- A Google Workspace SMTP relay configuration (advanced)

This app expects an **App Password**.

## Setup

1) Create and activate a virtual environment (optional but recommended).

2) Install dependencies:

```bash
pip3 install -r requirements.txt
```

3) Set environment variables:

- `SMTP_EMAIL`: your Gmail address (sender)
- `SMTP_APP_PASSWORD`: Gmail App Password (not your normal password)
- `SMTP_HOST` (optional): default `smtp.gmail.com`
- `SMTP_PORT` (optional): default `587`
- `SMTP_SENDER_NAME` (optional): display name in the From header

### Option A: Export in your shell

```bash
export SMTP_EMAIL="you@gmail.com"
export SMTP_APP_PASSWORD="xxxx xxxx xxxx xxxx"
```

### Option B: Local `.env` (for development only)

Create a `.env` file in the project root:

```env
SMTP_EMAIL=you@gmail.com
SMTP_APP_PASSWORD=xxxx xxxx xxxx xxxx
SMTP_SENDER_NAME=Your Name
```

## Run

```bash
streamlit run app.py
```

## CSV / Excel format

Your file must include columns:

- `Name`
- `Email`

Extra columns are allowed and ignored.

## Personalization

Use `{name}` in subject/body to personalize:

- Subject: `Hi {name} — quick question`
- Body: `Hello {name},\n\n...`

