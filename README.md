# AI Email Outreach System

A Flask web application for running B2B email campaigns. It manages reusable templates, reads leads from an Excel file, sends emails with built-in rate limiting, and schedules multi-step follow-ups.

## Features

- Web interface for managing outreach campaigns across multiple markets
- Template management with saved, reusable email templates
- Multi-step automated follow-ups
- Daily send limit and configurable delay between emails
- Lead import from Excel (columns: business_name, email, city, website, category)

## Tech stack

Python, Flask, Gmail SMTP, HTML/Jinja2

## Setup

1. Enable 2-Step Verification on your Gmail account and create an App Password (Google Account > Security > App passwords).
2. Copy `.env.example` to `.env` and fill in your own values:

        SENDER_EMAIL=you@example.com
        SENDER_APP_PASSWORD=your-gmail-app-password
        SENDER_NAME=Your Name
        DAILY_LIMIT=50
        DELAY_SECONDS=30

3. Install dependencies:

        pip install -r requirements.txt

4. Place your leads file next to the app and start it:

        python app.py

## Notes

- `.env` and lead files are gitignored. Never commit credentials or real contact data.
- Use responsibly and comply with applicable anti-spam and data-protection laws (CAN-SPAM, GDPR, PECR). Always include an opt-out option.