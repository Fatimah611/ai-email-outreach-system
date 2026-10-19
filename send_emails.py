"""
Your Company Email Marketing Sender
--------------------------------
Reads leads from an Excel file, personalizes a template, and sends
emails via Gmail SMTP (App Password). No external AI API required.

USAGE:
    python send_emails.py --dry-run          # preview only, sends nothing
    python send_emails.py --test you@x.com   # send ONE test email to yourself
    python send_emails.py --send             # real batch send to all leads

Required Excel columns (case-insensitive, order doesn't matter):
    business_name, email, city   (website, category, phone are optional)

The script writes back 3 columns after sending:
    status, sent_at, error
"""

import argparse
import os
import smtplib
import ssl
import sys
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_APP_PASSWORD = os.getenv("SENDER_APP_PASSWORD")
SENDER_NAME = os.getenv("SENDER_NAME", "Your Company")
DAILY_LIMIT = int(os.getenv("DAILY_LIMIT", "80"))
DELAY_SECONDS = int(os.getenv("DELAY_SECONDS", "8"))

DEMO_LINK = "https://example.com/demo"

LEADS_FILE = "leads.xlsx"
TEMPLATE_FILE = "template.txt"

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587


def load_template():
    with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
        raw = f.read()
    subject_line = ""
    body_lines = []
    for line in raw.splitlines():
        if line.strip().upper().startswith("SUBJECT:"):
            subject_line = line.split(":", 1)[1].strip()
        else:
            body_lines.append(line)
    body = "\n".join(body_lines).strip()
    return subject_line, body


def personalize(text, row):
    return text.format(
        business_name=row.get("business_name", ""),
        city=row.get("city", ""),
        website=row.get("website", ""),
        category=row.get("category", ""),
        demo_link=DEMO_LINK,
        sender_email=SENDER_EMAIL,
    )


def build_message(to_email, subject, body):
    msg = MIMEMultipart()
    msg["From"] = f"{SENDER_NAME} <{SENDER_EMAIL}>"
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    return msg


def send_one(server, to_email, subject, body):
    msg = build_message(to_email, subject, body)
    server.sendmail(SENDER_EMAIL, to_email, msg.as_string())


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Preview only, sends nothing")
    parser.add_argument("--test", type=str, help="Send one test email to this address")
    parser.add_argument("--send", action="store_true", help="Real batch send to all leads")
    args = parser.parse_args()

    if not SENDER_EMAIL or not SENDER_APP_PASSWORD:
        print("ERROR: SENDER_EMAIL / SENDER_APP_PASSWORD missing in .env file.")
        sys.exit(1)

    subject_template, body_template = load_template()

    if args.test:
        fake_row = {
            "business_name": "Sample Dental Clinic",
            "city": "Sydney",
            "website": "sampledental.com",
            "category": "Dental Clinic",
        }
        subject = personalize(subject_template, fake_row)
        body = personalize(body_template, fake_row)
        print(f"Sending TEST email to {args.test} ...")
        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.starttls(context=context)
            server.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
            send_one(server, args.test, subject, body)
        print("Test email sent successfully.")
        return

    if not os.path.exists(LEADS_FILE):
        print(f"ERROR: {LEADS_FILE} not found. Put your leads Excel file next to this script.")
        sys.exit(1)

    df = pd.read_excel(LEADS_FILE)
    df.columns = [c.strip().lower() for c in df.columns]

    required = {"business_name", "email"}
    if not required.issubset(set(df.columns)):
        print(f"ERROR: Excel must contain at least these columns: {required}")
        sys.exit(1)

    if "status" not in df.columns:
        df["status"] = ""
    if "sent_at" not in df.columns:
        df["sent_at"] = ""
    if "error" not in df.columns:
        df["error"] = ""

    if args.dry_run:
        print(f"\n--- DRY RUN: {len(df)} leads found ---\n")
        for i, row in df.iterrows():
            if str(row.get("status", "")).strip().lower() == "sent":
                continue
            subject = personalize(subject_template, row)
            print(f"[{i+1}] To: {row['email']}  |  Subject: {subject}")
        print("\nNo emails were actually sent. Run with --send to actually send.")
        return

    if args.send:
        context = ssl.create_default_context()
        sent_count = 0
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.starttls(context=context)
            server.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
            for i, row in df.iterrows():
                if str(row.get("status", "")).strip().lower() == "sent":
                    continue
                if sent_count >= DAILY_LIMIT:
                    print(f"Daily limit of {DAILY_LIMIT} reached. Stopping for today.")
                    break
                to_email = str(row.get("email", "")).strip()
                if not to_email or "@" not in to_email:
                    df.at[i, "status"] = "skipped"
                    df.at[i, "error"] = "invalid email"
                    continue
                try:
                    subject = personalize(subject_template, row)
                    body = personalize(body_template, row)
                    send_one(server, to_email, subject, body)
                    df.at[i, "status"] = "sent"
                    df.at[i, "sent_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    df.at[i, "error"] = ""
                    sent_count += 1
                    print(f"[{i+1}] Sent to {to_email}")
                except Exception as e:
                    df.at[i, "status"] = "failed"
                    df.at[i, "error"] = str(e)
                    print(f"[{i+1}] FAILED for {to_email}: {e}")

                df.to_excel(LEADS_FILE, index=False)
                time.sleep(DELAY_SECONDS)

        df.to_excel(LEADS_FILE, index=False)
        print(f"\nDone. {sent_count} emails sent this run. Progress saved to {LEADS_FILE}.")
        return

    print("Specify one of: --dry-run, --test you@email.com, or --send")


if __name__ == "__main__":
    main()
