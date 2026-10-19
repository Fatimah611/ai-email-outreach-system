"""
Your Company Email Marketing - Local Web UI
-----------------------------------------
Run this, then open http://localhost:5001 in your browser.
No cloud hosting, no external server - everything runs on your own computer.

UPDATED: Added automatic follow-up system.
- Leads that were sent the initial email and haven't replied get a follow-up
  automatically after FOLLOWUP_GAP_DAYS days (max MAX_FOLLOWUP_STAGE follow-ups).
- Visit http://localhost:5001/followups to send due follow-ups.
- If a lead replies, manually change their "status" in leads.xlsx to "replied"
  so they stop receiving follow-ups.
"""

import json
import os
import smtplib
import ssl
import threading
import time
from datetime import datetime

import pandas as pd
from flask import Flask, jsonify, render_template, request
from werkzeug.utils import secure_filename

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LEADS_FILE = os.path.join(BASE_DIR, "leads.xlsx")
TEMPLATE_FILE = os.path.join(BASE_DIR, "template.txt")
ENV_FILE = os.path.join(BASE_DIR, ".env")

# ---- Follow-up settings ----
FOLLOWUP_TEMPLATE_FILES = {
    1: os.path.join(BASE_DIR, "template_followup1.txt"),
    2: os.path.join(BASE_DIR, "template_followup2.txt"),
}
FOLLOWUP_GAP_DAYS = 4        # days to wait before sending next follow-up
MAX_FOLLOWUP_STAGE = 2       # total number of follow-ups per lead (after initial email)

DEFAULT_FOLLOWUP_TEMPLATES = {
    1: {
        "subject": "Following up - Your Company x {business_name}",
        "body": (
            "Hi,\n\n"
            "Following up on my earlier note. We've been helping dental clinics automate "
            "patient inquiries and after-hours calls using AI voice agents - reducing missed-call "
            "losses and freeing up front-desk time.\n\n"
            "Happy to share more details or a short demo if it's useful for {business_name}: {demo_link}\n\n"
            "Best regards,\n"
            "Your Name\n"
            "Your Title, Your Company"
        ),
    },
    2: {
        "subject": "One last note - Your Company",
        "body": (
            "Hi,\n\n"
            "Just checking in one last time, in case my earlier emails reached you at a busy moment.\n\n"
            "If this isn't a priority for {business_name} right now, no worries at all - happy to "
            "reconnect whenever the timing works better.\n\n"
            "Best regards,\n"
            "Your Name\n"
            "Your Title, Your Company"
        ),
    },
}

DEMO_LINK = "https://example.com/demo"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587

app = Flask(__name__)

# In-memory progress tracker for the currently running initial-send job
send_state = {
    "running": False,
    "done": False,
    "sent_count": 0,
    "total": 0,
    "log": [],
}
send_lock = threading.Lock()

# Separate in-memory progress tracker for follow-up jobs (kept independent
# so a follow-up run never collides with an initial-send run)
followup_state = {
    "running": False,
    "done": False,
    "sent_count": 0,
    "total": 0,
    "log": [],
}
followup_lock = threading.Lock()


# ---------- Helpers ----------

def read_env():
    data = {"SENDER_EMAIL": "", "SENDER_APP_PASSWORD": "", "SENDER_NAME": "Your Company",
             "DAILY_LIMIT": "80", "DELAY_SECONDS": "8"}
    if os.path.exists(ENV_FILE):
        with open(ENV_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                data[k.strip()] = v.strip()
    return data


def write_env(data):
    with open(ENV_FILE, "w", encoding="utf-8") as f:
        for k, v in data.items():
            f.write(f"{k}={v}\n")


def read_template(path=TEMPLATE_FILE, defaults=None):
    """Generic template reader. Works for the initial template and follow-up templates."""
    if not os.path.exists(path):
        if defaults:
            return defaults
        return {"subject": "", "body": ""}
    with open(path, "r", encoding="utf-8") as f:
        raw = f.read()
    subject = ""
    body_lines = []
    for line in raw.splitlines():
        if line.strip().upper().startswith("SUBJECT:"):
            subject = line.split(":", 1)[1].strip()
        else:
            body_lines.append(line)
    return {"subject": subject, "body": "\n".join(body_lines).strip()}


def write_template(subject, body, path=TEMPLATE_FILE):
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"SUBJECT: {subject}\n\n{body}\n")


def read_leads_df():
    if not os.path.exists(LEADS_FILE):
        return pd.DataFrame(columns=["business_name", "email", "city", "website", "status",
                                       "sent_at", "error", "follow_up_stage"])
    df = pd.read_excel(LEADS_FILE)
    df.columns = [c.strip().lower() for c in df.columns]
    for col in ["status", "sent_at", "error"]:
        if col not in df.columns:
            df[col] = ""
    if "follow_up_stage" not in df.columns:
        df["follow_up_stage"] = 0
    df["follow_up_stage"] = pd.to_numeric(df["follow_up_stage"], errors="coerce").fillna(0).astype(int)
    df = df.fillna("")
    return df


def personalize(text, row, env):
    return text.format(
        business_name=row.get("business_name", ""),
        city=row.get("city", ""),
        website=row.get("website", ""),
        category=row.get("category", ""),
        demo_link=DEMO_LINK,
        sender_email=env.get("SENDER_EMAIL", ""),
    )


def get_due_followups(df):
    """Leads eligible for their next follow-up today."""
    today = datetime.now().date()

    def is_due(row):
        if str(row.get("status", "")).strip().lower() != "sent":
            return False
        if int(row.get("follow_up_stage", 0)) >= MAX_FOLLOWUP_STAGE:
            return False
        sent_at_raw = row.get("sent_at", "")
        if not sent_at_raw:
            return False
        try:
            last_sent = pd.to_datetime(sent_at_raw).date()
        except Exception:
            return False
        return (today - last_sent).days >= FOLLOWUP_GAP_DAYS

    return df[df.apply(is_due, axis=1)]


# ---------- Routes: pages ----------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/followups")
def followups_page():
    return render_template("followups.html")


# ---------- Routes: settings ----------

@app.route("/api/settings", methods=["GET"])
def get_settings():
    env = read_env()
    masked = "•" * 12 if env.get("SENDER_APP_PASSWORD") else ""
    return jsonify({
        "sender_email": env.get("SENDER_EMAIL", ""),
        "sender_name": env.get("SENDER_NAME", "Your Company"),
        "daily_limit": env.get("DAILY_LIMIT", "80"),
        "delay_seconds": env.get("DELAY_SECONDS", "8"),
        "has_password": bool(env.get("SENDER_APP_PASSWORD")),
        "password_masked": masked,
    })


@app.route("/api/settings", methods=["POST"])
def save_settings():
    payload = request.json or {}
    env = read_env()
    env["SENDER_EMAIL"] = payload.get("sender_email", env.get("SENDER_EMAIL", ""))
    env["SENDER_NAME"] = payload.get("sender_name", env.get("SENDER_NAME", "Your Company"))
    env["DAILY_LIMIT"] = str(payload.get("daily_limit", env.get("DAILY_LIMIT", "80")))
    env["DELAY_SECONDS"] = str(payload.get("delay_seconds", env.get("DELAY_SECONDS", "8")))
    new_password = payload.get("sender_app_password", "").strip()
    if new_password:
        env["SENDER_APP_PASSWORD"] = new_password.replace(" ", "")
    write_env(env)
    return jsonify({"ok": True})


# ---------- Routes: template (initial email) ----------

@app.route("/api/template", methods=["GET"])
def get_template():
    return jsonify(read_template())


@app.route("/api/template", methods=["POST"])
def save_template_route():
    payload = request.json or {}
    write_template(payload.get("subject", ""), payload.get("body", ""))
    return jsonify({"ok": True})


# ---------- Routes: follow-up templates ----------

@app.route("/api/followup-template/<int:stage>", methods=["GET"])
def get_followup_template(stage):
    path = FOLLOWUP_TEMPLATE_FILES.get(stage)
    if not path:
        return jsonify({"error": "Invalid stage"}), 400
    return jsonify(read_template(path, defaults=DEFAULT_FOLLOWUP_TEMPLATES.get(stage)))


@app.route("/api/followup-template/<int:stage>", methods=["POST"])
def save_followup_template(stage):
    path = FOLLOWUP_TEMPLATE_FILES.get(stage)
    if not path:
        return jsonify({"error": "Invalid stage"}), 400
    payload = request.json or {}
    write_template(payload.get("subject", ""), payload.get("body", ""), path=path)
    return jsonify({"ok": True})


# ---------- Routes: leads ----------

@app.route("/api/leads", methods=["GET"])
def get_leads():
    df = read_leads_df()
    return jsonify({
        "leads": df.to_dict(orient="records"),
        "total": len(df),
        "sent": int((df["status"].astype(str).str.lower() == "sent").sum()),
    })


@app.route("/api/leads/upload", methods=["POST"])
def upload_leads():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file uploaded"}), 400
    f = request.files["file"]
    filename = secure_filename(f.filename or "")
    if not filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"ok": False, "error": "Please upload an .xlsx or .xls file"}), 400
    f.save(LEADS_FILE)
    try:
        df = read_leads_df()
    except Exception as e:
        return jsonify({"ok": False, "error": f"Could not read file: {e}"}), 400
    required = {"business_name", "email"}
    if not required.issubset(set(df.columns)):
        return jsonify({"ok": False, "error": "File must have 'business_name' and 'email' columns"}), 400
    return jsonify({"ok": True, "total": len(df)})


# ---------- Routes: dry run (initial email) ----------

@app.route("/api/dry-run", methods=["POST"])
def dry_run():
    df = read_leads_df()
    tmpl = read_template()
    env = read_env()
    preview = []
    for _, row in df.iterrows():
        if str(row.get("status", "")).strip().lower() == "sent":
            continue
        subject = personalize(tmpl["subject"], row, env)
        to_email = str(row.get("email", "")).strip()
        preview.append({
            "business_name": row.get("business_name", ""),
            "email": to_email if to_email and "@" in to_email else "(missing email)",
            "subject": subject,
            "sendable": bool(to_email and "@" in to_email),
        })
    return jsonify({"preview": preview})


# ---------- Routes: follow-up dry run / due count ----------

@app.route("/api/followups/due", methods=["GET"])
def followups_due():
    df = read_leads_df()
    due_df = get_due_followups(df)
    preview = []
    for _, row in due_df.iterrows():
        next_stage = int(row.get("follow_up_stage", 0)) + 1
        preview.append({
            "business_name": row.get("business_name", ""),
            "email": row.get("email", ""),
            "next_stage": next_stage,
            "last_sent": row.get("sent_at", ""),
        })
    return jsonify({"due_count": len(due_df), "preview": preview})


# ---------- Routes: send (initial email) ----------

def run_send_job():
    env = read_env()
    tmpl = read_template()
    df = read_leads_df()

    daily_limit = int(env.get("DAILY_LIMIT", "80") or 80)
    delay_seconds = int(env.get("DELAY_SECONDS", "8") or 8)
    sender_email = env.get("SENDER_EMAIL", "")
    sender_password = env.get("SENDER_APP_PASSWORD", "")
    sender_name = env.get("SENDER_NAME", "Your Company")

    pending_idx = [i for i, r in df.iterrows() if str(r.get("status", "")).strip().lower() != "sent"]

    with send_lock:
        send_state["running"] = True
        send_state["done"] = False
        send_state["sent_count"] = 0
        send_state["total"] = len(pending_idx)
        send_state["log"] = []

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.starttls(context=context)
            server.login(sender_email, sender_password)

            count = 0
            for i in pending_idx:
                row = df.loc[i]
                if count >= daily_limit:
                    with send_lock:
                        send_state["log"].append(f"Daily limit of {daily_limit} reached. Stopping.")
                    break

                to_email = str(row.get("email", "")).strip()
                if not to_email or "@" not in to_email:
                    df.at[i, "status"] = "skipped"
                    df.at[i, "error"] = "invalid email"
                    with send_lock:
                        send_state["log"].append(f"Skipped {row.get('business_name','')} (no email)")
                    continue

                try:
                    subject = personalize(tmpl["subject"], row, env)
                    body = personalize(tmpl["body"], row, env)
                    from email.mime.multipart import MIMEMultipart
                    from email.mime.text import MIMEText
                    msg = MIMEMultipart()
                    msg["From"] = f"{sender_name} <{sender_email}>"
                    msg["To"] = to_email
                    msg["Subject"] = subject
                    msg.attach(MIMEText(body, "plain"))
                    server.sendmail(sender_email, to_email, msg.as_string())

                    df.at[i, "status"] = "sent"
                    df.at[i, "sent_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    df.at[i, "error"] = ""
                    df.at[i, "follow_up_stage"] = 0
                    count += 1
                    with send_lock:
                        send_state["sent_count"] = count
                        send_state["log"].append(f"Sent to {row.get('business_name','')} ({to_email})")
                except Exception as e:
                    df.at[i, "status"] = "failed"
                    df.at[i, "error"] = str(e)
                    with send_lock:
                        send_state["log"].append(f"FAILED {row.get('business_name','')}: {e}")

                df.to_excel(LEADS_FILE, index=False)
                time.sleep(delay_seconds)

        df.to_excel(LEADS_FILE, index=False)
    except Exception as e:
        with send_lock:
            send_state["log"].append(f"ERROR: {e}")
    finally:
        with send_lock:
            send_state["running"] = False
            send_state["done"] = True


@app.route("/api/send/start", methods=["POST"])
def start_send():
    with send_lock:
        if send_state["running"]:
            return jsonify({"ok": False, "error": "A send job is already running"}), 400
    env = read_env()
    if not env.get("SENDER_EMAIL") or not env.get("SENDER_APP_PASSWORD"):
        return jsonify({"ok": False, "error": "Please set your sender email and app password in Settings first"}), 400
    thread = threading.Thread(target=run_send_job, daemon=True)
    thread.start()
    return jsonify({"ok": True})


@app.route("/api/send/status", methods=["GET"])
def send_status():
    with send_lock:
        return jsonify(dict(send_state))


# ---------- Routes: send follow-ups ----------

def run_followup_job():
    env = read_env()
    df = read_leads_df()

    daily_limit = int(env.get("DAILY_LIMIT", "80") or 80)
    delay_seconds = int(env.get("DELAY_SECONDS", "8") or 8)
    sender_email = env.get("SENDER_EMAIL", "")
    sender_password = env.get("SENDER_APP_PASSWORD", "")
    sender_name = env.get("SENDER_NAME", "Your Company")

    due_df = get_due_followups(df)
    due_idx = list(due_df.index)

    with followup_lock:
        followup_state["running"] = True
        followup_state["done"] = False
        followup_state["sent_count"] = 0
        followup_state["total"] = len(due_idx)
        followup_state["log"] = []

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.starttls(context=context)
            server.login(sender_email, sender_password)

            count = 0
            for i in due_idx:
                row = df.loc[i]
                if count >= daily_limit:
                    with followup_lock:
                        followup_state["log"].append(f"Daily limit of {daily_limit} reached. Stopping.")
                    break

                next_stage = int(row.get("follow_up_stage", 0)) + 1
                tmpl_path = FOLLOWUP_TEMPLATE_FILES.get(next_stage)
                tmpl = read_template(tmpl_path, defaults=DEFAULT_FOLLOWUP_TEMPLATES.get(next_stage))
                to_email = str(row.get("email", "")).strip()

                try:
                    subject = personalize(tmpl["subject"], row, env)
                    body = personalize(tmpl["body"], row, env)
                    from email.mime.multipart import MIMEMultipart
                    from email.mime.text import MIMEText
                    msg = MIMEMultipart()
                    msg["From"] = f"{sender_name} <{sender_email}>"
                    msg["To"] = to_email
                    msg["Subject"] = subject
                    msg.attach(MIMEText(body, "plain"))
                    server.sendmail(sender_email, to_email, msg.as_string())

                    df.at[i, "sent_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    df.at[i, "follow_up_stage"] = next_stage
                    df.at[i, "error"] = ""
                    count += 1
                    with followup_lock:
                        followup_state["sent_count"] = count
                        followup_state["log"].append(
                            f"Follow-up #{next_stage} sent to {row.get('business_name','')} ({to_email})"
                        )
                except Exception as e:
                    df.at[i, "error"] = str(e)
                    with followup_lock:
                        followup_state["log"].append(f"FAILED {row.get('business_name','')}: {e}")

                df.to_excel(LEADS_FILE, index=False)
                time.sleep(delay_seconds)

        df.to_excel(LEADS_FILE, index=False)
    except Exception as e:
        with followup_lock:
            followup_state["log"].append(f"ERROR: {e}")
    finally:
        with followup_lock:
            followup_state["running"] = False
            followup_state["done"] = True


@app.route("/api/followup/start", methods=["POST"])
def start_followup():
    with followup_lock:
        if followup_state["running"]:
            return jsonify({"ok": False, "error": "A follow-up job is already running"}), 400
    env = read_env()
    if not env.get("SENDER_EMAIL") or not env.get("SENDER_APP_PASSWORD"):
        return jsonify({"ok": False, "error": "Please set your sender email and app password in Settings first"}), 400
    thread = threading.Thread(target=run_followup_job, daemon=True)
    thread.start()
    return jsonify({"ok": True})


@app.route("/api/followup/status", methods=["GET"])
def followup_status():
    with followup_lock:
        return jsonify(dict(followup_state))


if __name__ == "__main__":
    print("\n  Your Company Email Marketing is running!")
    print("  Open this in your browser:  http://localhost:5001")
    print("  Follow-ups page:            http://localhost:5001/followups\n")
    app.run(host="127.0.0.1", port=5001, debug=False)
