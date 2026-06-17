import argparse
import csv
import os
import queue
import random
import threading
import time
from datetime import datetime, timezone
from email.message import EmailMessage
from typing import List
from typing import Tuple

import pandas as pd
import smtplib

# ===================== CONFIG =====================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT_SSL = 465
CONNECTION_TIMEOUT = 15
CONTACT_PHONE = "(512) 761-6455"

# ===================== TEMPLATES =====================
SUBJECT_TEMPLATES = [
    "Available Loads - Let's Work Together",
    "Steady Freight - Let's Keep Your Trucks Moving",
    "Freight Opportunities in {state} - Let's Connect",
    "Let's Keep Your Fleet Busy - Freight Available Now",
    "Dispatch Support for Your Fleet - Let's Talk Loads",
    "Open Lanes in {state} - Let's Fill Your Schedule",
    "Freight Dispatch Opportunities - Let's Build Together",
    "Let's Maximize Your Miles - Freight Available in {state}",
    "Ready to Dispatch - Let's Get Your Trucks Moving",
    "Let's Partner Up - Freight Opportunities in {state}"
]

BODY_TEMPLATES = [
    """Hi {Name},

I hope this message finds you well. My name is {SENDER_NAME}, and I'm currently looking to partner with reliable carriers for consistent freight dispatch opportunities.

We have loads available in {state}, and I'd love to learn more about your equipment and availability. If you're looking for steady work and fair rates, I believe we can build a strong partnership.

Please let me know if you're interested, and feel free to share your MC number and preferred lanes.

Looking forward to working with you!

Best regards,
{SENDER_NAME}
INDUS TRANSPORTS LLC
{CONTACT_PHONE}
{SENDER_EMAIL}
""",
    """Hi {Name},

I'm reaching out to connect with dependable carriers who are ready for consistent freight opportunities. We're currently dispatching loads in {state}, and I'd love to help keep your trucks running profitably.

If you're open to new lanes or want to maximize your current routes, send me your MC number and ZIP code - I'll match you with the best loads available.

Let's build something solid together.

Best regards,
{SENDER_NAME}
INDUS TRANSPORTS LLC

{CONTACT_PHONE}
{SENDER_EMAIL}
""",
    """Hi {Name},

We're currently dispatching freight in {state}, and I'm looking to team up with reliable carriers who want consistent work and competitive rates.

Whether you run local or OTR, I can help fill your schedule with profitable loads. Just send over your MC number and preferred lanes, and I'll get started.

Excited to work with you!

Best,
{SENDER_NAME}
INDUS TRANSPORTS LLC

{CONTACT_PHONE}
{SENDER_EMAIL}
""",
    """Hi {Name},

I'm {SENDER_NAME} with INDUS TRANSPORTS LLC, and I'm currently onboarding carriers for active freight lanes in {state}.

We offer consistent dispatch, fair rates, and backhaul options to keep your trucks moving. Send me your MC number and ZIP code, and I'll match you with loads that fit your operation.

Looking forward to building a strong partnership.

Best regards,
{SENDER_NAME}
INDUS TRANSPORTS LLC

{CONTACT_PHONE}
{SENDER_EMAIL}
""",
    """Hi {Name},

Are you looking for reliable dispatch support and steady freight? I'm currently working with carriers in {state} and would love to connect.

We offer local and OTR loads tailored to your lanes. Just send me your MC number and ZIP code, and I'll get started right away.

Let's make your routes more profitable.

Best,
{SENDER_NAME}
INDUS TRANSPORTS LLC

{CONTACT_PHONE}
{SENDER_EMAIL}
""",
    """Hi {Name},

We have open lanes and freight ready to move in {state}, and I'm looking for dependable carriers to dispatch.

If you're interested in consistent work and competitive rates, send me your MC number and preferred lanes. I'll match you with loads that fit your fleet.

Let's keep your trucks rolling.

Best regards,
{SENDER_NAME}
INDUS TRANSPORTS LLC

{CONTACT_PHONE}
{SENDER_EMAIL}
""",
    """Hi {Name},

I'm currently expanding my carrier network and would love to work with you. We have freight available in {state} and offer reliable dispatch with fair compensation.

Send me your MC number and ZIP code, and I'll get to work finding the best loads for your lanes.

Looking forward to connecting!

Best,
{SENDER_NAME}
INDUS TRANSPORTS LLC

{CONTACT_PHONE}
{SENDER_EMAIL}
""",
    """Hi {Name},

I'm reaching out to offer dispatch support for carriers operating in {state}. We have freight ready to move and can help you keep your trucks loaded.

Just send me your MC number and ZIP code, and I'll match you with profitable loads that fit your schedule.

Let's make every mile count.

Best regards,
{SENDER_NAME}
INDUS TRANSPORTS LLC

{CONTACT_PHONE}
{SENDER_EMAIL}
""",
    """Hi {Name},

We're actively dispatching freight in {state}, and I'm looking to connect with carriers who want consistent work and reliable support.

If you're interested, send me your MC number and ZIP code, and I'll match you with loads that fit your fleet and schedule.

Let's move freight together.

Best,
{SENDER_NAME}
INDUS TRANSPORTS LLC

{CONTACT_PHONE}
{SENDER_EMAIL}
""",
    """Hi {Name},

I'm {SENDER_NAME} with INDUS TRANSPORTS LLC, and I'm currently looking for carriers to dispatch in {state}. We offer steady loads, fair rates, and personalized support.

Send me your MC number and ZIP code, and I'll get started matching you with the best freight for your lanes.

Looking forward to working together.

Best regards,
{SENDER_NAME}
INDUS TRANSPORTS LLC

{CONTACT_PHONE}
{SENDER_EMAIL}
"""
]

# ===================== UTILITIES =====================

def validate_email(email: str) -> bool:
    if not isinstance(email, str):
        return False
    email = email.strip()
    return "@" in email and "." in email.split("@")[1]

def load_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(path, engine="openpyxl")
    elif ext == ".csv":
        for enc in ["utf-8", "utf-8-sig", "cp1252", "latin-1"]:
            try:
                df = pd.read_csv(path, encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise UnicodeDecodeError("Could not decode CSV file")
    else:
        raise ValueError(f"Unsupported file type for {path}")
    
    # Clean data
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
    return df

def safe_get(value, default=""):
    if pd.isna(value) or value is None:
        return default
    return str(value).strip()

# ===================== LOGGING =====================

class CsvLogger:
    def __init__(self, path: str):
        self.path = path
        self.lock = threading.Lock()
        self._init_file()

    def _init_file(self):
        if not os.path.exists(self.path):
            with open(self.path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["timestamp", "sender", "recipient", "subject", "status", "error"])

    def log(self, sender: str, recipient: str, subject: str, status: str, error: str = ""):
        ts = datetime.now(timezone.utc).isoformat()
        with self.lock:
            with open(self.path, "a", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow([ts, sender, recipient, subject, status, error])

def load_already_sent(path: str) -> set:
    sent = set()
    if os.path.exists(path):
        try:
            df = pd.read_csv(path, encoding="utf-8", on_bad_lines='skip')
            if "recipient" in df.columns:
                for val in df["recipient"].dropna().astype(str):
                    sent.add(val.strip().lower())
            elif "email" in df.columns:
                for val in df["email"].dropna().astype(str):
                    sent.add(val.strip().lower())
        except Exception:
            pass
    return sent

# ===================== SMTP CLIENT =====================

class SmtpClient:
    def __init__(self, email: str, app_password: str):
        self.email = email
        self.app_password = app_password
        self.server = None

    def connect(self):
        try:
            srv = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT_SSL, timeout=CONNECTION_TIMEOUT)
            srv.login(self.email, self.app_password)
            print(f"[OK] SMTP login successful: {self.email}")
            return srv
        except Exception as e:
            print(f"[ERROR] SMTP login failed for {self.email}: {e}")
            raise

    def send_message(self, msg: EmailMessage):
        if self.server is None:
            self.server = self.connect()
        
        try:
            self.server.send_message(msg)
            return True
        except Exception as e:
            print(f"[WARN] Send failed for {self.email}: {e}")
            try:
                self.server.quit()
            except:
                pass
            self.server = None
            return False

    def close(self):
        if self.server:
            try:
                self.server.quit()
            except:
                pass
            self.server = None

# ===================== EMAIL FUNCTIONS =====================

def choose_subject_and_body(state: str) -> Tuple[str, str]:
    subject_template = random.choice(SUBJECT_TEMPLATES)
    subject = subject_template.format(state=state)
    body = random.choice(BODY_TEMPLATES)
    return subject, body

def build_message(sender_name: str, sender_email: str, recipient_email: str, subject: str, body_template: str, row: pd.Series) -> EmailMessage:
    name = safe_get(row.get("Name", "")) or "there"
    state = safe_get(row.get("State", "N/A")).upper() or "N/A"

    body = body_template.format(
        Name=name,
        state=state,
        SENDER_NAME=sender_name,
        SENDER_EMAIL=sender_email,
        CONTACT_PHONE=CONTACT_PHONE,
    )

    msg = EmailMessage()
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = recipient_email
    msg["Subject"] = subject
    msg.set_content(body, charset="us-ascii", cte="7bit")
    return msg

def warm_up_account(smtp_client, sender_email, sender_name):
    try:
        msg = EmailMessage()
        msg["From"] = f"{sender_name} <{sender_email}>"
        msg["To"] = sender_email
        msg["Subject"] = "Test Email - Account Warmup"
        msg.set_content("This is a test email to warm up your account before sending bulk emails.")
        
        if smtp_client.send_message(msg):
            print(f"[OK] Warm-up email sent successfully for {sender_email}")
            time.sleep(random.uniform(30, 60))
            return True
        return False
    except Exception as e:
        print(f"[ERROR] Warm-up failed for {sender_email}: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Single account email sender")
    parser.add_argument("--email", required=True, help="Sender email")
    parser.add_argument("--password", required=True, help="Sender app password")
    parser.add_argument("--name", required=True, help="Sender name")
    parser.add_argument("--recipients", default="recipients.xlsx", help="Path to recipients file")
    parser.add_argument("--log", default="sent_log.csv", help="Central CSV log path")
    parser.add_argument("--min-delay", type=float, default=5.0)
    parser.add_argument("--max-delay", type=float, default=15.0)
    parser.add_argument("--max-per-sender", type=int, default=100)
    parser.add_argument("--retries", type=int, default=2)
    parser.add_argument("--limit", type=int, default=None)
    
    args = parser.parse_args()

    # Load recipients
    try:
        recipients = load_table(args.recipients)
        recipients = recipients.dropna(subset=["Email"])
        recipients["Email"] = recipients["Email"].astype(str).str.strip().str.lower()
    except Exception as e:
        print(f"Error loading recipients: {e}")
        return

    # Load already sent
    already_sent = load_already_sent(args.log)
    print(f"[INFO] Already sent: {len(already_sent)} emails")

    # Filter out already sent
    recipients = recipients[~recipients["Email"].isin(already_sent)]
    
    if args.limit and args.limit > 0:
        recipients = recipients.head(args.limit)
    
    print(f"[INFO] Emails to send: {len(recipients)}")
    
    if len(recipients) == 0:
        print("No recipients to send to")
        return

    logger = CsvLogger(args.log)
    smtp = SmtpClient(args.email, args.password)
    sent_count = 0

    try:
        # Warm up account
        if not warm_up_account(smtp, args.email, args.name):
            print("Account warm-up failed, exiting")
            return

        for _, row in recipients.iterrows():
            recipient = safe_get(row.get("Email", ""))
            if not validate_email(recipient):
                logger.log(args.email, recipient, "", "skip", "invalid email")
                continue

            state = safe_get(row.get("State", "N/A")).upper() or "N/A"
            subject, body_template = choose_subject_and_body(state)
            msg = build_message(args.name, args.email, recipient, subject, body_template, row)

            # Try to send with retries
            success = False
            for attempt in range(args.retries + 1):
                if smtp.send_message(msg):
                    success = True
                    break
                elif attempt < args.retries:
                    time.sleep(3)

            if success:
                logger.log(args.email, recipient, subject, "sent", "")
                sent_count += 1
                print(f"[OK] Sent {sent_count}/{len(recipients)} to {recipient}")
            else:
                logger.log(args.email, recipient, subject, "fail", "All retries failed")
                print(f"[ERROR] Failed to send to {recipient}")

            # Delay between sends
            delay = random.uniform(args.min_delay, args.max_delay)
            time.sleep(delay)

            if sent_count >= args.max_per_sender:
                print(f"[INFO] Reached max per sender limit: {args.max_per_sender}")
                break

    except KeyboardInterrupt:
        print("Interrupted by user")
    except Exception as e:
        print(f"Unexpected error: {e}")
    finally:
        smtp.close()
        print(f"[DONE] Finished! Sent {sent_count} emails from {args.email}")

if __name__ == "__main__":
    main()

