"""
Email Sender
============
Sends per-zone cancellation report Excel files via Gmail OAuth.

Credentials come from environment variables (GitHub Secrets):
    GMAIL_CREDENTIALS  - contents of credentials.json  (JSON string)
    GMAIL_TOKEN        - contents of token.json         (JSON string)

Zone recipient emails come from email_config.json (editable flat file):
    { "north_email": "...", "south_email": "...", "west_email": "..." }
"""

import base64
import json
import os
import tempfile
import logging
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders

log = logging.getLogger(__name__)

_SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
_EMAIL_CFG   = os.path.join(_SCRIPT_DIR, "email_config.json")
_GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


# ==============================================================================
#  LOAD EMAIL CONFIG
# ==============================================================================

def load_email_config():
    """
    Returns dict: {"north": "...", "south": "...", "west": "..."}
    Reads email_config.json from the same folder as this script.
    """
    if not os.path.exists(_EMAIL_CFG):
        raise FileNotFoundError(
            "email_config.json not found at {}. "
            "Create it with keys: north_email, south_email, west_email".format(_EMAIL_CFG)
        )
    with open(_EMAIL_CFG) as f:
        cfg = json.load(f)

    mapping = {
        "North": cfg.get("north_email", ""),
        "South": cfg.get("south_email", ""),
        "West":  cfg.get("west_email",  ""),
    }
    for zone, email in mapping.items():
        if not email:
            log.warning("email_config.json: {} zone email is empty".format(zone))
    return mapping


# ==============================================================================
#  GMAIL SERVICE
# ==============================================================================

def _build_gmail_service():
    """
    Builds and returns an authenticated Gmail API service.

    Credential source (in priority order):
      1. GMAIL_CREDENTIALS + GMAIL_TOKEN environment variables (GitHub Actions)
      2. credentials.json + token.json files on disk (local dev)
    """
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
    except ImportError:
        raise ImportError(
            "Google API packages not installed. "
            "Run: pip install google-auth google-auth-oauthlib google-api-python-client"
        )

    creds = None

    # ── Option 1: env vars (GitHub Actions) ──────────────────────────────────
    creds_json_str  = os.environ.get("GMAIL_CREDENTIALS", "")
    token_json_str  = os.environ.get("GMAIL_TOKEN", "")

    if creds_json_str and token_json_str:
        log.info("Loading Gmail credentials from environment variables.")
        # Write to temp files (google-auth library needs file paths)
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as tf:
            tf.write(token_json_str)
            token_tmp = tf.name
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as cf:
            cf.write(creds_json_str)
            creds_tmp = cf.name

        try:
            creds = Credentials.from_authorized_user_file(token_tmp, _GMAIL_SCOPES)
            if creds.expired and creds.refresh_token:
                log.info("Refreshing expired Gmail token...")
                creds.refresh(Request())
        finally:
            os.unlink(token_tmp)
            os.unlink(creds_tmp)

    # ── Option 2: local files (dev machine) ──────────────────────────────────
    else:
        token_path = os.path.join(_SCRIPT_DIR, "token.json")
        creds_path = os.path.join(_SCRIPT_DIR, "credentials.json")

        if not os.path.exists(token_path):
            raise FileNotFoundError(
                "Gmail not authorized. Run: python authorize_gmail.py\n"
                "(Or set GMAIL_CREDENTIALS + GMAIL_TOKEN env vars for GitHub Actions)"
            )
        if not os.path.exists(creds_path):
            raise FileNotFoundError(
                "credentials.json not found. Download from Google Cloud Console."
            )

        log.info("Loading Gmail credentials from local files.")
        creds = Credentials.from_authorized_user_file(token_path, _GMAIL_SCOPES)
        if creds.expired and creds.refresh_token:
            log.info("Refreshing expired Gmail token...")
            creds.refresh(Request())
            with open(token_path, "w") as f:
                f.write(creds.to_json())

    service = build("gmail", "v1", credentials=creds)
    return service


# ==============================================================================
#  SEND EMAIL
# ==============================================================================

def send_zone_email(to_email, zone, excel_bytes, filename, date_str, top3_stores=None):
    """
    Sends a zone cancellation report as an email attachment.

    Args:
        to_email    : recipient email address
        zone        : "North", "South", or "West"
        excel_bytes : BytesIO object with Excel file contents
        filename    : attachment filename (e.g. "North_Cancellation_2026-05-13.xlsx")
        date_str    : human-readable date/range string for email body
        top3_stores : optional list of dicts with "Store Name" and "Count" keys
    """
    service = _build_gmail_service()

    # ── Build email body ──────────────────────────────────────────────────────
    top3_block = ""
    if top3_stores:
        lines = [
            "  {}. {} — {}".format(i + 1, s["Store Name"], s["Count"])
            for i, s in enumerate(top3_stores[:3])
        ]
        top3_block = "\nTop Kitchen Cancellation Stores:\n" + "\n".join(lines) + "\n"

    body = (
        "Hi Team,\n\n"
        "Please find attached the {zone} Zone Cancellation Report for {date}.\n"
        "{top3}\n"
        "This is an automated report. Do not reply to this email.\n"
    ).format(zone=zone, date=date_str, top3=top3_block)

    # ── Compose MIME message ──────────────────────────────────────────────────
    msg = MIMEMultipart()
    msg["To"]      = to_email
    msg["Subject"] = "[{zone} Zone] Cancellation Report — {date}".format(
        zone=zone, date=date_str
    )
    msg.attach(MIMEText(body, "plain"))

    attachment = MIMEBase("application", "octet-stream")
    attachment.set_payload(excel_bytes.getvalue())
    encoders.encode_base64(attachment)
    attachment.add_header(
        "Content-Disposition",
        "attachment; filename={}".format(filename)
    )
    msg.attach(attachment)

    # ── Send ──────────────────────────────────────────────────────────────────
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw}).execute()
    log.info("Email sent to {} ({} zone)".format(to_email, zone))


def send_all_zones(zone_excels, date_str, zone_kitchen_top3=None):
    """
    Sends reports to all three zones.

    Args:
        zone_excels       : dict {"North": BytesIO, "South": BytesIO, "West": BytesIO}
        date_str          : human-readable date label (for subject + body)
        zone_kitchen_top3 : optional dict {"North": [...], "South": [...], "West": [...]}
    """
    email_cfg   = load_email_config()
    zone_kitchen_top3 = zone_kitchen_top3 or {}

    results = {}
    for zone, excel_buf in zone_excels.items():
        to_email = email_cfg.get(zone, "")
        if not to_email:
            log.warning("Skipping {} zone — no email configured.".format(zone))
            results[zone] = "skipped (no email)"
            continue

        filename = "{}_Cancellation_Report_{}.xlsx".format(zone, date_str.replace(" ", "_"))
        top3     = zone_kitchen_top3.get(zone, [])

        try:
            send_zone_email(to_email, zone, excel_buf, filename, date_str, top3)
            results[zone] = "sent → {}".format(to_email)
        except Exception as e:
            log.error("Failed to send {} zone email: {}".format(zone, e))
            results[zone] = "FAILED: {}".format(e)

    return results
