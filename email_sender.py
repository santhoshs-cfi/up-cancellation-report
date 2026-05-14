"""
Email Sender
============
Sends per-zone cancellation report CSV files via Gmail OAuth.

If a CSV attachment exceeds SIZE_LIMIT_BYTES (25 MB), the file is uploaded
to Google Drive (folder "UrbanPiper Reports") and a view/download link is
sent in the email body instead.

Credentials come from environment variables (GitHub Secrets):
    GMAIL_CREDENTIALS  - contents of credentials.json  (JSON string)
    GMAIL_TOKEN        - contents of token.json         (JSON string)

Zone recipient emails — checked in this order:
    1. Env vars: NORTH_EMAIL, SOUTH_EMAIL, WEST_EMAIL  (GitHub Secrets)
    2. email_config.json on disk                        (local dev fallback)
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
import email.policy

log = logging.getLogger(__name__)

_SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
_EMAIL_CFG    = os.path.join(_SCRIPT_DIR, "email_config.json")
_GMAIL_SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/drive.file",
]

# Files larger than this go to Google Drive instead of direct attachment
SIZE_LIMIT_BYTES = 25 * 1024 * 1024   # 25 MB


# ==============================================================================
#  LOAD EMAIL CONFIG
# ==============================================================================

def load_email_config():
    """
    Priority order:
      1. NORTH_EMAIL / SOUTH_EMAIL / WEST_EMAIL env vars (GitHub Secrets)
      2. email_config.json on disk (local dev fallback)
    """
    north = os.environ.get("NORTH_EMAIL", "")
    south = os.environ.get("SOUTH_EMAIL", "")
    west  = os.environ.get("WEST_EMAIL",  "")

    if north or south or west:
        log.info("Zone emails loaded from environment variables.")
        mapping = {"North": north, "South": south, "West": west}
        for zone, addr in mapping.items():
            if not addr:
                log.warning("Env var missing for {} zone email.".format(zone))
        return mapping

    if not os.path.exists(_EMAIL_CFG):
        raise FileNotFoundError(
            "No zone emails found.\n"
            "Either set NORTH_EMAIL / SOUTH_EMAIL / WEST_EMAIL env vars\n"
            "or create email_config.json with keys: north_email, south_email, west_email"
        )
    with open(_EMAIL_CFG) as f:
        cfg = json.load(f)
    log.info("Zone emails loaded from email_config.json.")
    mapping = {
        "North": cfg.get("north_email", ""),
        "South": cfg.get("south_email", ""),
        "West":  cfg.get("west_email",  ""),
    }
    for zone, addr in mapping.items():
        if not addr:
            log.warning("email_config.json: {} zone email is empty.".format(zone))
    return mapping


# ==============================================================================
#  CREDENTIALS (shared by Gmail + Drive)
# ==============================================================================

def _get_credentials():
    """Returns google.oauth2.credentials.Credentials with Gmail + Drive scopes."""
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
    except ImportError:
        raise ImportError("Run: pip install google-auth google-auth-oauthlib google-api-python-client")

    creds_json_str = os.environ.get("GMAIL_CREDENTIALS", "")
    token_json_str = os.environ.get("GMAIL_TOKEN", "")

    if creds_json_str and token_json_str:
        log.info("Loading credentials from environment variables.")
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as tf:
            tf.write(token_json_str)
            token_tmp = tf.name
        try:
            creds = Credentials.from_authorized_user_file(token_tmp, _GMAIL_SCOPES)
            if creds.expired and creds.refresh_token:
                log.info("Refreshing expired token...")
                creds.refresh(Request())
        finally:
            os.unlink(token_tmp)
    else:
        token_path = os.path.join(_SCRIPT_DIR, "token.json")
        if not os.path.exists(token_path):
            raise FileNotFoundError(
                "Gmail not authorized. Run: python authorize_gmail.py"
            )
        log.info("Loading credentials from local token.json.")
        creds = Credentials.from_authorized_user_file(token_path, _GMAIL_SCOPES)
        if creds.expired and creds.refresh_token:
            log.info("Refreshing expired token...")
            creds.refresh(Request())
            with open(token_path, "w") as f:
                f.write(creds.to_json())

    return creds


# ==============================================================================
#  GOOGLE DRIVE UPLOAD
# ==============================================================================

def _get_or_create_drive_folder(drive, folder_name):
    """Return the Drive folder ID, creating it if it doesn't exist."""
    q = (
        "name='{}' and mimeType='application/vnd.google-apps.folder'"
        " and trashed=false"
    ).format(folder_name)
    results = drive.files().list(q=q, fields="files(id,name)").execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]
    folder = drive.files().create(
        body={"name": folder_name, "mimeType": "application/vnd.google-apps.folder"},
        fields="id"
    ).execute()
    log.info("Created Drive folder: {}".format(folder_name))
    return folder["id"]


def _cleanup_old_drive_files(drive, folder_id, days_to_keep):
    """Delete files in the Drive folder older than days_to_keep days."""
    from datetime import datetime, timedelta, timezone
    cutoff = (datetime.now(timezone.utc) - timedelta(days=days_to_keep)).isoformat()
    q = "'{}' in parents and trashed=false and createdTime < '{}'".format(folder_id, cutoff)
    results = drive.files().list(q=q, fields="files(id,name,createdTime)").execute()
    old_files = results.get("files", [])
    for f in old_files:
        drive.files().delete(fileId=f["id"]).execute()
        log.info("Deleted old Drive file: {} (created {})".format(f["name"], f["createdTime"]))
    if old_files:
        log.info("Cleaned up {} old file(s) from Drive.".format(len(old_files)))


def _upload_to_drive(creds, csv_bytes, filename, days_to_keep=90):
    """
    Uploads the CSV BytesIO to Google Drive under 'UrbanPiper Reports'.
    Sets 'anyone with link can view' permission.
    Auto-deletes files older than days_to_keep.
    Returns the webViewLink.
    """
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload

    drive     = build("drive", "v3", credentials=creds)
    folder_id = _get_or_create_drive_folder(drive, "UrbanPiper Reports")

    # Clean up old files to save storage
    _cleanup_old_drive_files(drive, folder_id, days_to_keep)

    csv_bytes.seek(0)
    media = MediaIoBaseUpload(
        csv_bytes,
        mimetype="text/csv",
        resumable=True,
    )
    file_meta = {"name": filename, "parents": [folder_id]}
    uploaded  = drive.files().create(
        body=file_meta, media_body=media, fields="id,webViewLink"
    ).execute()

    # Anyone with the link can view (no sign-in required)
    drive.permissions().create(
        fileId=uploaded["id"],
        body={"type": "anyone", "role": "reader"},
    ).execute()

    link = uploaded["webViewLink"]
    log.info("Uploaded to Drive: {} -> {}".format(filename, link))
    return link


# ==============================================================================
#  SEND EMAIL
# ==============================================================================

def send_zone_email(to_email, zone, csv_bytes, filename, date_str, summary_text="", creds=None):
    """
    Sends a zone cancellation report as a CSV attachment.
    - Attaches the CSV directly if <= 25 MB
    - Uploads to Google Drive and sends a link if > 25 MB
    """
    from googleapiclient.discovery import build

    if creds is None:
        creds = _get_credentials()

    gmail = build("gmail", "v1", credentials=creds)

    # ── Check file size ────────────────────────────────────────────────────────
    csv_data  = csv_bytes.getvalue()
    size_mb   = len(csv_data) / (1024 * 1024)
    log.info("Preparing {} — {:.1f} MB".format(filename, size_mb))

    use_drive  = len(csv_data) >= SIZE_LIMIT_BYTES
    drive_link = None

    if use_drive:
        log.info("File exceeds 25 MB — uploading to Google Drive instead.")
        drive_link = _upload_to_drive(creds, csv_bytes, filename)

    # ── Build email body ───────────────────────────────────────────────────────
    if use_drive:
        delivery_note = (
            "NOTE: The CSV report is {:.1f} MB which exceeds Gmail's 25 MB limit.\n"
            "Download it directly from Google Drive (no sign-in required):\n"
            "{}\n\n"
        ).format(size_mb, drive_link)
    else:
        delivery_note = "The CSV report is attached ({:.1f} MB).\n\n".format(size_mb)

    body = (
        "Hi Team,\n\n"
        "Please find the {zone} Zone Cancellation Report for {date}.\n\n"
        "{delivery}"
        "The attached CSV contains ALL cancelled orders across all zones.\n"
        "Use the 'Zone' column to filter for {zone} zone data.\n\n"
        "{summary}\n"
        "This is an automated report. Do not reply to this email.\n"
    ).format(
        zone=zone,
        date=date_str,
        delivery=delivery_note,
        summary=summary_text,
    )

    # ── Compose MIME message ───────────────────────────────────────────────────
    subject = "[{zone} Zone] Cancellation Report - {date}".format(
        zone=zone, date=date_str
    )
    msg = MIMEMultipart()
    msg["To"]      = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    if not use_drive:
        attachment = MIMEBase("text", "csv")
        attachment.set_payload(csv_data)
        encoders.encode_base64(attachment)
        attachment.add_header(
            "Content-Disposition",
            'attachment; filename="{}"'.format(filename)
        )
        msg.attach(attachment)

    # ── Send ───────────────────────────────────────────────────────────────────
    raw    = base64.urlsafe_b64encode(
        msg.as_bytes(policy=email.policy.SMTP)
    ).decode("utf-8")
    result = gmail.users().messages().send(
        userId="me", body={"raw": raw}
    ).execute()

    log.info("Email sent to {} ({} zone) | ID: {} | Via: {}".format(
        to_email, zone, result.get("id"), "Drive link" if use_drive else "attachment"
    ))
    return result.get("id")


# ==============================================================================
#  SEND ALL ZONES
# ==============================================================================

def send_all_zones(zone_csvs, date_str, zone_summaries=None):
    """
    Sends CSV reports to all three zones.
    Builds credentials once, reuses for all zones + any Drive uploads.

    Args:
        zone_csvs     - dict {zone: BytesIO}
        date_str      - human-readable date label for email subject/body
        zone_summaries - dict {zone: str} with plain-text summary for each zone
    """
    email_cfg      = load_email_config()
    zone_summaries = zone_summaries or {}
    creds          = _get_credentials()   # build once, reuse for all zones + Drive

    results = {}
    for zone, csv_buf in zone_csvs.items():
        to_email = email_cfg.get(zone, "")
        if not to_email:
            log.warning("Skipping {} zone - no email configured.".format(zone))
            results[zone] = "skipped (no email)"
            continue

        filename = "{}_Cancellation_Report_{}.csv".format(
            zone, date_str.replace(" ", "_")
        )
        summary = zone_summaries.get(zone, "")

        try:
            send_zone_email(to_email, zone, csv_buf, filename, date_str, summary, creds)
            results[zone] = "sent -> {}".format(to_email)
        except Exception as e:
            log.error("Failed to send {} zone email: {}".format(zone, e))
            results[zone] = "FAILED: {}".format(e)

    return results
