"""
Email Sender
============
Sends per-zone cancellation report Excel files via Gmail OAuth.

If an attachment exceeds SIZE_LIMIT_BYTES (25 MB), the file is uploaded
to Google Drive (folder "UrbanPiper Reports") and a download link is sent
in the email body. Old Drive files are auto-deleted after 90 days.

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
            raise FileNotFoundError("Gmail not authorized. Run: python authorize_gmail.py")
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


def _cleanup_old_drive_files(drive, folder_id, days_to_keep=90):
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
        log.info("Cleaned up {} old Drive file(s).".format(len(old_files)))


def _upload_to_drive(creds, excel_bytes, filename):
    """Upload Excel to Google Drive, set public link, auto-clean old files. Returns link."""
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload

    drive     = build("drive", "v3", credentials=creds)
    folder_id = _get_or_create_drive_folder(drive, "UrbanPiper Reports")

    _cleanup_old_drive_files(drive, folder_id, days_to_keep=90)

    excel_bytes.seek(0)
    media = MediaIoBaseUpload(
        excel_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    uploaded = drive.files().create(
        body={"name": filename, "parents": [folder_id]},
        media_body=media,
        fields="id,webViewLink"
    ).execute()

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

def send_zone_email(to_email, zone, excel_bytes, filename, date_str, top3_stores=None, creds=None):
    """
    Sends a zone cancellation report Excel.
    - Attaches directly if <= 25 MB
    - Uploads to Google Drive and sends link if > 25 MB
    """
    from googleapiclient.discovery import build

    if creds is None:
        creds = _get_credentials()

    gmail = build("gmail", "v1", credentials=creds)

    excel_data = excel_bytes.getvalue()
    size_mb    = len(excel_data) / (1024 * 1024)
    log.info("Preparing {} - {:.1f} MB".format(filename, size_mb))

    use_drive  = len(excel_data) >= SIZE_LIMIT_BYTES
    drive_link = None

    if use_drive:
        log.info("File exceeds 25 MB - uploading to Google Drive.")
        drive_link = _upload_to_drive(creds, excel_bytes, filename)

    # ── Email body ─────────────────────────────────────────────────────────────
    top3_block = ""
    if top3_stores:
        lines = [
            "  {}. {} - {}".format(i + 1, s.get("Store Name", ""), s.get("Count", ""))
            for i, s in enumerate(top3_stores[:3])
        ]
        top3_block = "\nTop Kitchen Cancellation Stores:\n" + "\n".join(lines) + "\n"

    if use_drive:
        delivery_note = (
            "\nThe report file is {:.1f} MB which exceeds Gmail's 25 MB limit.\n"
            "Download it directly from Google Drive (no sign-in required):\n"
            "{}\n"
        ).format(size_mb, drive_link)
    else:
        delivery_note = "\nThe Excel report is attached ({:.1f} MB).\n".format(size_mb)

    body = (
        "Hi Team,\n\n"
        "Please find the {zone} Zone Cancellation Report for {date}.\n"
        "\nThe report contains 2 sheets:\n"
        "  - Sheet 1 (Cancelled Orders): all cancelled orders across all zones\n"
        "    Filter by the 'Zone' column to view {zone} zone data only.\n"
        "  - Sheet 2 (Cancellation Summary): category breakdown, zone pivot,\n"
        "    and top kitchen cancellation stores for {zone} zone.\n"
        "{top3}"
        "{delivery}\n"
        "This is an automated report. Do not reply to this email.\n"
    ).format(zone=zone, date=date_str, top3=top3_block, delivery=delivery_note)

    # ── MIME message ───────────────────────────────────────────────────────────
    msg = MIMEMultipart()
    msg["To"]      = to_email
    msg["Subject"] = "[{zone} Zone] Cancellation Report - {date}".format(
        zone=zone, date=date_str)
    msg.attach(MIMEText(body, "plain", "utf-8"))

    if not use_drive:
        attachment = MIMEBase(
            "application",
            "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        attachment.set_payload(excel_data)
        encoders.encode_base64(attachment)
        attachment.add_header(
            "Content-Disposition",
            'attachment; filename="{}"'.format(filename)
        )
        msg.attach(attachment)

    raw    = base64.urlsafe_b64encode(
        msg.as_bytes(policy=email.policy.SMTP)
    ).decode("utf-8")
    result = gmail.users().messages().send(
        userId="me", body={"raw": raw}
    ).execute()

    log.info("Email sent to {} ({} zone) | ID: {} | Via: {}".format(
        to_email, zone, result.get("id"),
        "Drive link" if use_drive else "attachment"
    ))
    return result.get("id")


# ==============================================================================
#  SEND ALL ZONES
# ==============================================================================

def send_all_zones(zone_excels, date_str, zone_kitchen_top3=None):
    """Sends Excel reports to all zones. Builds credentials once, reuses."""
    email_cfg         = load_email_config()
    zone_kitchen_top3 = zone_kitchen_top3 or {}
    creds             = _get_credentials()

    results = {}
    for zone, excel_buf in zone_excels.items():
        to_email = email_cfg.get(zone, "")
        if not to_email:
            log.warning("Skipping {} zone - no email configured.".format(zone))
            results[zone] = "skipped (no email)"
            continue

        filename = "{}_Cancellation_Report_{}.xlsx".format(
            zone, date_str.replace(" ", "_")
        )
        top3 = zone_kitchen_top3.get(zone, [])

        try:
            send_zone_email(to_email, zone, excel_buf, filename, date_str, top3, creds)
            results[zone] = "sent -> {}".format(to_email)
        except Exception as e:
            log.error("Failed to send {} zone email: {}".format(zone, e))
            results[zone] = "FAILED: {}".format(e)

    return results
