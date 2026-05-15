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
#  HTML TABLE BUILDERS (for email body)
# ==============================================================================

_TABLE_CSS = (
    "border-collapse:collapse;font-family:Arial,sans-serif;"
    "font-size:13px;width:100%"
)
_TH = (
    "style='background:#1F4E79;color:#fff;padding:7px 12px;"
    "text-align:{}border:1px solid #ccc;white-space:nowrap'"
)
_TD      = "style='padding:6px 12px;border:1px solid #ddd;text-align:{};white-space:nowrap'"
_TD_ALT  = "style='padding:6px 12px;border:1px solid #ddd;text-align:{};white-space:nowrap;background:#f5f8fc'"
_TD_BOLD = "style='padding:6px 12px;border:1px solid #ccc;text-align:{};font-weight:bold;background:#e8eef5'"


def _th(text, align="center"):
    return "<th {}>{}".format(_TH.format(align + ";"), text) + "</th>"


def _td(text, align="center", alt=False, bold=False):
    if bold:
        return "<td {}>{}</td>".format(_TD_BOLD.format(align + ";"), text)
    style = _TD_ALT if alt else _TD
    return "<td {}>{}</td>".format(style.format(align + ";"), text)


def _build_summary_html(pivot_df):
    rows_html = ""
    totals = {"Total Cancelled", "Total Orders"}
    for i, (_, row) in enumerate(pivot_df.iterrows()):
        cat   = str(row["Cancellation Category"])
        count = "{:,}".format(int(row["Count"]))
        p_ord = row["% of Total Orders"]
        p_can = row["% of Total Cancelled"]
        pct_o = "{:.2f}%".format(float(p_ord)) if str(p_ord) != "-" else "-"
        pct_c = "{:.2f}%".format(float(p_can)) if str(p_can) != "-" else "-"
        bold  = cat in totals
        alt   = (i % 2 == 0) and not bold
        rows_html += "<tr>{}{}{}{}</tr>".format(
            _td(cat,   "left",   alt, bold),
            _td(count, "center", alt, bold),
            _td(pct_o, "center", alt, bold),
            _td(pct_c, "center", alt, bold),
        )
    return (
        "<h3 style='font-family:Arial;color:#1F4E79;margin:24px 0 6px'>"
        "Cancellation Category Summary</h3>"
        "<table style='{}'><tr>{}{}{}{}</tr>{}</table>"
    ).format(
        _TABLE_CSS,
        _th("Category",         "left"),
        _th("Count"),
        _th("% of Total Orders"),
        _th("% of Total Cancelled"),
        rows_html,
    )


def _build_zone_html(zone_pivot_df, zones=("North", "South", "West")):
    hdrs = _th("Category", "left")
    for z in zones:
        hdrs += _th(z) + _th("% " + z)
    hdrs += _th("Grand Total") + _th("Overall %")

    rows_html = ""
    for i, (_, row) in enumerate(zone_pivot_df.iterrows()):
        cat      = str(row["Cancellation Category"])
        is_grand = cat == "Grand Total"
        alt      = (i % 2 == 0) and not is_grand
        cells    = _td(cat, "left", alt, is_grand)
        for z in zones:
            cells += _td("{:,}".format(int(row.get(z, 0))),          "center", alt, is_grand)
            cells += _td("{:.2f}%".format(float(row.get("% "+z,0))), "center", alt, is_grand)
        cells += _td("{:,}".format(int(row.get("Grand Total", 0))),   "center", alt, is_grand)
        cells += _td("{:.2f}%".format(float(row.get("Overall %", 0))),"center", alt, is_grand)
        rows_html += "<tr>{}</tr>".format(cells)

    return (
        "<h3 style='font-family:Arial;color:#1F4E79;margin:24px 0 6px'>"
        "Zone-wise Cancellation Breakdown</h3>"
        "<table style='{}'><tr>{}</tr>{}</table>"
    ).format(_TABLE_CSS, hdrs, rows_html)


def _build_top3_html(top3_stores, zone):
    if not top3_stores:
        return ""
    items = "".join(
        "<li style='margin:3px 0'><b>{}</b> &mdash; {:,} cancellations</li>".format(
            s.get("Store Name", ""), int(s.get("Count", 0))
        )
        for s in top3_stores[:3]
    )
    return (
        "<h3 style='font-family:Arial;color:#1F4E79;margin:24px 0 6px'>"
        "Top Kitchen Cancellation Stores &mdash; {} Zone</h3>"
        "<ol style='font-family:Arial;font-size:13px;margin:0;padding-left:20px'>"
        "{}</ol>"
    ).format(zone, items)


# ==============================================================================
#  SEND EMAIL
# ==============================================================================

def send_zone_email(to_email, zone, excel_bytes, filename, date_str,
                    top3_stores=None, creds=None, pivot_df=None, zone_pivot_df=None,
                    mode="daily"):
    """
    Sends a zone cancellation report Excel.
    - Attaches directly if <= 25 MB
    - Uploads to Google Drive and sends link if > 25 MB
    - Includes category summary + zone breakdown tables in email body
    - mode: daily / weekly / monthly — used in subject line
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

    # ── Build HTML email body ──────────────────────────────────────────────────
    summary_html = ""
    if pivot_df is not None:
        try:
            summary_html = _build_summary_html(pivot_df)
        except Exception as e:
            log.warning("Could not build summary table: {}".format(e))

    zone_html = ""
    if zone_pivot_df is not None:
        try:
            zone_html = _build_zone_html(zone_pivot_df)
        except Exception as e:
            log.warning("Could not build zone table: {}".format(e))

    top3_html = ""
    if top3_stores:
        try:
            top3_html = _build_top3_html(top3_stores, zone)
        except Exception as e:
            log.warning("Could not build top-3 block: {}".format(e))

    if use_drive:
        delivery_html = (
            "<p style='font-family:Arial;font-size:14px'>"
            "The report file is <b>{:.1f} MB</b> which exceeds Gmail's 25&nbsp;MB limit.<br>"
            "Download it directly from Google Drive (no sign-in required):<br>"
            "<a href='{link}'>{link}</a></p>"
        ).format(size_mb, link=drive_link)
    else:
        delivery_html = (
            "<p style='font-family:Arial;font-size:14px'>"
            "The Excel report is attached (<b>{:.1f}&nbsp;MB</b>).</p>"
        ).format(size_mb)

    body = """
<html><body style='margin:0;padding:16px;background:#fff'>
<p style='font-family:Arial;font-size:14px'>Hi Team,</p>
<p style='font-family:Arial;font-size:14px'>
Please find the <b>{zone} Zone</b> Cancellation Report for <b>{date}</b>.
</p>
{delivery}
{summary}
{zone_table}
{top3}
<p style='font-family:Arial;font-size:12px;color:#888;margin-top:32px'>
This is an automated report. Do not reply to this email.
</p>
</body></html>
""".format(
        zone=zone, date=date_str,
        delivery=delivery_html,
        summary=summary_html,
        zone_table=zone_html,
        top3=top3_html,
    )

    # ── MIME message ───────────────────────────────────────────────────────────
    mode_label = {"daily": "Daily", "weekly": "Weekly", "monthly": "Monthly",
                  "test": "Daily"}.get(mode.lower(), "Daily")
    msg = MIMEMultipart()
    msg["To"]      = to_email
    msg["Subject"] = "[{zone} Zone] {label} Cancellation Report - {date}".format(
        zone=zone, label=mode_label, date=date_str)
    msg.attach(MIMEText(body, "html", "utf-8"))

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

def send_all_zones(zone_excels, date_str, zone_kitchen_top3=None,
                   pivot_df=None, zone_pivot_df=None, mode="daily"):
    """Sends Excel reports to all zones. Builds credentials once, reuses.

    pivot_df and zone_pivot_df are included as plain-text tables in the body.
    mode (daily/weekly/monthly) is used in the email subject line.
    If TEST_EMAIL env var is set, every zone report is also sent to that
    address. Remove the secret from GitHub when testing is done.
    """
    email_cfg         = load_email_config()
    zone_kitchen_top3 = zone_kitchen_top3 or {}
    creds             = _get_credentials()
    test_email        = os.environ.get("TEST_EMAIL", "").strip()

    if test_email:
        log.info("TEST_EMAIL is set - all zone reports will also go to {}.".format(test_email))

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
            send_zone_email(to_email, zone, excel_buf, filename, date_str,
                            top3, creds, pivot_df, zone_pivot_df, mode)
            results[zone] = "sent -> {}".format(to_email)
        except Exception as e:
            log.error("Failed to send {} zone email: {}".format(zone, e))
            results[zone] = "FAILED: {}".format(e)

        # Also send to test email if configured
        if test_email and test_email != to_email:
            try:
                excel_buf.seek(0)
                send_zone_email(test_email, zone, excel_buf, filename, date_str,
                                top3, creds, pivot_df, zone_pivot_df, mode)
                log.info("Test copy sent to {} ({} zone).".format(test_email, zone))
            except Exception as e:
                log.error("Failed to send test copy for {} zone: {}".format(zone, e))

    return results
