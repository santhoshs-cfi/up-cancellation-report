"""
UrbanPiper Atlas - Automated Report Generator
==============================================
Usage:
    python urbanpiper_report.py smart    -> Auto-detects what to send based on today's date:
                                           Always:        daily report (yesterday's data)
                                           If Monday:   + weekly report (last 7 days)
                                           If 1st:      + monthly report (full prev month)

    python urbanpiper_report.py daily    -> Yesterday's data only
    python urbanpiper_report.py weekly   -> Last 7 days only
    python urbanpiper_report.py monthly  -> Previous full month only
    python urbanpiper_report.py test     -> Same as daily, for testing right now

Requirements:
    pip install playwright requests
    python -m playwright install chromium
"""

import sys
import json
import os
import time
import requests
import logging
from datetime import date, timedelta

# -- Setup logging -------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("urbanpiper_report.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger(__name__)

# -- Import config -------------------------------------------------------------
try:
    import config
except ImportError:
    log.error("config.py not found! Make sure config.py is in the same folder.")
    sys.exit(1)


# ==============================================================================
#  DATE LOGIC
# ==============================================================================

def get_time_period(mode):
    """Returns the timePeriod string expected by the UrbanPiper GraphQL API."""
    today = date.today()

    if mode in ("daily", "test"):
        return "YESTERDAY"

    elif mode == "weekly":
        # Run on Monday -> last 7 days = previous Mon-Sun exactly
        return "LAST_7_DAYS"

    elif mode == "monthly":
        # Calculate exact previous month (handles 28/29/30/31 day months correctly)
        first_of_this_month = today.replace(day=1)
        last_of_prev_month  = first_of_this_month - timedelta(days=1)
        first_of_prev_month = last_of_prev_month.replace(day=1)
        return "{},{}".format(first_of_prev_month.isoformat(), last_of_prev_month.isoformat())

    else:
        raise ValueError("Unknown mode '{}'. Use: smart, daily, weekly, monthly, test".format(mode))


def get_smart_modes():
    """
    Decides which reports to run today:
    - Daily:   always
    - Weekly:  only on Monday (weekday == 0)
    - Monthly: only on 1st of month
    """
    today = date.today()
    modes = ["daily"]

    if today.weekday() == 0:
        modes.append("weekly")
        log.info("Today is Monday -> weekly report will also be sent.")

    if today.day == 1:
        modes.append("monthly")
        log.info("Today is 1st of month -> monthly report will also be sent.")

    return modes


# ==============================================================================
#  SESSION MANAGEMENT
# ==============================================================================

def load_cached_session():
    """Load previously saved cookies from disk. Returns dict or None."""
    session_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.SESSION_FILE)
    if os.path.exists(session_path):
        try:
            with open(session_path) as f:
                data = json.load(f)
            log.info("Loaded cached session from {}".format(config.SESSION_FILE))
            return data
        except Exception as e:
            log.warning("Could not load session cache: {}".format(e))
    return None


def save_session(cookies):
    """Save cookies to disk for reuse."""
    session_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.SESSION_FILE)
    with open(session_path, "w") as f:
        json.dump(cookies, f)
    log.info("Session saved to {}".format(config.SESSION_FILE))


def delete_session():
    """Remove cached session (forces re-login on next attempt)."""
    session_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.SESSION_FILE)
    if os.path.exists(session_path):
        os.remove(session_path)
        log.info("Cached session deleted.")


# ==============================================================================
#  PLAYWRIGHT LOGIN
# ==============================================================================

def login_and_get_cookies():
    """
    Opens a headless Chromium browser, logs into Atlas, returns session cookies.
    Saves cookies to disk for reuse so login only happens when session expires.
    """
    from playwright.sync_api import sync_playwright

    log.info("Starting browser login to UrbanPiper Atlas...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page    = context.new_page()

        try:
            page.goto(config.ATLAS_LOGIN_URL, wait_until="networkidle", timeout=30000)
            log.info("Login page loaded.")

            page.fill(
                'input[type="email"], input[name="email"], input[name="username"], input[placeholder*="mail" i]',
                config.UP_EMAIL
            )
            page.fill('input[type="password"]', config.UP_PASSWORD)
            page.click('button[type="submit"], button:has-text("Login"), button:has-text("Sign in")')
            page.wait_for_url("**/home**", timeout=20000)
            log.info("Login successful!")

        except Exception as e:
            page.screenshot(path="login_debug.png")
            log.error("Login failed: {}".format(e))
            log.error("Screenshot saved as login_debug.png")
            browser.close()
            raise

        raw_cookies = context.cookies()
        browser.close()

    cookies = {c["name"]: c["value"] for c in raw_cookies}
    save_session(cookies)
    log.info("Extracted {} cookies from browser session.".format(len(cookies)))
    return cookies


# ==============================================================================
#  GRAPHQL API CALL
# ==============================================================================

GENERATE_REPORT_MUTATION = """
mutation generateReport($input: ReportGenerationInput!) {
  generateReportV2(input: $input) {
    status {
      success
      messages {
        field
        message
        __typename
      }
      __typename
    }
    __typename
  }
}
"""

def generate_report(cookies, time_period):
    """
    Calls the UrbanPiper GraphQL API to queue report generation.
    UrbanPiper emails the CSV directly to everyone in config.REPORT_EMAILS.
    Returns True on success, False on failure.
    """
    payload = {
        "operationName": "generateReport",
        "variables": {
            "input": {
                "id":           config.REPORT_ID,
                "name":         config.REPORT_NAME,
                "exportFormat": config.EXPORT_FORMAT,
                "emails":       config.REPORT_EMAILS,
                "filters": {
                    "platforms":   config.PLATFORMS,
                    "locations":   config.LOCATION_IDS,
                    "orderStates": config.ORDER_STATES,
                    "timePeriod":  time_period,
                },
                "newVersion": True
            }
        },
        "query": GENERATE_REPORT_MUTATION
    }

    headers = {
        "Content-Type": "application/json",
        "Origin":       "https://atlas.urbanpiper.com",
        "Referer":      "https://atlas.urbanpiper.com/",
        "User-Agent":   "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/148.0.0.0 Safari/537.36",
    }

    log.info("Calling API | timePeriod={} | recipients={}".format(time_period, config.REPORT_EMAILS))

    try:
        response = requests.post(
            config.ATLAS_GRAPHQL_URL,
            json=payload,
            headers=headers,
            cookies=cookies,
            timeout=30
        )
        response.raise_for_status()
        result = response.json()

        gql_status = result.get("data", {}).get("generateReportV2", {}).get("status", {})
        if gql_status.get("success"):
            for m in gql_status.get("messages", []):
                log.info("API: {}".format(m.get("message")))
            log.info("Report queued successfully. UrbanPiper will email it to recipients.")
            return True
        else:
            log.error("API returned failure: {}".format(gql_status.get("messages")))
            return False

    except requests.exceptions.RequestException as e:
        log.error("HTTP error: {}".format(e))
        return False
    except Exception as e:
        log.error("Unexpected error: {}".format(e))
        return False


# ==============================================================================
#  POLL + DOWNLOAD
# ==============================================================================

GET_TASK_LIST_QUERY = """
query getReportsTaskList($limit: Int, $offset: Int, $filters: [ListFilterArgument]) {
  reportTaskList(limit: $limit, offset: $offset, filters: $filters) {
    count
    objects {
      status
      name
      jobId
      message
      downloadUrl
      __typename
    }
    __typename
  }
}
"""

REPORT_DOWNLOAD_BASE = "https://report.svc.urbanpiper.com/v1/report/{job_id}/download"


def _query_task_list(cookies, status_filter=None):
    """
    Calls getReportsTaskList GraphQL query.
    Returns list of task objects: [{status, name, jobId, message, downloadUrl}, ...]
    """
    filters = []
    if status_filter:
        filters = [{"field": "status", "value": status_filter}]

    payload = {
        "operationName": "getReportsTaskList",
        "variables": {"limit": 50, "offset": 0, "filters": filters},
        "query": GET_TASK_LIST_QUERY,
    }
    headers = {
        "Content-Type": "application/json",
        "Origin":       "https://atlas.urbanpiper.com",
        "Referer":      "https://atlas.urbanpiper.com/",
        "User-Agent":   "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/148.0.0.0 Safari/537.36",
    }

    resp = requests.post(
        config.ATLAS_GRAPHQL_URL,
        json=payload,
        headers=headers,
        cookies=cookies,
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    return data.get("data", {}).get("reportTaskList", {}).get("objects", [])


def generate_and_download_csv(cookies, time_period, poll_interval=15, max_wait=600):
    """
    Triggers report generation and polls until the CSV is ready, then downloads it.

    Args:
        cookies       : session cookies dict
        time_period   : Atlas timePeriod string (e.g. "YESTERDAY")
        poll_interval : seconds between status polls (default 15)
        max_wait      : max total seconds to wait (default 600 = 10 minutes)

    Returns:
        (bytes, job_id) — raw CSV bytes and the Atlas job ID
    """
    # Step 1: Queue the report (no email recipients — download-only)
    log.info("Requesting report generation | timePeriod={}".format(time_period))
    payload = {
        "operationName": "generateReport",
        "variables": {
            "input": {
                "id":           config.REPORT_ID,
                "name":         config.REPORT_NAME,
                "exportFormat": config.EXPORT_FORMAT,
                "emails":       [],          # empty — we download directly
                "filters": {
                    "platforms":   config.PLATFORMS,
                    "locations":   config.LOCATION_IDS,
                    "orderSta