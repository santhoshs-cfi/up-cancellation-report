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
#  CORE EXECUTION
# ==============================================================================

def get_cookies_with_retry():
    """
    Returns valid session cookies.
    Uses cache first; only opens a browser if no cache exists.
    Login happens once per run even when multiple reports are sent.
    """
    cookies = load_cached_session()
    if cookies:
        return cookies
    log.info("No cached session found — performing fresh login...")
    return login_and_get_cookies()


def run_report(cookies, mode):
    """
    Runs a single report for the given mode.
    Auto re-logins if the session has expired.
    Returns the (possibly refreshed) cookies dict.
    """
    time_period = get_time_period(mode)
    log.info("--- {} report | timePeriod={} ---".format(mode.upper(), time_period))

    success = generate_report(cookies, time_period)

    if not success:
        log.warning("Request failed. Session may be expired. Re-logging in...")
        delete_session()
        cookies = login_and_get_cookies()
        success = generate_report(cookies, time_period)

    if not success:
        log.error("FAILED: {} report failed even after re-login.".format(mode.upper()))
        sys.exit(1)

    return cookies


# ==============================================================================
#  MAIN
# ==============================================================================

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    mode = sys.argv[1].lower().strip()

    # Validate credentials
    if not config.UP_EMAIL or config.UP_EMAIL == "your-email@company.com":
        log.error("Set UP_EMAIL in config.py (or as env var) before running!")
        sys.exit(1)
    if not config.UP_PASSWORD or config.UP_PASSWORD == "your-password-here":
        log.error("Set UP_PASSWORD in config.py (or as env var) before running!")
        sys.exit(1)

    if not config.REPORT_EMAILS:
        log.warning("REPORT_EMAILS is empty — report will be generated but NOT emailed.")

    today = date.today()
    log.info("=" * 60)
    log.info("UrbanPiper Automation | Mode: {} | Date: {} ({})".format(
        mode.upper(), today.isoformat(), today.strftime("%A")
    ))

    # Determine which reports to send
    if mode == "smart":
        modes_to_run = get_smart_modes()
    else:
        modes_to_run = [mode]

    log.info("Reports to send this run: {}".format([m.upper() for m in modes_to_run]))

    # Login once, reuse the session for all reports in this run
    cookies = get_cookies_with_retry()

    for m in modes_to_run:
        cookies = run_report(cookies, m)

    log.info("Done. {} report(s) sent successfully.".format(len(modes_to_run)))
    log.info("=" * 60)


if __name__ == "__main__":
    main()
