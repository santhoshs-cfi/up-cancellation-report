"""
UrbanPiper Atlas - Automated Report Generator
==============================================
Usage:
    python urbanpiper_report.py smart    -> Always daily; +weekly on Mon; +monthly on 1st
    python urbanpiper_report.py daily    -> Yesterday only
    python urbanpiper_report.py weekly   -> Last 7 days
    python urbanpiper_report.py monthly  -> Previous full month
    python urbanpiper_report.py test     -> Same as daily, for quick testing
"""

import sys
import json
import os
import time
import requests
import logging
from datetime import date, timedelta

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("urbanpiper_report.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger(__name__)

try:
    import config
except ImportError:
    log.error("config.py not found!")
    sys.exit(1)


# ==============================================================================
#  DATE LOGIC
# ==============================================================================

def get_time_period(mode):
    today = date.today()
    if mode in ("daily", "test"):
        return "YESTERDAY"
    elif mode == "weekly":
        return "LAST_7_DAYS"
    elif mode == "monthly":
        first_of_this = today.replace(day=1)
        last_of_prev  = first_of_this - timedelta(days=1)
        first_of_prev = last_of_prev.replace(day=1)
        return "{},{}".format(first_of_prev.isoformat(), last_of_prev.isoformat())
    else:
        raise ValueError("Unknown mode: {}".format(mode))


def get_smart_modes():
    today = date.today()
    modes = ["daily"]
    if today.weekday() == 0:
        modes.append("weekly")
        log.info("Today is Monday -> weekly report also.")
    if today.day == 1:
        modes.append("monthly")
        log.info("Today is 1st -> monthly report also.")
    return modes


# ==============================================================================
#  SESSION MANAGEMENT
# ==============================================================================

def load_cached_session():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.SESSION_FILE)
    if os.path.exists(path):
        try:
            with open(path) as f:
                data = json.load(f)
            log.info("Loaded cached session.")
            return data
        except Exception as e:
            log.warning("Could not load session cache: {}".format(e))
    return None


def save_session(cookies):
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.SESSION_FILE)
    with open(path, "w") as f:
        json.dump(cookies, f)
    log.info("Session saved.")


def delete_session():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.SESSION_FILE)
    if os.path.exists(path):
        os.remove(path)
        log.info("Cached session deleted.")


# ==============================================================================
#  PLAYWRIGHT LOGIN
# ==============================================================================

def login_and_get_cookies():
    from playwright.sync_api import sync_playwright
    log.info("Starting browser login...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        try:
            # Navigate to Atlas — it JS-redirects to login.urbanpiper.com
            # Use "commit" so goto returns as soon as the first bytes arrive,
            # then we wait directly for the email input (works across redirects).
            page.goto(config.ATLAS_LOGIN_URL, wait_until="commit", timeout=30000)
            log.info("Initial URL: {}".format(page.url))

            # ── Step 1: Email / Mobile input ──────────────────────────────────
            # Wait for the email field — this naturally waits through the redirect.
            # Broad selector: the input is type=text (accepts email OR mobile).
            email_sel = (
                "input[type=email], input[type=text], input[name=email], "
                "input[name=username], input[placeholder*=mail i], "
                "input[placeholder*=mobile i], input[placeholder*=email i]"
            )
            page.wait_for_selector(email_sel, state="visible", timeout=45000)
            log.info("Email/mobile field visible. URL: {}".format(page.url))
            page.screenshot(path="login_step1_loaded.png")
            page.locator(email_sel).first.fill(config.UP_EMAIL)
            page.screenshot(path="login_step2_email_filled.png")

            # Click the submit button to proceed
            page.locator("button[type=submit], input[type=submit]").first.click()
            log.info("Clicked continue. URL: {}".format(page.url))

            # ── Step 2: Password field ─────────────────────────────────────────
            # May appear on same URL (SPA) or a new /password/ route
            page.wait_for_selector("input[type=password]", state="visible", timeout=30000)
            log.info("Password field visible. URL: {}".format(page.url))
            page.screenshot(path="login_step3_pwd_visible.png")
            page.locator("input[type=password]").first.fill(config.UP_PASSWORD)

            # Submit password — use expect_navigation BEFORE click so we
            # never miss the redirect event (avoids already-fired race condition)
            with page.expect_navigation(
                url="*atlas.urbanpiper.com*",
                wait_until="commit",
                timeout=30000
            ):
                page.locator("button[type=submit], input[type=submit]").first.click()
            page.screenshot(path="login_step4_success.png")
            log.info("Login successful! URL: {}".format(page.url))

        except Exception as e:
            page.screenshot(path="login_debug.png")
            log.error("Login failed: {}".format(e))
            log.error("Current URL: {}".format(page.url))
            browser.close()
            raise

        raw_cookies = context.cookies()
        browser.close()

    cookies = {c["name"]: c["value"] for c in raw_cookies}
    save_session(cookies)
    log.info("Extracted {} cookies.".format(len(cookies)))
    return cookies


# ==============================================================================
#  GRAPHQL DEFINITIONS
# ==============================================================================

GENERATE_REPORT_MUTATION = """
mutation generateReport($input: ReportGenerationInput!) {
  generateReportV2(input: $input) {
    status {
      success
      messages { field message __typename }
      __typename
    }
    __typename
  }
}
"""

GET_TASK_LIST_QUERY = """
query getReportsTaskList($limit: Int, $offset: Int, $filters: [ListFilterArgument]) {
  reportTaskList(limit: $limit, offset: $offset, filters: $filters) {
    count
    objects { status name jobId message downloadUrl __typename }
    __typename
  }
}
"""

REPORT_DOWNLOAD_BASE = "https://report.svc.urbanpiper.com/v1/report/{job_id}/download"

_HEADERS = {
    "Content-Type": "application/json",
    "Origin":       "https://atlas.urbanpiper.com",
    "Referer":      "https://atlas.urbanpiper.com/",
    "User-Agent":   "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
}


# ==============================================================================
#  GENERATE REPORT (email delivery mode)
# ==============================================================================

def generate_report(cookies, time_period):
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
    log.info("Calling API | timePeriod={} | recipients={}".format(time_period, config.REPORT_EMAILS))
    try:
        r = requests.post(config.ATLAS_GRAPHQL_URL, json=payload, headers=_HEADERS, cookies=cookies, timeout=30)
        r.raise_for_status()
        gql = r.json().get("data", {}).get("generateReportV2", {}).get("status", {})
        if gql.get("success"):
            for m in gql.get("messages", []):
                log.info("API: {}".format(m.get("message")))
            log.info("Report queued. UrbanPiper will email it.")
            return True
        else:
            log.error("API failure: {}".format(gql.get("messages")))
            return False
    except Exception as e:
        log.error("Error: {}".format(e))
        return False


# ==============================================================================
#  POLL + DOWNLOAD (used by run_automation.py)
# ==============================================================================

def _query_task_list(cookies):
    payload = {
        "operationName": "getReportsTaskList",
        "variables": {"limit": 50, "offset": 0, "filters": []},
        "query": GET_TASK_LIST_QUERY,
    }
    r = requests.post(config.ATLAS_GRAPHQL_URL, json=payload, headers=_HEADERS, cookies=cookies, timeout=30)
    r.raise_for_status()
    return r.json().get("data", {}).get("reportTaskList", {}).get("objects", [])


def generate_and_download_csv(cookies, time_period, poll_interval=15, max_wait=600):
    """
    Generates the report via GraphQL, polls until complete, then downloads the CSV.
    Returns (csv_bytes, job_id).
    """
    # Step 1: Queue report (no email recipients)
    log.info("Queuing report | timePeriod={}".format(time_period))
    payload = {
        "operationName": "generateReport",
        "variables": {
            "input": {
                "id":           config.REPORT_ID,
                "name":         config.REPORT_NAME,
                "exportFormat": config.EXPORT_FORMAT,
                "emails":       [],
                "filters": {
                    "platforms":   config.PLATFORMS,
                    "locations":   config.LOCATION_IDS,
                    "orderStates": config.ORDER_STATES,
                    "timePeriod":  time_period,
                },
                "newVersion": True,
            }
        },
        "query": GENERATE_REPORT_MUTATION,
    }
    r = requests.post(config.ATLAS_GRAPHQL_URL, json=payload, headers=_HEADERS, cookies=cookies, timeout=30)
    r.raise_for_status()
    gql = r.json().get("data", {}).get("generateReportV2", {}).get("status", {})
    if not gql.get("success"):
        raise RuntimeError("generateReportV2 failed: {}".format(gql.get("messages")))
    log.info("Report queued. Polling every {}s (max {}s)...".format(poll_interval, max_wait))

    # Step 2: Poll for completion
    deadline  = time.time() + max_wait
    found_job = None

    while time.time() < deadline:
        time.sleep(poll_interval)
        tasks = _query_task_list(cookies)
        for task in tasks:
            if config.REPORT_NAME.lower() in task.get("name", "").lower():
                found_job = task
                break

        if not found_job:
            log.info("Not in task list yet -- waiting...")
            continue

        status = found_job.get("status", "")
        job_id = found_job.get("jobId", "")
        log.info("Status: {} | jobId: {}".format(status, job_id))

        if status == "completed":
            break
        elif status == "failed":
            raise RuntimeError("Atlas report FAILED: {}".format(found_job.get("message")))

        found_job = None   # reset, keep polling
    else:
        raise TimeoutError("Report did not complete within {} seconds.".format(max_wait))

    job_id       = found_job.get("jobId", "")
    download_url = found_job.get("downloadUrl") or REPORT_DOWNLOAD_BASE.format(job_id=job_id)
    log.info("Downloading from: {}".format(download_url))

    # Step 3: Download CSV
    dl = requests.get(download_url, cookies=cookies, headers={"User-Agent": _HEADERS["User-Agent"]}, timeout=120, stream=True)
    dl.raise_for_status()
    csv_bytes = dl.content
    log.info("Downloaded {:,} bytes.".format(len(csv_bytes)))
    return csv_bytes, job_id


# ==============================================================================
#  CORE EXECUTION
# ==============================================================================

def get_cookies_with_retry():
    cookies = load_cached_session()
    if cookies:
        return cookies
    log.info("No cached session -- fresh login...")
    return login_and_get_cookies()


def run_report(cookies, mode):
    time_period = get_time_period(mode)
    log.info("--- {} | timePeriod={} ---".format(mode.upper(), time_period))
    success = generate_report(cookies, time_period)
    if not success:
        log.warning("Failed -- re-logging in...")
        delete_session()
        cookies = login_and_get_cookies()
        success = generate_report(cookies, time_period)
    if not success:
        log.error("FAILED after re-login: {}".format(mode.upper()))
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

    if not config.UP_EMAIL or config.UP_EMAIL == "your-email@company.com":
        log.error("UP_EMAIL not set!")
        sys.exit(1)
    if not config.UP_PASSWORD or config.UP_PASSWORD == "your-password-here":
        log.error("UP_PASSWORD not set!")
        sys.exit(1)

    today = date.today()
    log.info("=" * 60)
    log.info("Mode: {} | Date: {} ({})".format(mode.upper(), today.isoformat(), today.strftime("%A")))

    modes_to_run = get_smart_modes() if mode == "smart" else [mode]
    log.info("Reports: {}".format([m.upper() for m in modes_to_run]))

    cookies = get_cookies_with_retry()
    for m in modes_to_run:
        cook