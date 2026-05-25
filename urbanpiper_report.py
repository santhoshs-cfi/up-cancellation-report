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
    """
    Returns the timePeriod value to send to UrbanPiper API.

    Daily/weekly use UrbanPiper's own recognised constants (YESTERDAY, LAST_7_DAYS)
    because the API generates those quickly from cache.  Custom date strings cause
    UrbanPiper to spin up a cold query that can take >10 min and never appear in
    the task list.

    Monthly uses explicit dates because there is no LAST_MONTH constant.

    The snapshot-based polling + date validation downstream ensures we always
    download the correct freshly-generated report, not a stale cached one.
    """
    today = date.today()
    if mode in ("daily", "test"):
        # YESTERDAY constant = previous day, recognised and cached by UrbanPiper
        return "YESTERDAY"
    elif mode == "weekly":
        # Explicit Mon-Sun of previous week (today is always Monday when this runs)
        # today-1 = Sunday, today-7 = previous Monday
        end   = today - timedelta(days=1)   # Sunday
        start = today - timedelta(days=7)   # Monday of previous week
        return "{},{}".format(start.isoformat(), end.isoformat())
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


def get_auto_modes():
    """Auto mode: weekly on Monday, monthly on 1st — never daily.
    Used by cron-job.org Job 2 (fires on Mon + 1st via cron: 30 3 1 * 1).
    Returns empty list on any other day so the pipeline exits cleanly.
    """
    today = date.today()
    modes = []
    if today.weekday() == 0:
        modes.append("weekly")
        log.info("Auto mode: Monday -> running weekly report.")
    if today.day == 1:
        modes.append("monthly")
        log.info("Auto mode: 1st of month -> running monthly report.")
    if not modes:
        log.info("Auto mode: today ({}) is neither Monday nor 1st — skipping.".format(
            today.strftime("%A %d %b")))
    return modes


# ==============================================================================
#  SESSION MANAGEMENT
#  Session file stores: {"cookies": {name: value}, "token": "JWT string or null"}
# ==============================================================================

def load_cached_session():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.SESSION_FILE)
    if os.path.exists(path):
        try:
            with open(path) as f:
                data = json.load(f)
            # Support new format {"cookies": {...}, "token": "..."}
            if isinstance(data, dict) and "cookies" in data:
                log.info("Loaded cached session (token present: {}).".format(
                    bool(data.get("token"))))
                return data
            # Old format: just a flat cookies dict — discard it
            log.info("Old session format detected — will re-login.")
        except Exception as e:
            log.warning("Could not load session cache: {}".format(e))
    return None


def save_session(session):
    """session = {"cookies": {...}, "token": "JWT or None"}"""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.SESSION_FILE)
    with open(path, "w") as f:
        json.dump(session, f)
    log.info("Session saved (token present: {}).".format(bool(session.get("token"))))


def delete_session():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.SESSION_FILE)
    if os.path.exists(path):
        os.remove(path)
        log.info("Cached session deleted.")


# ==============================================================================
#  PLAYWRIGHT LOGIN
# ==============================================================================

def login_and_get_session():
    """
    Logs in via Playwright headless browser.
    Returns session dict: {"cookies": {name: value, ...}, "token": "JWT or None"}

    Auth flow:
      1. Browser navigates to atlas.urbanpiper.com → redirects to login.urbanpiper.com
      2. Two-step form: email → Continue → password → Submit
      3. Login.urbanpiper.com redirects to atlas.urbanpiper.com/auth-service?access_token=JWT
      4. Atlas SPA calls authServiceLogin GraphQL mutation → gets a session JWT back
      5. JWT stored in localStorage; sent as "Authorization: JWT <token>" on every API call
    We capture the JWT from localStorage after login completes.
    """
    from playwright.sync_api import sync_playwright
    log.info("Starting browser login...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        # Intercept authServiceLogin response to capture the JWT
        captured_token = {"value": None}

        def handle_response(response):
            if "atlas-backend.svc.urbanpiper.com/graphql" not in response.url:
                return
            try:
                body = response.json()
                data = (body.get("data") or {})
                login_data = data.get("loginAuthServiceAccess", {})
                if login_data and login_data.get("status", {}).get("success"):
                    token = login_data.get("token")
                    if token:
                        captured_token["value"] = token
                        log.info("Captured JWT from authServiceLogin response.")
            except Exception:
                pass

        page.on("response", handle_response)

        try:
            # Navigate to Atlas — JS-redirects to login.urbanpiper.com
            page.goto(config.ATLAS_LOGIN_URL, wait_until="commit", timeout=30000)
            log.info("Initial URL: {}".format(page.url))

            # ── Step 1: Email / Mobile input ──────────────────────────────────
            email_sel = (
                "input[type=email], input[type=text], input[name=email], "
                "input[name=username], input[placeholder*=mail i], "
                "input[placeholder*=mobile i], input[placeholder*=email i]"
            )
            page.wait_for_selector(email_sel, state="visible", timeout=45000)
            log.info("Email field visible. URL: {}".format(page.url))
            page.screenshot(path="login_step1_loaded.png")
            page.locator(email_sel).first.fill(config.UP_EMAIL)
            page.screenshot(path="login_step2_email_filled.png")

            page.locator("button[type=submit], input[type=submit]").first.click()
            log.info("Clicked continue.")

            # ── Step 2: Password field ─────────────────────────────────────────
            page.wait_for_selector("input[type=password]", state="visible", timeout=30000)
            log.info("Password field visible. URL: {}".format(page.url))
            page.screenshot(path="login_step3_pwd_visible.png")
            page.locator("input[type=password]").first.fill(config.UP_PASSWORD)

            page.locator("button[type=submit], input[type=submit]").first.click()

            # ── Step 3: Business selector (may appear if account has multiple businesses) ──
            # Wait briefly to see if the business selector page appears
            try:
                page.wait_for_url("**/business**", timeout=8000)
                log.info("Business selector page detected. URL: {}".format(page.url))
                page.screenshot(path="login_step3b_business_selector.png")

                # Click on "Eatfit - MOC"
                eatfit_locator = page.locator("text=Eatfit - MOC")
                eatfit_locator.wait_for(state="visible", timeout=10000)
                eatfit_locator.click()
                log.info("Clicked 'Eatfit - MOC' business.")
            except Exception:
                # Business selector did not appear — single business account, continue normally
                log.info("No business selector — proceeding directly.")

            # Wait until we're back on atlas.urbanpiper.com (past all redirects)
            page.wait_for_function(
                "window.location.hostname === 'atlas.urbanpiper.com' "
                "&& window.location.pathname !== '/login'",
                timeout=45000
            )
            page.screenshot(path="login_step4_success.png")
            log.info("Login successful! URL: {}".format(page.url))

            # Give the SPA a moment to fire authServiceLogin and store the token
            page.wait_for_timeout(3000)

            # Fallback: try to read token directly from localStorage
            if not captured_token["value"]:
                log.info("Trying to read token from localStorage...")
                try:
                    storage_json = page.evaluate(
                        "() => JSON.stringify(Object.assign({}, localStorage))"
                    )
                    storage = json.loads(storage_json)
                    log.info("localStorage keys: {}".format(list(storage.keys())))
                    for key, val in storage.items():
                        if isinstance(val, str) and len(val) > 50:
                            # JWT tokens are long base64 strings
                            if val.startswith("eyJ") or "token" in key.lower() or "auth" in key.lower():
                                captured_token["value"] = val
                                log.info("Found token in localStorage key: {}".format(key))
                                break
                except Exception as e:
                    log.warning("localStorage read failed: {}".format(e))

        except Exception as e:
            page.screenshot(path="login_debug.png")
            log.error("Login failed: {}".format(e))
            log.error("Current URL: {}".format(page.url))
            browser.close()
            raise

        raw_cookies = context.cookies()
        browser.close()

    cookies = {c["name"]: c["value"] for c in raw_cookies}
    token   = captured_token["value"]

    if token:
        log.info("JWT token captured successfully (length {}).".format(len(token)))
    else:
        log.warning("No JWT token captured — API calls may fail. "
                    "Proceeding with cookies-only session.")

    session = {"cookies": cookies, "token": token}
    save_session(session)
    log.info("Extracted {} cookies.".format(len(cookies)))
    return session


# ==============================================================================
#  API HEADERS
# ==============================================================================

def _make_headers(token=None):
    """Build request headers, adding Authorization if we have a JWT."""
    headers = {
        "Content-Type":   "application/json",
        "Origin":         "https://atlas.urbanpiper.com",
        "Referer":        "https://atlas.urbanpiper.com/",
        "User-Agent":     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "sec-fetch-site": "same-site",
        "sec-fetch-mode": "cors",
        "sec-fetch-dest": "empty",
    }
    if token:
        headers["Authorization"] = "JWT {}".format(token)
    return headers


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


# ==============================================================================
#  GENERATE REPORT (email delivery mode — used by urbanpiper_report.py main)
# ==============================================================================

def generate_report(session, time_period):
    cookies = session.get("cookies", {})
    token   = session.get("token")
    headers = _make_headers(token)

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
        r = requests.post(config.ATLAS_GRAPHQL_URL, json=payload, headers=headers,
                          cookies=cookies, timeout=30)
        r.raise_for_status()
        resp_json = r.json()
        if resp_json.get("data") is None:
            log.error("GraphQL returned data=null. Full response: {}".format(r.text[:500]))
            errors = resp_json.get("errors", [])
            if errors:
                log.error("GraphQL errors: {}".format(errors))
            return False
        gql = (resp_json.get("data") or {}).get("generateReportV2", {}).get("status", {})
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

def _query_task_list(session):
    cookies = session.get("cookies", {})
    token   = session.get("token")
    headers = _make_headers(token)

    payload = {
        "operationName": "getReportsTaskList",
        "variables": {"limit": 200, "offset": 0, "filters": []},
        "query": GET_TASK_LIST_QUERY,
    }
    r = requests.post(config.ATLAS_GRAPHQL_URL, json=payload, headers=headers,
                      cookies=cookies, timeout=30)
    r.raise_for_status()
    resp_json = r.json()
    if resp_json.get("data") is None:
        log.error("getReportsTaskList returned data=null: {}".format(r.text[:300]))
        return []
    return (resp_json.get("data") or {}).get("reportTaskList", {}).get("objects", [])


def generate_and_download_csv(session, time_period, poll_interval=15, max_wait=1200):
    """
    Generates the report via GraphQL, polls until complete, then downloads the CSV.
    Returns (csv_bytes, job_id).
    """
    cookies = session.get("cookies", {})
    token   = session.get("token")
    headers = _make_headers(token)

    # Step 1a: Snapshot existing jobs BEFORE queuing so we can identify the new one
    existing_ids = set()
    try:
        for t in _query_task_list(session):
            jid = t.get("jobId", "")
            if jid:
                existing_ids.add(jid)
        log.info("Pre-queue snapshot: {} existing job(s).".format(len(existing_ids)))
    except Exception as e:
        log.warning("Could not snapshot task list: {}".format(e))

    # Step 1b: Queue report
    # We must pass at least one email so UrbanPiper creates a visible task entry
    # in the task list. Without an email, the report is queued silently and never
    # appears — causing the poll to wait forever.
    # UP_EMAIL is used as the notify address; UrbanPiper will also send a raw CSV
    # to that inbox (can be ignored — our processed zone reports go out separately).
    notify_email = config.UP_EMAIL
    log.info("Queuing report | timePeriod={} | notify={}".format(time_period, notify_email))
    payload = {
        "operationName": "generateReport",
        "variables": {
            "input": {
                "id":           config.REPORT_ID,
                "name":         config.REPORT_NAME,
                "exportFormat": config.EXPORT_FORMAT,
                "emails":       [notify_email],
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
    r = requests.post(config.ATLAS_GRAPHQL_URL, json=payload, headers=headers,
                      cookies=cookies, timeout=30)
    r.raise_for_status()
    resp_json = r.json()

    if resp_json.get("data") is None:
        log.error("generateReportV2 returned data=null. Full response: {}".format(r.text[:500]))
        errors = resp_json.get("errors", [])
        raise RuntimeError("generateReportV2 failed (data=null). errors={}".format(errors))

    gql = (resp_json.get("data") or {}).get("generateReportV2", {}).get("status", {})
    if not gql.get("success"):
        raise RuntimeError("generateReportV2 failed: {}".format(gql.get("messages")))

    log.info("Report queued. Polling every {}s (max {}s)...".format(poll_interval, max_wait))

    # Step 2: Snapshot existing job IDs so we only track the NEWLY queued report
    # (snapshot was already taken before queuing — reused here)
    log.info("Polling for new job (skipping {} pre-existing jobs)...".format(len(existing_ids)))

    # Step 3: Poll for the NEW report (ignore any pre-existing jobs)
    deadline  = time.time() + max_wait
    found_job = None

    while time.time() < deadline:
        time.sleep(poll_interval)
        tasks = _query_task_list(session)
        for task in tasks:
            jid = task.get("jobId", "")
            if jid in existing_ids:
                continue   # skip old reports
            # Match by timePeriod embedded in the task name, OR by report name alone
            # The task name from Atlas typically contains the timePeriod string
            task_name = task.get("name", "").lower()
            time_period_matches = time_period.lower().replace(",", "") in task_name.replace(" ", "").replace(",", "")
            report_name_matches = config.REPORT_NAME.lower() in task_name
            if report_name_matches or time_period_matches:
                found_job = task
                break

        if not found_job:
            log.info("New report not in task list yet -- waiting...")
            continue

        status = found_job.get("status", "")
        job_id = found_job.get("jobId", "")
        log.info("Status: {} | jobId: {}".format(status, job_id))

        if status == "completed":
            break
        elif status == "failed":
            raise RuntimeError("Atlas report FAILED: {}".format(found_job.get("message")))

        found_job = None   # still pending/processing, keep polling
    else:
        raise TimeoutError("Report did not complete within {} seconds.".format(max_wait))

    job_id       = found_job.get("jobId", "")
    download_url = found_job.get("downloadUrl") or REPORT_DOWNLOAD_BASE.format(job_id=job_id)
    log.info("Downloading from: {}".format(download_url))

    # Step 3: Download CSV
    dl = requests.get(download_url, cookies=cookies,
                      headers={"User-Agent": headers["User-Agent"],
                               "Authorization": headers.get("Authorization", "")},
                      timeout=120, stream=True)
    dl.raise_for_status()
    csv_bytes = dl.content
    log.info("Downloaded {:,} bytes.".format(len(csv_bytes)))
    return csv_bytes, job_id


# ==============================================================================
#  CORE EXECUTION
# ==============================================================================

def get_session_with_retry():
    session = load_cached_session()
    if session:
        return session
    log.info("No cached session -- fresh login...")
    return login_and_get_session()


# Backward-compat alias (run_automation.py imports get_cookies_with_retry)
def get_cookies_with_retry():
    return get_session_with_retry()


def delete_and_relogin():
    delete_session()
    return login_and_get_session()


def run_report(session, mode):
    time_period = get_time_period(mode)
    log.info("--- {} | timePeriod={} ---".format(mode.upper(), time_period))
    success = generate_report(session, time_period)
    if not success:
        log.warning("Failed -- re-logging in...")
        session = delete_and_relogin()
        success = generate_report(session, time_period)
    if not success:
        log.error("FAILED after re-login: {}".format(mode.upper()))
        sys.exit(1)
    return session


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

    if mode == "smart":
        modes_to_run = get_smart_modes()
    elif mode == "auto":
        modes_to_run = get_auto_modes()
        if not modes_to_run:
            log.info("Auto mode: today is neither Monday nor 1st — nothing to run.")
            return
    else:
        modes_to_run = [mode]
    log.info("Reports: {}".format([m.upper() for m in modes_to_run]))

    session = get_session_with_retry()
    for m in modes_to_run:
        session = run_report(session, m)

    log.info("=" * 60)
    log.info("All done.")


if __name__ == "__main__":
    main()