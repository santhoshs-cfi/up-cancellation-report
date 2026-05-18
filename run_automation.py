"""
UrbanPiper Full Automation
==========================
Orchestrates the complete pipeline in one run:

  1. Login to UrbanPiper Atlas (headless Playwright)
  2. Generate the "Order Status Transitions" report for the correct time period
  3. Poll until Atlas has finished generating it
  4. Download the CSV
  5. Process it: categorize cancellations, add zones, build pivot tables
  6. Build per-zone Excel reports — 2 sheets each:
       Sheet 1: all cancelled orders (all zones, filter by Zone column)
       Sheet 2: category summary + zone breakdown + kitchen stores
  7. Email each zone's Excel (direct attachment if <= 25 MB, Drive link if larger)

Usage:
    python run_automation.py smart      # daily always; weekly on Mon; monthly on 1st
    python run_automation.py auto       # weekly on Mon; monthly on 1st (NO daily)
    python run_automation.py daily
    python run_automation.py weekly
    python run_automation.py monthly
    python run_automation.py test       # same as daily, for quick testing

Called by GitHub Actions (report.yml) — runs entirely headless.

Cron-job.org setup (2 jobs):
  Job 1: every day at 02:30 UTC (8:00 AM IST)  -> mode: daily
  Job 2: cron 30 3 1 * 1  (Mon + 1st of month) -> mode: auto
"""

import sys
import os
import logging
from datetime import date, timedelta

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("automation.log"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

try:
    import config
except ImportError:
    log.error("config.py not found.")
    sys.exit(1)

try:
    from urbanpiper_report import (
        get_time_period,
        get_smart_modes,
        get_auto_modes,
        get_session_with_retry,
        login_and_get_session,
        delete_session,
        generate_and_download_csv,
    )
except ImportError as e:
    log.error("Could not import from urbanpiper_report.py: {}".format(e))
    sys.exit(1)

try:
    from cancellation_processor import process_csv, build_kitchen_zone_breakdown
except ImportError as e:
    log.error("Could not import from cancellation_processor.py: {}".format(e))
    sys.exit(1)

try:
    from email_sender import send_all_zones
except ImportError as e:
    log.error("Could not import from email_sender.py: {}".format(e))
    sys.exit(1)


# ==============================================================================
#  DATE LABEL
# ==============================================================================

def make_date_label(mode):
    today = date.today()
    if mode in ("daily", "test"):
        return (today - timedelta(days=1)).strftime("%Y-%m-%d")
    elif mode == "weekly":
        end   = today - timedelta(days=1)
        start = today - timedelta(days=7)
        return "{} to {}".format(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))
    elif mode == "monthly":
        first_of_this = today.replace(day=1)
        last_of_prev  = first_of_this - timedelta(days=1)
        first_of_prev = last_of_prev.replace(day=1)
        return "{} to {}".format(
            first_of_prev.strftime("%Y-%m-%d"),
            last_of_prev.strftime("%Y-%m-%d"),
        )
    return today.strftime("%Y-%m-%d")


# ==============================================================================
#  PIPELINE
# ==============================================================================

def run_pipeline(session, mode):
    time_period = get_time_period(mode)
    date_label  = make_date_label(mode)

    log.info("=" * 60)
    log.info("Pipeline | mode={} | timePeriod={} | date={}".format(
        mode.upper(), time_period, date_label))

    # Step 1: Generate + download CSV
    try:
        csv_bytes, job_id = generate_and_download_csv(session, time_period)
    except Exception as e:
        log.warning("Download failed ({}). Re-logging in...".format(e))
        delete_session()
        session   = login_and_get_session()
        csv_bytes, job_id = generate_and_download_csv(session, time_period)

    log.info("CSV ready - {:,} bytes | jobId={}".format(len(csv_bytes), job_id))

    # Step 2: Process CSV -> zone Excel files
    log.info("Processing CSV...")
    zone_excels, pivot_df, zone_pivot_df, total_orders, total_cancelled = process_csv(csv_bytes)

    log.info("Processing complete - {:,} total orders, {:,} cancelled.".format(
        total_orders, total_cancelled))
    for zone, buf in zone_excels.items():
        size_mb = len(buf.getvalue()) / (1024 * 1024)
        log.info("  {}: {:.1f} MB Excel -> {}".format(
            zone, size_mb,
            "direct attachment" if size_mb < 25 else "Google Drive link"
        ))

    # Step 3: Build top-3 kitchen stores per zone (for email body)
    zone_kitchen_top3 = {}
    try:
        import io
        import pandas as pd
        from cancellation_processor import add_zone_city, categorize_all, load_zone_mapping
        df_raw  = pd.read_csv(io.BytesIO(csv_bytes), low_memory=False)
        zone_df = load_zone_mapping()
        df_raw  = add_zone_city(df_raw, zone_df)
        df_raw  = categorize_all(df_raw)
        for zone in ["North", "South", "West"]:
            kdf = build_kitchen_zone_breakdown(df_raw, zone)
            if not kdf.empty:
                zone_kitchen_top3[zone] = kdf.head(3).to_dict("records")
    except Exception as e:
        log.warning("Could not build kitchen top-3: {}".format(e))

    # Step 4: Send emails (pass pivot tables for inline body display)
    log.info("Sending emails...")
    results = send_all_zones(zone_excels, date_label, zone_kitchen_top3,
                             pivot_df=pivot_df, zone_pivot_df=zone_pivot_df, mode=mode)
    for zone, result in results.items():
        log.info("  {} -> {}".format(zone, result))

    log.info("Pipeline complete for {} mode.".format(mode.upper()))
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
        log.error("UP_EMAIL not configured.")
        sys.exit(1)
    if not config.UP_PASSWORD or config.UP_PASSWORD == "your-password-here":
        log.error("UP_PASSWORD not configured.")
        sys.exit(1)

    today = date.today()
    log.info("=" * 60)
    log.info("UrbanPiper Full Automation | Mode: {} | Date: {} ({})".format(
        mode.upper(), today.isoformat(), today.strftime("%A")))

    if mode == "smart":
        modes_to_run = get_smart_modes()
    elif mode == "auto":
        modes_to_run = get_auto_modes()
      