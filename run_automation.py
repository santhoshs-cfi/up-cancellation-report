"""
UrbanPiper Full Automation
==========================
Orchestrates the complete pipeline in one run:

  1. Login to UrbanPiper Atlas (headless Playwright)
  2. Generate the "Order Status Transitions" report for the correct time period
  3. Poll until Atlas has finished generating it
  4. Download the CSV
  5. Process it: categorize cancellations, add zones, build pivot tables
  6. Build per-zone CSV reports (all cancelled orders, filterable by Zone column)
  7. Email each zone's CSV to the configured recipient
     - Files <= 25 MB: attached directly
     - Files >  25 MB: uploaded to Google Drive, link sent in email

Usage (same modes as urbanpiper_report.py):
    python run_automation.py smart      # daily always; weekly on Mon; monthly on 1st
    python run_automation.py daily
    python run_automation.py weekly
    python run_automation.py monthly
    python run_automation.py test       # same as daily, for quick testing

Called by GitHub Actions (report.yml) — runs entirely headless, no local machine needed.
"""

import sys
import os
import logging
from datetime import date, timedelta

# ── Setup logging ──────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("automation.log"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# ── Local imports ──────────────────────────────────────────────────────────────
try:
    import config
except ImportError:
    log.error("config.py not found. Make sure it's in the same folder as this script.")
    sys.exit(1)

try:
    from urbanpiper_report import (
        get_time_period,
        get_smart_modes,
        get_session_with_retry,
        login_and_get_session,
        delete_session,
        generate_and_download_csv,
    )
except ImportError as e:
    log.error("Could not import from urbanpiper_report.py: {}".format(e))
    sys.exit(1)

try:
    from cancellation_processor import process_csv, build_zone_summaries
except ImportError as e:
    log.error("Could not import from cancellation_processor.py: {}".format(e))
    sys.exit(1)

try:
    from email_sender import send_all_zones
except ImportError as e:
    log.error("Could not import from email_sender.py: {}".format(e))
    sys.exit(1)


# ==============================================================================
#  DATE LABEL HELPERS
# ==============================================================================

def make_date_label(mode):
    """Human-readable date range label for email subject + body."""
    today = date.today()

    if mode in ("daily", "test"):
        yesterday = today - timedelta(days=1)
        return yesterday.strftime("%Y-%m-%d")

    elif mode == "weekly":
        end   = today - timedelta(days=1)          # yesterday
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
#  PIPELINE — ONE MODE (daily / weekly / monthly)
# ==============================================================================

def run_pipeline(session, mode):
    """
    Full end-to-end pipeline for one report mode.
    Returns the (possibly refreshed) session after execution.
    """
    time_period = get_time_period(mode)
    date_label  = make_date_label(mode)

    log.info("=" * 60)
    log.info("Pipeline | mode={} | timePeriod={} | date={}".format(
        mode.upper(), time_period, date_label
    ))

    # ── Step 1: Generate + download CSV ────────────────────────────────────────
    try:
        csv_bytes, job_id = generate_and_download_csv(session, time_period)
    except Exception as e:
        log.warning("Download failed ({}). Trying re-login...".format(e))
        delete_session()
        session   = login_and_get_session()
        csv_bytes, job_id = generate_and_download_csv(session, time_period)

    log.info("CSV ready — {:,} bytes | jobId={}".format(len(csv_bytes), job_id))

    # ── Step 2: Process CSV -> per-zone CSV files + pivots ─────────────────────
    log.info("Processing CSV...")
    zone_csvs, pivot_df, zone_pivot_df, total_orders, total_cancelled = process_csv(csv_bytes)

    log.info("Processing complete — {:,} total orders, {:,} cancelled.".format(
        total_orders, total_cancelled
    ))
    for zone, buf in zone_csvs.items():
        size_mb = len(buf.getvalue()) / (1024 * 1024)
        log.info("  {}: {:.1f} MB CSV ({} via {})".format(
            zone, size_mb,
            "attachment" if size_mb < 25 else "Google Drive link",
            "direct" if size_mb < 25 else "Drive upload",
        ))

    # ── Step 3: Build zone-specific summary text for email bodies ──────────────
    log.info("Building zone summaries...")
    try:
        import io
        import pandas as pd
        df_raw = pd.read_csv(io.BytesIO(csv_bytes), low_memory=False)
        from cancellation_processor import add_zone_city, categorize_all, load_zone_mapping
        zone_df = load_zone_mapping()
        df_raw  = add_zone_city(df_raw, zone_df)
        df_raw  = categorize_all(df_raw)
        zone_summaries = build_zone_summaries(df_raw, pivot_df, zone_pivot_df)
    except Exception as e:
        log.warning("Could not build zone summaries for email body: {}".format(e))
        zone_summaries = {}

    # ── Step 4: Email zone reports ─────────────────────────────────────────────
    log.info("Sending emails...")
    results = send_all_zones(zone_csvs, date_label, zone_summaries)

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

    # ── Validate credentials ───────────────────────────────────────────────────
    if not config.UP_EMAIL or config.UP_EMAIL == "your-email@company.com":
        log.error("UP_EMAIL not configured. Set it in config.py or as an env var.")
        sys.exit(1)
    if not config.UP_PASSWORD or config.UP_PASSWORD == "your-password-here":
        log.error("UP_PASSWORD not configured. Set it in config.py or as an env var.")
        sys.exit(1)

    today = date.today()
    log.info("=" * 60)
    log.info("UrbanPiper Full Automation | Mode: {} | Date: {} ({})".format(
        mode.upper(), today.isoformat(), today.strftime("%A")
    ))

    # ── Determine which modes to run ───────────────────────────────────────────
    if mode == "smart":
        modes_to_run = get_smart_modes()
    else:
        modes_to_run = [mode]

    log.info("Modes to run: {}".format([m.upper() for m in modes_to_run]))

    # ── Login once ─────────────────────────────────────────────────────────────
    session = get_session_with_retry()

    # ── Run pipeline for each mode ─────────────────────────────────────────────
    for m in modes_to_run:
        session = run_pipeline(session, m)

    log.info("=" * 60)
    log.info("All done. {} pipeline(s) completed.".format(len(modes_to_run)))


if __name__ == "__main__":
    main()
