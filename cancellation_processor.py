"""
Cancellation Processor
======================
Pure processing logic — no UI, no Excel.
Categorizes cancellations, adds zones, produces per-zone CSV bytes
and plain-text summaries for email bodies.

Used by run_automation.py for headless/scheduled execution.
"""

import io
import os
import pandas as pd

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
_SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
_MAPPING_FILE = os.path.join(_SCRIPT_DIR, "keyword_mapping.xlsx")
_ZONE_FILE    = os.path.join(_SCRIPT_DIR, "zone_city_mapping.xlsx")

ZONES = ["North", "South", "West"]

CATEGORIES = [
    "Auto Cancellation",
    "Kitchen Cancellation",
    "Non Kitchen Cancellation",
    "Post Dispatch Cancellation",
    "Cancelled post MFR",
]

# ---------------------------------------------------------------------------
# Default keywords (used if keyword_mapping.xlsx is missing)
# ---------------------------------------------------------------------------
_DEFAULT_PRIORITY_KITCHEN = [
    "could not confirm the order",
    "because your order was running late",
    "significant delay in your order",
]
_DEFAULT_KITCHEN = [
    "out of stock", "item_out_of_stock", "item out of stock",
    "items out of stock", "oos", "ordered items unavailable",
    "kitchen is full", "kitchen busy",
    "outlet closed", "outlet not open", "restaurant is now closed", "restaurant closed",
    "store_busy", "store busy", "store_closed", "store closed",
    "connectivity_issue", "connectivity issue", "merchant device", "electricity issue",
    "rx closed", "rx busy", "restaurant lpg", "lpg issue",
    "technical issue", "system issue",
    "merchant non-responsive", "merchant delaying", "merchant denying",
    "content or menu issue", "wrong merchant address",
    "reject-order got timeout",
]
_DEFAULT_NON_KITCHEN = [
    "customer cancellation", "customer_cancellation", "customer cancel",
    "customer placed the order by mistake", "customer reject due to",
    "swiggy order modification", "order modification",
    "merchant has accepted customer", "order cancelled by swiggy",
    "unable to see order on pos", "pos/ rms",
    "delivery partner cancellation", "rider_not_available", "rider not found", "no delivery partner",
    "vehicle breakdown", "fe - accident",
    "location issue", "wrong customer address",
    "identical orders", "duplicate order",
    "test order", "test_order", "test_order_aborted", "someone else picked",
    "rude behavior", "hygiene issues", "poor packaging", "wrong order delivered",
    "no cash", "order taking longer",
    "because your order was running late", "because there was a significant delay",
    "unspecified",
]
_DEFAULT_POST_DISPATCH = [
    "delivery partner cancellation",
    "rider not available",
    "no delivery partner",
]
_DELAY_CONTEXT_MESSAGES = [
    "reject-order rejected due to customer cancellation",
    "reject-order rejected due to 'customer cancellation'",
    "order cancelled by swiggy",
    "reject-order rejected due to customer cancellation due to delay",
    "reject-order rejected due to 'customer cancellation due to delay'",
    "customer reject due to delay",
]


# ---------------------------------------------------------------------------
# Keyword loading
# ---------------------------------------------------------------------------
def load_keywords():
    if os.path.exists(_MAPPING_FILE):
        try:
            km = pd.read_excel(_MAPPING_FILE)
            km.columns = [c.strip() for c in km.columns]
            km["Keyword"]  = km["Keyword"].astype(str).str.strip().str.lower()
            km["Category"] = km["Category"].astype(str).str.strip()
            priority     = km.loc[km["Category"] == "Priority Kitchen",          "Keyword"].tolist()
            kitchen      = km.loc[km["Category"] == "Kitchen Cancellation",      "Keyword"].tolist()
            nkitchen     = km.loc[km["Category"] == "Non Kitchen Cancellation",  "Keyword"].tolist()
            postdispatch = km.loc[km["Category"] == "Post Dispatch Cancellation","Keyword"].tolist()
            print("Keywords loaded: Priority={} Kitchen={} NonKitchen={} PostDispatch={}".format(
                len(priority), len(kitchen), len(nkitchen), len(postdispatch)))
            return (
                priority     or _DEFAULT_PRIORITY_KITCHEN,
                kitchen      or _DEFAULT_KITCHEN,
                nkitchen     or _DEFAULT_NON_KITCHEN,
                postdispatch or _DEFAULT_POST_DISPATCH,
            )
        except Exception as e:
            print("Warning: could not load keyword_mapping.xlsx ({}). Using defaults.".format(e))
    return _DEFAULT_PRIORITY_KITCHEN, _DEFAULT_KITCHEN, _DEFAULT_NON_KITCHEN, _DEFAULT_POST_DISPATCH


# ---------------------------------------------------------------------------
# Zone mapping
# ---------------------------------------------------------------------------
def load_zone_mapping():
    if os.path.exists(_ZONE_FILE):
        try:
            zm = pd.read_excel(_ZONE_FILE)
            zm.columns = [c.strip() for c in zm.columns]
            zm["Store Name"] = zm["Store Name"].astype(str).str.strip()
            zm["City"]       = zm["City"].astype(str).str.strip()
            zm["Zone"]       = zm["Zone"].astype(str).str.strip()
            print("Zone mapping loaded: {} stores".format(len(zm)))
            return zm[["Store Name", "City", "Zone"]]
        except Exception as e:
            print("Warning: could not load zone_city_mapping.xlsx ({})".format(e))
    return pd.DataFrame(columns=["Store Name", "City", "Zone"])


def add_zone_city(df, zone_df):
    df = df.copy()
    if zone_df.empty or "Store Name" not in df.columns:
        df["City"] = "Unknown"
        df["Zone"] = "Unknown"
        return df
    merged = df.merge(zone_df, on="Store Name", how="left")
    merged["City"] = merged["City"].fillna("Unknown")
    merged["Zone"] = merged["Zone"].fillna("Unknown")
    cols = [c for c in merged.columns if c not in ("City", "Zone")]
    if "Store Name" in cols:
        idx = cols.index("Store Name")
        cols.insert(idx + 1, "City")
        cols.insert(idx + 2, "Zone")
    else:
        cols = cols + ["City", "Zone"]
    return merged[cols]


# ---------------------------------------------------------------------------
# Categorization
# ---------------------------------------------------------------------------
def categorize_row(row, priority_kw, non_kitchen_kw, kitchen_kw, post_dispatch_kw):
    if pd.isna(row["Order Cancelled Time"]):
        return ""

    cancel_reason = str(row.get("Order Cancel Reason", "") or "").strip()
    if cancel_reason == "14":
        return "Non Kitchen Cancellation"

    if not pd.isna(row["Dispatch Time"]):
        return "Post Dispatch Cancellation"

    if (not pd.isna(row.get("MFR Time")) and
            pd.isna(row["Dispatch Time"]) and
            str(row.get("Order Cancelled By", "") or "").strip().lower() == "aggregator"):
        return "Cancelled post MFR"

    if str(row.get("Order Cancelled By", "") or "").strip().lower() == "merchant":
        raw = row.get("Order Cancellation Message", "")
        msg = "" if (pd.isna(raw) or str(raw).strip().lower() in ("", "none", "nan")) else str(raw).strip()
        if msg and msg.lower().startswith("reason:"):
            return "Kitchen Cancellation"
        if str(row.get("Order Cancel Reason", "") or "").strip().lower() == "item_out_of_stock":
            return "Kitchen Cancellation" if msg else "Auto Cancellation"

    text = (
        str(row.get("Order Cancellation Message", "") or "") + " " +
        str(row.get("Order Cancel Reason", "") or "")
    ).lower()

    if any(k in text for k in post_dispatch_kw):
        return "Post Dispatch Cancellation"

    if any(k in text for k in _DELAY_CONTEXT_MESSAGES):
        if not pd.isna(row["Order Creation Time"]):
            mins = (row["Order Cancelled Time"] - row["Order Creation Time"]).total_seconds() / 60
            if mins <= 10:
                return "Non Kitchen Cancellation"
            return "Post Dispatch Cancellation" if not pd.isna(row.get("MFR Time")) else "Kitchen Cancellation"
        return "Non Kitchen Cancellation"

    if any(k in text for k in priority_kw):
        return "Kitchen Cancellation"
    if any(k in text for k in non_kitchen_kw):
        return "Non Kitchen Cancellation"
    if any(k in text for k in kitchen_kw):
        return "Kitchen Cancellation"

    if not pd.isna(row["Order Creation Time"]):
        mins = (row["Order Cancelled Time"] - row["Order Creation Time"]).total_seconds() / 60
        if mins <= 8:
            return "Auto Cancellation"

    return ""


def categorize_all(df):
    priority_kw, kitchen_kw, non_kitchen_kw, post_dispatch_kw = load_keywords()
    print("Categorizing {} rows...".format(len(df)))

    for col in ["Order Creation Time", "MFR Time", "Dispatch Time", "Order Cancelled Time"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    df["Cancellation Category"] = df.apply(
        lambda row: categorize_row(row, priority_kw, non_kitchen_kw, kitchen_kw, post_dispatch_kw),
        axis=1
    )
    counts = df["Cancellation Category"].replace("", "(Not cancelled)").value_counts()
    print("Category breakdown:\n{}".format(counts.to_string()))
    return df


# ---------------------------------------------------------------------------
# Pivot builders
# ---------------------------------------------------------------------------
def build_pivot(df):
    total_orders    = len(df)
    has_cancel_time = df["Order Cancelled Time"].notna()
    total_cancelled = int(has_cancel_time.sum())
    unmatched       = int((has_cancel_time & (df["Cancellation Category"] == "")).sum())

    rows = []
    for cat in CATEGORIES:
        count = int((df["Cancellation Category"] == cat).sum())
        rows.append({
            "Cancellation Category": cat,
            "Count":                 count,
            "% of Total Orders":     round(count / total_orders    * 100, 2) if total_orders    > 0 else 0,
            "% of Total Cancelled":  round(count / total_cancelled * 100, 2) if total_cancelled > 0 else 0,
        })
    rows.append({"Cancellation Category": "Unmatched (Blank)", "Count": unmatched,
                 "% of Total Orders":    round(unmatched / total_orders    * 100, 2) if total_orders    > 0 else 0,
                 "% of Total Cancelled": round(unmatched / total_cancelled * 100, 2) if total_cancelled > 0 else 0})
    rows.append({"Cancellation Category": "Total Cancelled", "Count": total_cancelled,
                 "% of Total Orders":    round(total_cancelled / total_orders * 100, 2) if total_orders > 0 else 0,
                 "% of Total Cancelled": 100.0})
    rows.append({"Cancellation Category": "Total Orders", "Count": int(total_orders),
                 "% of Total Orders": 100.0, "% of Total Cancelled": "-"})
    return pd.DataFrame(rows), total_orders, total_cancelled


def build_zone_pivot(df):
    totals_by_zone = {z: max(len(df[df["Zone"] == z]), 1) for z in ZONES}
    total_all = max(len(df), 1)
    rows = []
    for cat in CATEGORIES:
        row = {"Cancellation Category": cat}
        cat_grand = 0
        for z in ZONES:
            count = int((df[df["Zone"] == z]["Cancellation Category"] == cat).sum())
            row[z]        = count
            row["% " + z] = round(count / totals_by_zone[z] * 100, 2)
            cat_grand    += count
        row["Grand Total"] = cat_grand
        row["Overall %"]   = round(cat_grand / total_all * 100, 2)
        rows.append(row)

    grand = {"Cancellation Category": "Grand Total"}
    for z in ZONES:
        z_total     = len(df[df["Zone"] == z])
        z_cancelled = int((df[df["Zone"] == z]["Cancellation Category"] != "").sum())
        grand[z]        = z_total
        grand["% " + z] = round(z_cancelled / max(z_total, 1) * 100, 2)
    grand["Grand Total"] = len(df)
    grand["Overall %"]   = round(int((df["Cancellation Category"] != "").sum()) / total_all * 100, 2)
    rows.append(grand)

    col_order = ["Cancellation Category"] + [c for z in ZONES for c in [z, "% " + z]] + ["Grand Total", "Overall %"]
    return pd.DataFrame(rows, columns=col_order)


def build_kitchen_zone_breakdown(df, zone):
    mask = (df["Zone"] == zone) & (df["Cancellation Category"] == "Kitchen Cancellation")
    if "Store Name" not in df.columns or mask.sum() == 0:
        return pd.DataFrame(columns=["Store Name", "Count"])
    breakdown = df[mask].groupby("Store Name").size().reset_index(name="Count")
    return breakdown.sort_values("Count", ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Plain-text summary builder (goes into email body)
# ---------------------------------------------------------------------------
def build_zone_summary_text(pivot_df, zone_pivot_df, kitchen_df, zone):
    """Returns a formatted plain-text summary block for one zone's email."""
    lines = []

    # ── Overall category summary ──────────────────────────────────────────────
    lines.append("=" * 56)
    lines.append("  CANCELLATION SUMMARY")
    lines.append("=" * 56)
    lines.append("{:<30} {:>7} {:>10} {:>12}".format(
        "Category", "Count", "% Orders", "% Cancelled"))
    lines.append("-" * 56)
    for _, row in pivot_df.iterrows():
        cat   = str(row["Cancellation Category"])
        count = row["Count"]
        p_ord = row["% of Total Orders"]
        p_can = row["% of Total Cancelled"]
        pct_o = "{:.2f}%".format(float(p_ord)) if str(p_ord) != "-" else "-"
        pct_c = "{:.2f}%".format(float(p_can)) if str(p_can) != "-" else "-"
        lines.append("{:<30} {:>7} {:>10} {:>12}".format(cat, count, pct_o, pct_c))
    lines.append("")

    # ── Zone breakdown row for this zone ─────────────────────────────────────
    lines.append("=" * 56)
    lines.append("  {} ZONE BREAKDOWN".format(zone.upper()))
    lines.append("=" * 56)
    lines.append("{:<30} {:>7} {:>10}".format("Category", "Count", "% of Zone"))
    lines.append("-" * 56)
    zone_total = zone_pivot_df.loc[
        zone_pivot_df["Cancellation Category"] == "Grand Total", zone
    ].values
    zone_total_val = int(zone_total[0]) if len(zone_total) > 0 else 1
    for _, row in zone_pivot_df.iterrows():
        cat = str(row["Cancellation Category"])
        if cat == "Grand Total":
            continue
        count = int(row.get(zone, 0))
        pct   = row.get("% " + zone, 0)
        lines.append("{:<30} {:>7} {:>10}".format(cat, count, "{:.2f}%".format(float(pct))))
    lines.append("-" * 56)
    lines.append("{:<30} {:>7}".format("Total Zone Orders", zone_total_val))
    lines.append("")

    # ── Top kitchen cancellation stores ──────────────────────────────────────
    if not kitchen_df.empty:
        lines.append("=" * 56)
        lines.append("  TOP KITCHEN CANCELLATION STORES ({})".format(zone.upper()))
        lines.append("=" * 56)
        for i, row in kitchen_df.head(10).iterrows():
            lines.append("  {:>2}. {:<40} {:>5}".format(
                i + 1, str(row["Store Name"])[:40], int(row["Count"])))
        lines.append("")

    return "\n".join(lines)


def build_zone_summaries(df, pivot_df, zone_pivot_df):
    """Returns dict {zone: summary_text_string} for all zones."""
    summaries = {}
    for zone in ZONES:
        kitchen_df = build_kitchen_zone_breakdown(df, zone)
        summaries[zone] = build_zone_summary_text(pivot_df, zone_pivot_df, kitchen_df, zone)
    return summaries


# ---------------------------------------------------------------------------
# CSV builder
# ---------------------------------------------------------------------------
def build_zone_csv(all_cancelled_df):
    """Convert the all-cancelled DataFrame to CSV bytes (BytesIO)."""
    buf = io.BytesIO()
    all_cancelled_df.to_csv(buf, index=False, encoding="utf-8-sig")  # utf-8-sig for Excel-safe open
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Main entry point: raw CSV bytes -> dict of {zone: BytesIO CSV}
# ---------------------------------------------------------------------------
def process_csv(csv_bytes):
    """
    Takes raw CSV bytes from UrbanPiper.
    Returns:
        zone_csvs       - dict {zone: BytesIO}  — same CSV for every zone
                          (all cancelled orders, all zones, filterable by Zone column)
        pivot_df        - overall cancellation category pivot
        zone_pivot_df   - zone-wise breakdown pivot
        total_orders    - int
        total_cancelled - int
    """
    print("Loading CSV data...")
    df = pd.read_csv(io.BytesIO(csv_bytes), low_memory=False)
    print("Loaded {} rows".format(len(df)))

    # Add zone / city columns
    zone_df = load_zone_mapping()
    df = add_zone_city(df, zone_df)

    # Categorize all rows
    df = categorize_all(df)

    # Build pivot tables
    pivot_df, total_orders, total_cancelled = build_pivot(df)
    zone_pivot_df = build_zone_pivot(df)

    # Prepare export df — convert datetimes to strings, fill NaN
    export_df = df.copy()
    for col in export_df.select_dtypes(include="datetime64[ns]").columns:
        export_df[col] = export_df[col].dt.strftime("%Y-%m-%d %H:%M:%S").fillna("")
    export_df = export_df.fillna("")

    # All cancelled orders (Cancellation Category != "") — all zones combined
    all_cancelled_df = export_df[
        export_df["Cancellation Category"] != ""
    ].reset_index(drop=True)

    print("Cancelled orders: {:,} rows (out of {:,} total)".format(
        len(all_cancelled_df), len(export_df)))

    # Build one CSV — same content sent to every zone (Zone column lets them filter)
    csv_buf = build_zone_csv(all_cancelled_df)
    csv_bytes_value = csv_buf.getvalue()  # bytes, reusable

    # Return one BytesIO per zone (each reads from same underlying bytes)
    zone_csvs = {
        zone: io.BytesIO(csv_bytes_value) for zone in ZONES
    }

    return zone_csvs, pivot_df, zone_pivot_df, total_orders, total_cancelled
