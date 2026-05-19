"""
Cancellation Processor
======================
Pure processing logic — no UI.
Categorizes cancellations, adds zones, builds per-zone Excel files.

Each Excel has 2 sheets:
  Sheet 1 - "Cancelled Orders"     : all cancelled rows across all zones
  Sheet 2 - "Cancellation Summary" : clean pivot + zone breakdown + kitchen stores
"""

import io
import os
import pandas as pd

_SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
_MAPPING_FILE = os.path.join(_SCRIPT_DIR, "keyword_mapping.xlsx")
_ZONE_FILE    = os.path.join(_SCRIPT_DIR, "zone_city_mapping.xlsx")

ZONES = ["North", "South", "West"]
ZONE_HEADER_COLORS = {"North": "#1F4E79", "South": "#375623", "West": "#7B2D2D"}

CATEGORIES = [
    "Auto Cancellation",
    "Kitchen Cancellation",
    "Non Kitchen Cancellation",
    "Post Dispatch Cancellation",
    "Cancelled post MFR",
]

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


def _add_formats(wb, zone=None):
    # Palette — light blue-grey theme, no black anywhere
    HDR_BG   = "#D9E1EA"   # light blue-grey header background
    HDR_FG   = "#2C4770"   # navy header text
    BODY_FG  = "#37474F"   # slate body text (not black)
    ALT_BG   = "#EEF4FB"   # very faint blue tint for alternating rows
    TOT_BG   = "#C5D5E4"   # slightly deeper blue for total/grand total rows
    TOT_FG   = "#2C4770"   # navy total text
    BORDER   = "#B0BEC5"   # soft blue-grey border
    TITLE_FG = "#2C4770"   # navy section title text

    base = {"font_name": "Arial", "font_size": 10, "valign": "vcenter",
            "border": 1, "border_color": BORDER, "font_color": BODY_FG}
    hdr  = {"font_name": "Arial", "font_size": 10, "bold": True, "valign": "vcenter",
            "border": 1, "border_color": BORDER,
            "bg_color": HDR_BG, "font_color": HDR_FG}
    tot  = {"font_name": "Arial", "font_size": 10, "bold": True, "valign": "vcenter",
            "border": 1, "border_color": BORDER,
            "bg_color": TOT_BG, "font_color": TOT_FG}
    alt  = dict(base, bg_color=ALT_BG)

    f = {}
    f["hdr"]   = wb.add_format({**hdr, "align": "center"})
    f["hdr_l"] = wb.add_format({**hdr, "align": "left"})

    f["body"]  = wb.add_format({**base, "align": "center"})
    f["body_l"]= wb.add_format({**base, "align": "left"})
    f["alt"]   = wb.add_format({**alt,  "align": "center"})
    f["alt_l"] = wb.add_format({**alt,  "align": "left"})

    f["title"] = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 12,
                                 "align": "left", "valign": "vcenter",
                                 "font_color": TITLE_FG, "bottom": 2, "bottom_color": HDR_BG})
    f["title2"]= wb.add_format({"bold": True, "font_name": "Arial", "font_size": 11,
                                 "align": "left", "valign": "vcenter",
                                 "font_color": TITLE_FG, "bottom": 2, "bottom_color": HDR_BG})
    f["title3"]= wb.add_format({"bold": True, "font_name": "Arial", "font_size": 11,
                                 "align": "left", "valign": "vcenter",
                                 "font_color": TITLE_FG, "bottom": 2, "bottom_color": HDR_BG})

    f["tot"]   = wb.add_format({**tot, "align": "center", "num_format": "#,##0"})
    f["tot_l"] = wb.add_format({**tot, "align": "left"})

    f["pctf"]  = wb.add_format({**base, "align": "center", "num_format": '0.00"%"'})
    f["pcta"]  = wb.add_format({**alt,  "align": "center", "num_format": '0.00"%"'})
    f["pctt"]  = wb.add_format({**tot,  "align": "center", "num_format": '0.00"%"'})

    f["num_b"] = wb.add_format({**base, "align": "center", "num_format": "#,##0"})
    f["num_a"] = wb.add_format({**alt,  "align": "center", "num_format": "#,##0"})
    f["num_t"] = wb.add_format({**tot,  "align": "center", "num_format": "#,##0"})
    return f


def build_zonal_excel(all_cancelled_df, pivot_df, zone_pivot_df, kitchen_df, zone):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    wb     = writer.book
    f      = _add_formats(wb, zone)

    # ── Sheet 1: Cancelled Orders ─────────────────────────────────────────────
    all_cancelled_df.to_excel(writer, sheet_name="Cancelled Orders", index=False)
    ws1 = writer.sheets["Cancelled Orders"]
    for c, col_name in enumerate(all_cancelled_df.columns):
        ws1.write(0, c, col_name, f["hdr"])
        w = max(len(str(col_name)), all_cancelled_df[col_name].astype(str).str.len().max())
        ws1.set_column(c, c, min(w + 2, 40))
    ws1.freeze_panes(1, 0)
    ws1.autofilter(0, 0, len(all_cancelled_df), len(all_cancelled_df.columns) - 1)

    # ── Sheet 2: Cancellation Summary ─────────────────────────────────────────
    ws2 = wb.add_worksheet("Cancellation Summary")
    r   = 0

    # Section 1: Category pivot
    ws2.merge_range(r, 0, r, 3, "  Cancellation Category Summary", f["title"])
    ws2.set_row(r, 28); r += 1
    for c, h in enumerate(["Cancellation Category", "Count", "% of Total Orders", "% of Total Cancelled"]):
        ws2.write(r, c, h, f["hdr"] if c > 0 else f["hdr_l"])
    ws2.set_row(r, 20); r += 1

    totals_set = {"Total Cancelled", "Total Orders"}
    for ri, prow in enumerate(pivot_df.itertuples(index=False)):
        cat, count, pct_ord, pct_can = prow
        is_tot = str(cat) in totals_set
        even   = (ri % 2 == 0)
        if is_tot:
            ws2.write(r, 0, cat,   f["tot_l"])
            ws2.write(r, 1, count, f["tot"])
            ws2.write(r, 2, float(pct_ord) if str(pct_ord) != "-" else "", f["pctt"])
            ws2.write(r, 3, float(pct_can) if str(pct_can) != "-" else "", f["pctt"])
        else:
            ws2.write(r, 0, cat,            f["alt_l"]  if even else f["body_l"])
            ws2.write(r, 1, count,          f["num_a"]  if even else f["num_b"])
            ws2.write(r, 2, float(pct_ord), f["pcta"]   if even else f["pctf"])
            pct_can_val = float(pct_can) if str(pct_can) != "-" else "-"
            ws2.write(r, 3, pct_can_val,    f["pcta"]   if even else f["pctf"])
        ws2.set_row(r, 18); r += 1

    r += 1

    # Section 2: Zone breakdown
    zone_hdrs = (["Cancellation Category"] +
                 [c for z in ZONES for c in [z, "% " + z]] +
                 ["Grand Total", "Overall %"])
    ws2.merge_range(r, 0, r, len(zone_hdrs) - 1, "  Zone-wise Cancellation Breakdown", f["title2"])
    ws2.set_row(r, 24); r += 1
    for c, h in enumerate(zone_hdrs):
        ws2.write(r, c, h, f["hdr"] if c > 0 else f["hdr_l"])
    ws2.set_row(r, 20); r += 1

    for ri, zrow in enumerate(zone_pivot_df.itertuples(index=False)):
        vals     = list(zrow)
        cat      = vals[0]
        is_grand = str(cat) == "Grand Total"
        even     = (ri % 2 == 0)
        ws2.write(r, 0, cat, f["tot_l"] if is_grand else (f["alt_l"] if even else f["body_l"]))
        col = 1; vi = 1
        for _ in ZONES:
            cnt_v = vals[vi]; vi += 1
            pct_v = float(vals[vi]); vi += 1
            if is_grand:
                ws2.write(r, col, cnt_v, f["num_t"]); ws2.write(r, col+1, pct_v, f["pctt"])
            else:
                ws2.write(r, col, cnt_v, f["num_a"] if even else f["num_b"])
                ws2.write(r, col+1, pct_v, f["pcta"] if even else f["pctf"])
            col += 2
        grand_cnt = vals[vi]; vi += 1; overall_pct = float(vals[vi])
        if is_grand:
            ws2.write(r, col, grand_cnt, f["num_t"]); ws2.write(r, col+1, overall_pct, f["pctt"])
        else:
            ws2.write(r, col, grand_cnt, f["num_a"] if even else f["num_b"])
            ws2.write(r, col+1, overall_pct, f["pcta"] if even else f["pctf"])
        ws2.set_row(r, 18); r += 1

    r += 1

    # Section 3: Kitchen by store
    ws2.merge_range(r, 0, r, 1,
        "  Kitchen Cancellation by Store - {} Zone".format(zone), f["title3"])
    ws2.set_row(r, 24); r += 1
    ws2.write(r, 0, "Store Name", f["hdr_l"])
    ws2.write(r, 1, "Count",      f["hdr"])
    ws2.set_row(r, 20); r += 1
    for ki, krow in enumerate(kitchen_df.itertuples(index=False)):
        store, count = krow
        even = (ki % 2 == 0)
        ws2.write(r, 0, store, f["alt_l"] if even else f["body_l"])
        ws2.write(r, 1, count, f["num_a"] if even else f["num_b"])
        ws2.set_row(r, 17); r += 1
    total_kitchen = int(kitchen_df["Count"].sum()) if not kitchen_df.empty else 0
    ws2.write(r, 0, "Grand Total", f["tot_l"])
    ws2.write(r, 1, total_kitchen, f["num_t"])

    ws2.set_column(0, 0, 34)
    ws2.set_column(1, 1, 12)
    ws2.set_column(2, 2, 14)
    ws2.set_column(3, 3, 14)
    ws2.set_column(4, 4, 12)
    ws2.set_column(5, 5, 14)
    ws2.set_column(6, 6, 12)
    ws2.set_column(7, 7, 14)
    ws2.set_column(8, 8, 12)
    ws2.freeze_panes(1, 0)

    writer.close()
    output.seek(0)
    return output


def process_csv(csv_bytes):
    print("Loading CSV data...")
    df = pd.read_csv(io.BytesIO(csv_bytes), low_memory=False)
    print("Loaded {} rows".format(len(df)))

    zone_df = load_zone_mapping()
    df = add_zone_city(df, zone_df)
    df = categorize_all(df)

    pivot_df, total_orders, total_cancelled = build_pivot(df)
    zone_pivot_df = build_zone_pivot(df)

    export_df = df.copy()
    for col in export_df.select_dtypes(include="datetime64[ns]").columns:
        export_df[col] = export_df[col].dt.strftime("%Y-%m-%d %H:%M:%S").fillna("")
    export_df = export_df.fillna("")

    # Label unmatched cancelled orders as "Other Cancellation" instead of dropping them
    export_df.loc[
        (export_df["Order Cancelled Time"] != "") &
        (export_df["Cancellation Category"] == ""),
        "Cancellation Category"
    ] = "Other Cancellation"

    all_cancelled_df = export_df[
        export_df["Order Cancelled Time"] != ""
    ].reset_index(drop=True)

    print("Cancelled orders: {:,} rows (Sheet 1 of every zone Excel)".format(len(all_cancelled_df)))

    zone_excels = {}
    for zone in ZONES:
        kitchen_df = build_kitchen_zone_breakdown(df, zone)
        zone_excels[zone] = build_zonal_excel(
            all_cancelled_df, pivot_df, zone_pivot_df, kitchen_df, zone
        )
        size_mb = len(zone_excels[zone].getvalue()) / (1024 * 1024)
        print("Built Excel for {} zone - {:.1f} MB".format(zone, size_mb))

    return zone_excels, pivot_df, zone_pivot_df, total_orders, total_cancelled
