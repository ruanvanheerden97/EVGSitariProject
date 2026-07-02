import streamlit as st
import pandas as pd
import calendar as cal
from datetime import datetime, date
import glob
import os
import json
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Sitari Evergreen — Meter Commissioning",
    page_icon="🔧",
    layout="wide",
)

# ---------- Styling ----------
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; max-width: 1300px;}
.stMetric, div[data-testid="stMetric"] {
  background: #FBF9F3 !important;
  border: 1px solid #DCD6C4;
  border-radius: 10px;
  padding: 10px 14px;
}
div[data-testid="stMetric"] * {
  color: #152B45 !important;
}
div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] * {
  font-size: 12px !important;
  text-transform: uppercase;
  letter-spacing: .05em;
  color: #3E5066 !important;
}
div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] * {
  color: #152B45 !important;
}
div[data-testid="stMetricDelta"], div[data-testid="stMetricDelta"] * {
  color: #3F7D5C !important;
  fill: #3F7D5C !important;
}
h1, h2, h3 {color: #152B45;}
.badge {display:inline-block; padding:2px 9px; border-radius:14px; font-size:11px; font-weight:600; font-family: monospace;}
.badge-overdue {background:#F5E2DB; color:#BD4B2C;}
.badge-upcoming {background:#FCEFDD; color:#B96E1E;}
.badge-ontrack {background:#E7EEF5; color:#1F3F66;}
.badge-installed {background:#E3EEE6; color:#3F7D5C;}

/* Calendar grid */
.cal-table {width:100%; border-collapse: collapse; table-layout: fixed;}
.cal-table th {font-size:11px; text-transform:uppercase; color:#3E5066; padding:6px; text-align:center; border-bottom:2px solid #152B45;}
.cal-table td {vertical-align:top; border:1px solid #E3DECB; height:92px; padding:5px; width:14.28%;}
.cal-day-num {font-family:monospace; font-size:12px; color:#3E5066; font-weight:600;}
.cal-day-num.today {color:#fff; background:#152B45; border-radius:4px; padding:0 5px;}
.cal-pill {display:block; font-size:10px; font-family:monospace; border-radius:4px; padding:1px 4px; margin-top:3px; line-height:1.5;}
.cal-pill.installed {background:#E3EEE6; color:#3F7D5C;}
.cal-pill.upcoming {background:#FCEFDD; color:#B96E1E;}
.cal-pill.overdue {background:#F5E2DB; color:#BD4B2C;}
.cal-empty {background:#FAFAF6;}
</style>
""", unsafe_allow_html=True)

WATER_SHEET_CANDIDATES = ["Water meters", "Water Meters"]
ELEC_SHEET_CANDIDATES = ["Elec Meters", "Electrical Meters"]
FILE_PATTERN = "EVG_SIT_FS_Meter_Commissioning_*.xlsx"

# ---------- Helpers ----------
def find_sheet(xls, candidates):
    for name in xls.sheet_names:
        if name.strip().lower() in [c.lower() for c in candidates]:
            return name
    return None

def coalesce_col(df, options):
    for o in options:
        if o in df.columns:
            return df[o]
    return pd.Series([None] * len(df))

def find_data_file():
    """The spreadsheet lives in this same repo folder — push an updated copy
    whenever it changes, and this picks up the most recently modified match."""
    matches = glob.glob(os.path.join(os.path.dirname(__file__), FILE_PATTERN))
    if not matches:
        return None
    return max(matches, key=os.path.getmtime)

@st.cache_data(show_spinner=False)
def load_data(file_path, _mtime):
    xls = pd.ExcelFile(file_path)
    water_name = find_sheet(xls, WATER_SHEET_CANDIDATES)
    elec_name = find_sheet(xls, ELEC_SHEET_CANDIDATES)

    records = []

    def fmt_serial(col):
        """Convert float serials like 14558067900.0 → '14558067900', blanking NaN."""
        def _fmt(v):
            if pd.isna(v):
                return ""
            try:
                return str(int(float(v)))
            except (ValueError, OverflowError):
                return str(v).strip()
        return col.apply(_fmt)

    if water_name:
        wdf = xls.parse(water_name)
        wdf.columns = [str(c).strip() for c in wdf.columns]
        out = pd.DataFrame()
        out["stand"] = coalesce_col(wdf, ["Stand Number"]).astype(str).str.strip()
        out["unit_type"] = coalesce_col(wdf, ["Section"])
        out["wbho_section"] = coalesce_col(wdf, ["WBHO Subsection"])
        out["manufacturer"] = coalesce_col(wdf, ["Manufacturer"])
        out["model"] = coalesce_col(wdf, ["Meter Model"])
        out["serial"] = fmt_serial(coalesce_col(wdf, ["Meter serial number"]))
        out["commission_date"] = pd.to_datetime(coalesce_col(wdf, ["Meter Commissioning Date", "Meter Commission Date"]), errors="coerce")
        out["amr"] = coalesce_col(wdf, ["AMR Commissioned"]).fillna(False).astype(bool)
        out["deadline"] = pd.to_datetime(coalesce_col(wdf, ["Snag Date 4"]), errors="coerce")
        out["meter_type"] = "Water"
        out = out[out["stand"].notna() & (out["stand"] != "") & (out["stand"].str.lower() != "none")]
        records.append(out)

    if elec_name:
        edf = xls.parse(elec_name)
        edf.columns = [str(c).strip() for c in edf.columns]
        out = pd.DataFrame()
        out["stand"] = coalesce_col(edf, ["Stand Number"]).astype(str).str.strip()
        out["unit_type"] = coalesce_col(edf, ["Section"])
        out["wbho_section"] = coalesce_col(edf, ["WBHO Subsection"])
        out["manufacturer"] = coalesce_col(edf, ["Manufacturer"])
        out["model"] = coalesce_col(edf, ["Meter Model"])
        out["serial"] = fmt_serial(coalesce_col(edf, ["Meter Serial"]))
        out["commission_date"] = pd.to_datetime(coalesce_col(edf, ["Meter Commission Date", "Meter Commissioning Date"]), errors="coerce")
        out["amr"] = coalesce_col(edf, ["AMR Installed"]).fillna(False).astype(bool)
        out["deadline"] = pd.to_datetime(coalesce_col(edf, ["Snag Date 4"]), errors="coerce")
        out["meter_type"] = "Electrical"
        out = out[out["stand"].notna() & (out["stand"] != "") & (out["stand"].str.lower() != "none")]
        records.append(out)

    if not records:
        return pd.DataFrame()

    df = pd.concat(records, ignore_index=True)
    df["installed"] = df["commission_date"].notna()

    today = pd.Timestamp(date.today())
    df["days_to_deadline"] = (df["deadline"] - today).dt.days

    def status_row(r):
        if r["installed"]:
            if pd.notna(r["deadline"]) and r["commission_date"] > r["deadline"]:
                return "Installed late"
            return "Installed"
        if pd.notna(r["deadline"]) and r["deadline"] < today:
            return "Overdue"
        if pd.notna(r["deadline"]) and r["days_to_deadline"] <= 14:
            return "Due soon"
        return "On track"

    df["status"] = df.apply(status_row, axis=1)
    return df


@st.cache_data(show_spinner=False)
def load_kiosk_data(file_path, _mtime):
    """Build the minisub → kiosk → meters hierarchy for the reticulation diagram."""
    xls = pd.ExcelFile(file_path)

    # --- Kiosk Plan (planned counts per kiosk) ---
    kp = xls.parse("Kiosk Plan")
    kp.columns = [str(c).strip() for c in kp.columns]
    kp = kp[kp["Kiosk Number"].notna() & kp["Minisub"].notna()].copy()
    kp["Minisub"] = kp["Minisub"].astype(int)
    kp["MS Serial"] = kp["MS Serial"].astype(int).astype(str)
    kp["planned"] = pd.to_numeric(kp["New planned units"], errors="coerce").fillna(0).astype(int)

    # --- Elec Meters (installed counts per kiosk, stand list, AMR) ---
    elec_name = find_sheet(xls, ["Elec Meters", "Electrical Meters"])
    edf = xls.parse(elec_name)
    edf.columns = [str(c).strip() for c in edf.columns]
    edf = edf[edf["Kiosk Number"].notna()].copy()
    edf["installed"] = edf["Meter Commission Date"].notna()
    edf["amr_done"] = edf["AMR Installed"].fillna(False).astype(bool)
    edf["stand_str"] = edf["Stand Number"].astype(str).str.strip()

    def _fmt_serial(v):
        if pd.isna(v): return ""
        try: return str(int(float(v)))
        except: return str(v).strip()

    edf["serial_str"] = edf["Meter Serial"].apply(_fmt_serial)

    # stand → {installed, amr, serial}
    stand_detail = {
        row["stand_str"]: {
            "installed": bool(row["installed"]),
            "amr": bool(row["amr_done"]),
            "serial": row["serial_str"],
        }
        for _, row in edf.iterrows()
    }

    # Aggregate per kiosk
    def agg_stands(x):
        return sorted(x.tolist())

    def agg_installed_stands(x):
        return sorted(edf.loc[x.index][edf.loc[x.index, "installed"]]["stand_str"].tolist())

    def agg_amr_stands(x):
        return sorted(edf.loc[x.index][edf.loc[x.index, "amr_done"]]["stand_str"].tolist())

    def agg_stand_serials(x):
        """Return dict of stand → serial for stands that have a serial."""
        subset = edf.loc[x.index][edf.loc[x.index, "serial_str"] != ""]
        return {row["stand_str"]: row["serial_str"] for _, row in subset.iterrows()}

    kiosk_agg = edf.groupby("Kiosk Number").agg(
        installed_count=("installed", "sum"),
        amr_count=("amr_done", "sum"),
        total_count=("Stand Number", "count"),
        stands=("stand_str", agg_stands),
        installed_stands=("stand_str", agg_installed_stands),
        amr_stands=("stand_str", agg_amr_stands),
        stand_serials=("stand_str", agg_stand_serials),
    ).reset_index()
    kiosk_agg.columns = ["kiosk", "installed", "amr_count", "total", "stands", "installed_stands", "amr_stands", "stand_serials"]

    # Merge plan vs actuals
    merged = kp.merge(kiosk_agg, left_on="Kiosk Number", right_on="kiosk", how="left")
    merged["installed"] = merged["installed"].fillna(0).astype(int)
    merged["amr_count"] = merged["amr_count"].fillna(0).astype(int)
    merged["total"] = merged["total"].fillna(0).astype(int)
    merged["planned"] = merged["planned"].astype(int)
    merged["stands"] = merged["stands"].apply(lambda x: x if isinstance(x, list) else [])
    merged["installed_stands"] = merged["installed_stands"].apply(lambda x: x if isinstance(x, list) else [])
    merged["amr_stands"] = merged["amr_stands"].apply(lambda x: x if isinstance(x, list) else [])
    merged["stand_serials"] = merged["stand_serials"].apply(lambda x: x if isinstance(x, dict) else {})

    # Build hierarchy: minisub → list of kiosks
    def sort_kiosk_key(k):
        import re
        m = re.match(r"(\d+)EVE(\d+)", str(k))
        return (int(m.group(2)), int(m.group(1))) if m else (99, 99)

    hierarchy = {}
    for _, row in merged.iterrows():
        ms_id = int(row["Minisub"])
        ms_serial = str(row["MS Serial"])
        if ms_id not in hierarchy:
            hierarchy[ms_id] = {"serial": ms_serial, "kiosks": []}
        hierarchy[ms_id]["kiosks"].append({
            "kiosk": row["Kiosk Number"],
            "planned": int(row["planned"]),
            "installed": int(row["installed"]),
            "amr_count": int(row["amr_count"]),
            "total_in_sheet": int(row["total"]),
            "stands": row["stands"],
            "installed_stands": row["installed_stands"],
            "amr_stands": row["amr_stands"],
            "stand_serials": row["stand_serials"],
            "comment": str(row.get("Comments", "") or ""),
        })

    # Sort kiosks within each minisub
    for ms_id in hierarchy:
        hierarchy[ms_id]["kiosks"].sort(key=lambda k: sort_kiosk_key(k["kiosk"]))

    return hierarchy


def summary_counters(view_df, full_df, label):
    """Shows total counts for this category, plus a water/electrical split,
    and how many remain after the filters above are applied."""
    total_in_category = len(full_df)
    water_n = int((full_df["meter_type"] == "Water").sum())
    elec_n = int((full_df["meter_type"] == "Electrical").sum())
    shown_n = len(view_df)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"Total {label}", total_in_category)
    c2.metric("Water", water_n)
    c3.metric("Electrical", elec_n)
    c4.metric("Matching current filters", shown_n)


def status_filters(df, key_prefix):
    c1, c2, c3 = st.columns(3)
    with c1:
        meter_types = st.multiselect("Meter type", sorted(df["meter_type"].unique()), default=list(df["meter_type"].unique()), key=f"{key_prefix}_type")
    with c2:
        sections = st.multiselect("Section (WBHO)", sorted(df["wbho_section"].dropna().unique()), key=f"{key_prefix}_section")
    with c3:
        search_stand = st.text_input("Search stand number", key=f"{key_prefix}_search")

    out = df[df["meter_type"].isin(meter_types)]
    if sections:
        out = out[out["wbho_section"].isin(sections)]
    if search_stand:
        out = out[out["stand"].str.contains(search_stand, case=False, na=False)]
    return out


def show_table(view_df, columns, rename, sort_col=None, ascending=True):
    if view_df.empty:
        st.info("Nothing here for the current filters.")
        return
    show = view_df[columns].rename(columns=rename)
    if sort_col:
        show = show.sort_values(rename.get(sort_col, sort_col), ascending=ascending)
    for c in show.columns:
        if pd.api.types.is_datetime64_any_dtype(show[c]):
            show[c] = show[c].dt.strftime("%d %b %Y")
    st.dataframe(show, use_container_width=True, hide_index=True)
    csv = show.to_csv(index=False).encode("utf-8")
    st.download_button("Download as CSV", csv, file_name="meters_export.csv", mime="text/csv", key=f"dl_{columns[0]}_{len(view_df)}")


# ---------- Load data ----------
st.title("🔧 Sitari Evergreen — Meter Commissioning")
st.caption("Erf 1186 Sitari · Lifestyle Retirement Village")

data_path = find_data_file()
if not data_path:
    st.error(f"No file matching `{FILE_PATTERN}` found in this app's folder. Push the latest spreadsheet to the repo to continue.")
    st.stop()

mtime = os.path.getmtime(data_path)
df = load_data(data_path, mtime)

if df.empty:
    st.error("Couldn't find a 'Water meters' or 'Elec Meters' tab in this file. Check the sheet names.")
    st.stop()

st.caption(f"📂 Using **{os.path.basename(data_path)}** · last updated {datetime.fromtimestamp(mtime).strftime('%d %b %Y, %H:%M')}")

# ---------- KPI strip (always visible) ----------
total = len(df)
installed_n = int(df["installed"].sum())
overdue_n = int((df["status"] == "Overdue").sum())
due_soon_n = int((df["status"] == "Due soon").sum())
outstanding_n = total - installed_n

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total meter points", total)
c2.metric("Installed", installed_n, f"{round(installed_n/total*100)}% complete")
c3.metric("Outstanding", outstanding_n)
c4.metric("Due within 14 days", due_soon_n)
c5.metric("Overdue", overdue_n, delta_color="inverse")

st.divider()

# ---------- Tabs ----------
tab_outstanding, tab_upcoming, tab_overdue, tab_installed, tab_calendar, tab_sections, tab_retic = st.tabs(
    ["🟦 Outstanding", "🟧 Upcoming", "🟥 Overdue", "🟩 Installed", "📅 Calendar", "📊 Sections", "⚡ Reticulation"]
)

COLS = ["stand", "meter_type", "unit_type", "wbho_section", "deadline", "status"]
RENAME = {"stand": "Stand", "meter_type": "Type", "unit_type": "Unit type", "wbho_section": "Section", "deadline": "Deadline (Snag 4)", "status": "Status"}

with tab_outstanding:
    st.subheader("All outstanding meters")
    outstanding_full = df[~df["installed"]]
    filtered = status_filters(outstanding_full, "outstanding")
    summary_counters(filtered, outstanding_full, "outstanding")
    show_table(filtered, COLS, RENAME, sort_col="deadline")

with tab_upcoming:
    st.subheader("Due within the next 14 days")
    due_soon_full = df[df["status"] == "Due soon"]
    filtered = status_filters(due_soon_full, "upcoming")
    summary_counters(filtered, due_soon_full, "due soon")
    show_table(filtered, COLS, RENAME, sort_col="deadline")
    st.markdown("##### Further ahead")
    further = status_filters(df[(~df["installed"]) & (df["deadline"].notna()) & (df["deadline"] >= pd.Timestamp(date.today())) & (df["status"] != "Due soon")], "further")
    if not further.empty:
        further = further.copy()
        further["week_start"] = further["deadline"].dt.to_period("W-SUN").apply(lambda p: p.start_time)
        for week_start, group in sorted(further.groupby("week_start")):
            week_end = week_start + pd.Timedelta(days=6)
            with st.expander(f"Week of {week_start.strftime('%d %b')} – {week_end.strftime('%d %b %Y')} · {len(group)} meters"):
                show_table(group, COLS, RENAME, sort_col="deadline")

with tab_overdue:
    st.subheader("Behind schedule — past Snag Date 4, not yet installed")
    overdue_full = df[df["status"] == "Overdue"]
    filtered = status_filters(overdue_full, "overdue")
    summary_counters(filtered, overdue_full, "overdue")
    if not filtered.empty:
        filtered = filtered.copy()
        filtered["days_overdue"] = (pd.Timestamp(date.today()) - filtered["deadline"]).dt.days
        show_table(
            filtered.sort_values("deadline"),
            ["stand", "meter_type", "unit_type", "wbho_section", "deadline", "days_overdue"],
            {"stand": "Stand", "meter_type": "Type", "unit_type": "Unit type", "wbho_section": "Section", "deadline": "Deadline (Snag 4)", "days_overdue": "Days overdue"},
        )
    else:
        st.success("Nothing overdue right now. 🎉")

with tab_installed:
    st.subheader("Installed meters log")
    installed_full = df[df["installed"]]

    # Serial search sits above the standard filters, prominent placement
    serial_search = st.text_input(
        "🔍 Search by meter serial number",
        placeholder="Type part of a serial — e.g. 10192017",
        key="installed_serial_search"
    )

    filtered = status_filters(installed_full, "installed")

    # Apply serial filter on top of section/type filters
    if serial_search.strip():
        filtered = filtered[filtered["serial"].str.contains(serial_search.strip(), case=False, na=False)]
        if filtered.empty:
            st.warning(f"No installed meter found with serial matching **{serial_search}**.")
        else:
            st.success(f"Found {len(filtered)} meter(s) matching serial **{serial_search}**.")

    summary_counters(filtered, installed_full, "installed")
    show_table(
        filtered,
        ["stand", "serial", "meter_type", "unit_type", "wbho_section", "commission_date", "deadline", "status", "amr"],
        {"stand": "Stand", "serial": "Meter serial", "meter_type": "Type", "unit_type": "Unit type", "wbho_section": "Section", "commission_date": "Commissioned", "deadline": "Deadline", "status": "Status", "amr": "AMR done"},
        sort_col="commission_date", ascending=False,
    )

with tab_calendar:
    st.subheader("Calendar view")
    st.caption("✅ installed · 🟧 upcoming deadline · 🟥 overdue deadline — based on Snag Date 4")

    cal_meter_type = st.radio("Meter type", ["All", "Water", "Electrical"], horizontal=True, key="cal_type")
    cal_df = df if cal_meter_type == "All" else df[df["meter_type"] == cal_meter_type]

    all_dates = pd.concat([cal_df["deadline"], cal_df["commission_date"]]).dropna()
    if all_dates.empty:
        st.info("No dated entries to show on a calendar.")
    else:
        months_available = sorted(set(zip(all_dates.dt.year, all_dates.dt.month)))
        month_labels = [f"{cal.month_name[m]} {y}" for y, m in months_available]
        today_ym = (date.today().year, date.today().month)
        default_idx = months_available.index(today_ym) if today_ym in months_available else len(months_available) - 1
        sel = st.selectbox("Month", month_labels, index=default_idx, key="cal_month")
        sel_year, sel_month = months_available[month_labels.index(sel)]

        # Build day -> events lookup
        installed_by_day = cal_df[cal_df["commission_date"].notna() & (cal_df["commission_date"].dt.year == sel_year) & (cal_df["commission_date"].dt.month == sel_month)]
        upcoming_by_day = cal_df[(~cal_df["installed"]) & cal_df["deadline"].notna() & (cal_df["deadline"].dt.year == sel_year) & (cal_df["deadline"].dt.month == sel_month) & (cal_df["status"] == "Due soon")]
        overdue_or_ontrack_deadline = cal_df[(~cal_df["installed"]) & cal_df["deadline"].notna() & (cal_df["deadline"].dt.year == sel_year) & (cal_df["deadline"].dt.month == sel_month) & cal_df["status"].isin(["Overdue", "On track"])]

        def counts_for_day(d):
            day_ts = pd.Timestamp(year=sel_year, month=sel_month, day=d)
            inst = int((installed_by_day["commission_date"].dt.day == d).sum())
            up = int((upcoming_by_day["deadline"].dt.day == d).sum())
            ov = int((overdue_or_ontrack_deadline[overdue_or_ontrack_deadline["status"] == "Overdue"]["deadline"].dt.day == d).sum())
            ot = int((overdue_or_ontrack_deadline[overdue_or_ontrack_deadline["status"] == "On track"]["deadline"].dt.day == d).sum())
            return inst, up, ov, ot

        month_matrix = cal.monthcalendar(sel_year, sel_month)
        today_d = date.today()

        html = "<table class='cal-table'><tr>" + "".join(f"<th>{d}</th>" for d in ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]) + "</tr>"
        for week in month_matrix:
            html += "<tr>"
            for d in week:
                if d == 0:
                    html += "<td class='cal-empty'></td>"
                    continue
                inst, up, ov, ot = counts_for_day(d)
                is_today = (sel_year == today_d.year and sel_month == today_d.month and d == today_d.day)
                day_num_cls = "cal-day-num today" if is_today else "cal-day-num"
                cell = f"<td><span class='{day_num_cls}'>{d}</span>"
                if inst: cell += f"<span class='cal-pill installed'>✅ {inst} installed</span>"
                if up: cell += f"<span class='cal-pill upcoming'>🟧 {up} due soon</span>"
                if ov: cell += f"<span class='cal-pill overdue'>🟥 {ov} overdue</span>"
                if ot: cell += f"<span class='cal-pill upcoming'>· {ot} scheduled</span>"
                cell += "</td>"
                html += cell
            html += "</tr>"
        html += "</table>"
        st.markdown(html, unsafe_allow_html=True)

with tab_sections:
    st.subheader("Section progress (WBHO subsections)")
    section_meter_type = st.radio("Meter type", ["All", "Water", "Electrical"], horizontal=True, key="sec_type")
    sec_df = df if section_meter_type == "All" else df[df["meter_type"] == section_meter_type]

    section_summary = sec_df.groupby("wbho_section").agg(
        total=("stand", "count"),
        installed=("installed", "sum"),
        deadline=("deadline", "max"),
        overdue=("status", lambda s: (s == "Overdue").sum()),
    ).reset_index()
    section_summary["progress"] = (section_summary["installed"] / section_summary["total"] * 100).round(0)
    section_summary = section_summary.sort_values(
        "wbho_section", key=lambda s: s.str.extract(r"(\d+)").fillna(0).astype(int)[0]
    )

    st.dataframe(
        section_summary.rename(columns={
            "wbho_section": "Section", "total": "Total", "installed": "Installed",
            "deadline": "Deadline", "overdue": "Overdue", "progress": "% complete",
        }),
        use_container_width=True,
        hide_index=True,
        column_config={
            "% complete": st.column_config.ProgressColumn("% complete", min_value=0, max_value=100, format="%d%%"),
            "Deadline": st.column_config.DatetimeColumn("Deadline", format="DD MMM YYYY"),
        },
    )

    st.markdown("##### Unit type progress")
    unit_summary = sec_df.groupby("unit_type").agg(total=("stand", "count"), installed=("installed", "sum")).reset_index()
    unit_summary["progress"] = (unit_summary["installed"] / unit_summary["total"] * 100).round(0)
    st.dataframe(
        unit_summary.rename(columns={"unit_type": "Unit type", "total": "Total", "installed": "Installed", "progress": "% complete"}),
        use_container_width=True, hide_index=True,
        column_config={"% complete": st.column_config.ProgressColumn("% complete", min_value=0, max_value=100, format="%d%%")},
    )

st.caption("Reads the spreadsheet pushed into this repo folder. Push an updated copy whenever meters are installed — the app picks up the most recently modified matching file automatically.")

# =====================================================================
# RETICULATION TAB — single-line diagram: Minisubs → Kiosks → Meters
# =====================================================================
with tab_retic:
    st.subheader("⚡ Electrical Reticulation — Single Line Diagram")
    st.caption("Supply → Minisub → Kiosk → Meters. Click a kiosk to expand stands. Click a stand chip to see its meter serial.")

    hierarchy = load_kiosk_data(data_path, mtime)

    # Serial search — highlights matching stand in diagram and shows result above
    retic_serial = st.text_input(
        "🔍 Search by meter serial number",
        placeholder="Type part of a serial to highlight the matching stand",
        key="retic_serial_search"
    )

    # Resolve serial → stand for feedback message
    if retic_serial.strip():
        match = df[df["installed"] & df["serial"].str.contains(retic_serial.strip(), case=False, na=False)]
        if match.empty:
            st.warning(f"No installed meter found with serial matching **{retic_serial}**.")
        else:
            for _, row in match.iterrows():
                kiosk_info = ""
                # Try to find its kiosk from the hierarchy
                for ms_id, ms in hierarchy.items():
                    for k in ms["kiosks"]:
                        if row["stand"] in k.get("stand_serials", {}):
                            kiosk_info = f" · Kiosk **{k['kiosk']}** (MS-{ms_id})"
                st.success(f"Stand **{row['stand']}** · Serial `{row['serial']}` · {row['meter_type']} · {row['wbho_section']}{kiosk_info} — expand that kiosk below to see the stand highlighted.")

    st.divider()

    # Summary KPIs across all kiosks
    all_kiosks = [k for ms in hierarchy.values() for k in ms["kiosks"]]
    total_planned = sum(k["planned"] for k in all_kiosks)
    total_installed = sum(k["installed"] for k in all_kiosks)
    total_amr = sum(k["amr_count"] for k in all_kiosks)
    total_outstanding = sum(max(0, k["planned"] - k["installed"]) for k in all_kiosks)

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total kiosks", sum(len(ms["kiosks"]) for ms in hierarchy.values()))
    k2.metric("Planned meter points", total_planned)
    k3.metric("Meters installed", total_installed, f"{round(total_installed/total_planned*100) if total_planned else 0}%")
    k4.metric("AMR commissioned", total_amr, f"{round(total_amr/total_installed*100) if total_installed else 0}% of installed")
    k5.metric("Outstanding", total_outstanding)

    st.divider()

    # Resolve which stands match the serial search so we can highlight them in JS
    highlight_stands = []
    if retic_serial.strip():
        match = df[df["installed"] & df["serial"].str.contains(retic_serial.strip(), case=False, na=False)]
        highlight_stands = match["stand"].tolist()

    diagram_data = []
    for ms_id in sorted(hierarchy.keys()):
        ms = hierarchy[ms_id]
        ms_installed = sum(k["installed"] for k in ms["kiosks"])
        ms_planned = sum(k["planned"] for k in ms["kiosks"])
        ms_amr = sum(k["amr_count"] for k in ms["kiosks"])
        diagram_data.append({
            "ms_id": ms_id,
            "serial": ms["serial"],
            "ms_installed": ms_installed,
            "ms_planned": ms_planned,
            "ms_amr": ms_amr,
            "kiosks": ms["kiosks"],
        })

    diagram_json = json.dumps(diagram_data)
    highlight_json = json.dumps(highlight_stands)

    html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'IBM Plex Mono', 'Courier New', monospace; background: #0e1117; color: #e0e0e0; padding: 16px; }}

  /* ---- Top bus bar ---- */
  .supply-bus {{
    display: flex; align-items: center; justify-content: center;
    margin-bottom: 0;
  }}
  .supply-box {{
    background: #1a2535; border: 2px solid #E69138; border-radius: 8px;
    padding: 10px 28px; font-size: 13px; font-weight: 700;
    color: #E69138; letter-spacing: .1em; text-transform: uppercase;
  }}
  .bus-line {{
    height: 4px; background: #E69138; flex: 1; max-width: 300px;
  }}

  /* ---- Minisub row ---- */
  .minisubs-row {{
    display: flex; justify-content: center; gap: 32px; margin-top: 0;
    align-items: flex-start;
  }}

  /* ---- Minisub column ---- */
  .ms-col {{ display: flex; flex-direction: column; align-items: center; min-width: 220px; }}

  .ms-vert-line {{ width: 3px; height: 32px; background: #E69138; }}

  .ms-box {{
    background: #1F3F66; border: 2px solid #5B86B3; border-radius: 10px;
    padding: 12px 18px; text-align: center; width: 100%; cursor: default;
    position: relative;
  }}
  .ms-box .ms-label {{ font-size: 12px; color: #9FB0C2; letter-spacing: .08em; margin-bottom: 4px; }}
  .ms-box .ms-title {{ font-size: 15px; font-weight: 700; color: #FFFFFF; }}
  .ms-box .ms-serial {{ font-size: 10px; color: #7A96B2; margin-top: 3px; }}
  .ms-box .ms-progress {{ margin-top: 8px; }}
  .progress-track {{ height: 5px; background: #2a3f55; border-radius: 3px; overflow: hidden; }}
  .progress-fill {{ height: 100%; background: #3F7D5C; border-radius: 3px; transition: width .3s; }}
  .ms-counts {{ font-size: 10px; color: #9FB0C2; margin-top: 4px; }}

  /* ---- Vertical connector from MS to kiosk row ---- */
  .kiosk-connector {{ width: 3px; height: 20px; background: #5B86B3; }}

  /* ---- Horizontal kiosk bus ---- */
  .kiosk-bus-wrap {{ position: relative; width: 100%; display: flex; justify-content: center; }}
  .kiosk-bus {{ height: 3px; background: #5B86B3; width: calc(100% - 20px); position: absolute; top: 0; }}

  /* ---- Kiosk grid ---- */
  .kiosk-grid {{
    display: flex; flex-direction: column; gap: 0; width: 100%;
    margin-top: 0; padding-top: 0;
  }}

  /* ---- Single kiosk entry ---- */
  .kiosk-entry {{ display: flex; flex-direction: column; align-items: center; width: 100%; }}
  .kiosk-drop-line {{ width: 3px; height: 20px; background: #5B86B3; }}

  .kiosk-node {{
    width: 100%; border-radius: 8px; border: 1.5px solid #334d6e;
    background: #131c2b; cursor: pointer; padding: 9px 12px;
    transition: border-color .15s, background .15s;
    position: relative;
  }}
  .kiosk-node:hover {{ border-color: #5B86B3; background: #1a2840; }}
  .kiosk-node.expanded {{ border-color: #E69138; background: #1e2b3a; }}
  .kiosk-node.all-installed {{ border-color: #3F7D5C; }}
  .kiosk-node.overdue {{ border-color: #BD4B2C; }}
  .kiosk-removed {{ opacity: .4; cursor: default; }}

  .kiosk-header {{ display: flex; align-items: center; gap: 8px; }}
  .kiosk-id {{ font-size: 12px; font-weight: 700; color: #c8d8eb; }}
  .kiosk-bar-wrap {{ flex: 1; }}
  .kiosk-mini-bar {{ height: 4px; border-radius: 2px; background: #2a3f55; overflow: hidden; }}
  .kiosk-mini-fill {{ height: 100%; border-radius: 2px; }}
  .kiosk-counts {{ font-size: 10px; color: #7A96B2; white-space: nowrap; }}
  .kiosk-chevron {{ font-size: 10px; color: #5B86B3; transition: transform .2s; }}
  .kiosk-chevron.open {{ transform: rotate(180deg); }}

  /* ---- Expandable stand list ---- */
  .kiosk-detail {{
    display: none; width: 100%; background: #0d1520; border: 1px solid #1e3050;
    border-top: none; border-radius: 0 0 8px 8px; padding: 8px 10px;
    font-size: 10px;
  }}
  .kiosk-detail.open {{ display: block; }}
  .stand-grid {{ display: flex; flex-wrap: wrap; gap: 4px; margin-top: 4px; }}
  .stand-chip {{
    padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 600;
    background: #3F7D5C22; color: #6eb88a; border: 1px solid #3F7D5C55;
    display: inline-flex; align-items: center; gap: 3px;
  }}
  .stand-chip.no-amr {{
    background: #E6913822; color: #d4902a; border: 1px solid #E6913866;
  }}
  .stand-chip.pending {{
    background: #1F3F6622; color: #7a9ec4; border: 1px solid #334d6e;
  }}
  .amr-dot {{
    display: inline-block; width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0;
  }}
  .amr-dot.ok {{ background: #3F7D5C; }}
  .amr-dot.missing {{ background: #E69138; }}
  .kiosk-amr-line {{
    font-size: 9px; color: #7A96B2; margin-top: 3px; display: flex; gap: 8px;
  }}
  .amr-ok-count {{ color: #6eb88a; }}
  .amr-miss-count {{ color: #d4902a; }}
  .legend-bar {{
    display: flex; gap: 14px; margin-bottom: 14px; flex-wrap: wrap;
    font-size: 10px; align-items: center; padding: 8px 12px;
    background: #0d1520; border: 1px solid #1e3050; border-radius: 6px;
  }}
  .legend-item {{ display: flex; align-items: center; gap: 5px; }}
  .legend-swatch {{
    width: 12px; height: 12px; border-radius: 3px; flex-shrink: 0;
  }}
  .stand-chip.highlight {{
    outline: 2px solid #E69138;
    box-shadow: 0 0 6px #E6913888;
  }}
  /* Serial popup tooltip */
  .serial-popup {{
    position: fixed; z-index: 9999;
    background: #152B45; border: 1px solid #5B86B3; border-radius: 8px;
    padding: 10px 14px; font-size: 11px; min-width: 200px;
    box-shadow: 0 6px 24px rgba(0,0,0,.5);
    display: none; pointer-events: none;
  }}
  .serial-popup.visible {{ display: block; }}
  .serial-popup .sp-stand {{ font-size: 13px; font-weight: 700; color: #fff; margin-bottom: 6px; }}
  .serial-popup .sp-row {{ display: flex; justify-content: space-between; gap: 16px; margin-bottom: 3px; }}
  .serial-popup .sp-label {{ color: #7A96B2; font-size: 10px; text-transform: uppercase; letter-spacing: .05em; }}
  .serial-popup .sp-val {{ color: #E0E8F0; font-family: monospace; font-size: 11px; }}
  .serial-popup .sp-amr-ok {{ color: #6eb88a; }}
  .serial-popup .sp-amr-miss {{ color: #d4902a; }}
  .serial-popup .sp-pending {{ color: #7a9ec4; font-style: italic; }}
</style>
</head>
<body>

<div class="serial-popup" id="serialPopup"></div>

<div class="legend-bar">
  <strong style="color:#9FB0C2;font-size:10px;letter-spacing:.06em;">STAND STATUS:</strong>
  <div class="legend-item"><span class="legend-swatch" style="background:#3F7D5C33;border:1px solid #3F7D5C66"></span><span style="color:#6eb88a">Meter ✓ · AMR ✓</span></div>
  <div class="legend-item"><span class="legend-swatch" style="background:#E6913833;border:1px solid #E6913866"></span><span style="color:#d4902a">Meter ✓ · AMR pending</span></div>
  <div class="legend-item"><span class="legend-swatch" style="background:#1F3F6633;border:1px solid #334d6e"></span><span style="color:#7a9ec4">Meter not yet installed</span></div>
</div>

<div class="supply-bus">
  <div class="bus-line"></div>
  <div class="supply-box">⚡ Utility Supply (MV)</div>
  <div class="bus-line"></div>
</div>

<div class="minisubs-row" id="diagramRoot"></div>

<script>
const data = {diagram_json};
const highlightStands = new Set({highlight_json});
const popup = document.getElementById('serialPopup');

function pct(inst, plan) {{
  return plan > 0 ? Math.round(inst / plan * 100) : 0;
}}

function barColor(p) {{
  if (p >= 100) return '#3F7D5C';
  if (p >= 60) return '#5B86B3';
  if (p >= 30) return '#E69138';
  return '#BD4B2C';
}}

function showPopup(e, stand, serial, inst, amr) {{
  const amrHtml = !inst
    ? `<span class="sp-pending">Meter not yet installed</span>`
    : amr
      ? `<span class="sp-amr-ok">✓ Commissioned</span>`
      : `<span class="sp-amr-miss">⚠ Pending</span>`;
  const serialHtml = serial
    ? `<span class="sp-val">${{serial}}</span>`
    : `<span class="sp-pending">No serial recorded</span>`;

  popup.innerHTML = `
    <div class="sp-stand">Stand ${{stand}}</div>
    <div class="sp-row"><span class="sp-label">Meter serial</span>${{serialHtml}}</div>
    <div class="sp-row"><span class="sp-label">Meter status</span><span class="sp-val">${{inst ? 'Installed' : 'Pending'}}</span></div>
    <div class="sp-row"><span class="sp-label">AMR</span>${{amrHtml}}</div>
  `;
  popup.classList.add('visible');
  const x = Math.min(e.clientX + 12, window.innerWidth - 220);
  const y = Math.min(e.clientY + 12, window.innerHeight - 140);
  popup.style.left = x + 'px';
  popup.style.top = y + 'px';
  e.stopPropagation();
}}

document.addEventListener('click', () => popup.classList.remove('visible'));

function buildDiagram() {{
  const root = document.getElementById('diagramRoot');

  data.forEach(ms => {{
    const col = document.createElement('div');
    col.className = 'ms-col';

    // Vertical line from supply
    const vline = document.createElement('div');
    vline.className = 'ms-vert-line';
    col.appendChild(vline);

    // Minisub box
    const p = pct(ms.ms_installed, ms.ms_planned);
    const msBox = document.createElement('div');
    msBox.className = 'ms-box';
    msBox.innerHTML = `
      <div class="ms-label">Minisub ${{ms.ms_id}}</div>
      <div class="ms-title">MS-${{ms.ms_id}}</div>
      <div class="ms-serial">Serial: ${{ms.serial}}</div>
      <div class="ms-progress">
        <div class="progress-track"><div class="progress-fill" style="width:${{p}}%; background:${{barColor(p)}}"></div></div>
        <div class="ms-counts">${{ms.ms_installed}} / ${{ms.ms_planned}} installed (${{p}}%)</div>
        <div class="ms-counts" style="margin-top:2px;">
          <span class="amr-ok-count">AMR: ${{ms.ms_amr}} ✓</span>
          ${{(ms.ms_installed - ms.ms_amr) > 0 ? ' &nbsp;<span class="amr-miss-count">⚠ ' + (ms.ms_installed - ms.ms_amr) + ' pending</span>' : ''}}
        </div>
      </div>
    `;
    col.appendChild(msBox);

    // Connector line down to kiosk bus
    const conn = document.createElement('div');
    conn.className = 'kiosk-connector';
    col.appendChild(conn);

    // Kiosk grid
    const grid = document.createElement('div');
    grid.className = 'kiosk-grid';

    ms.kiosks.forEach(k => {{
      const entry = document.createElement('div');
      entry.className = 'kiosk-entry';

      const dropLine = document.createElement('div');
      dropLine.className = 'kiosk-drop-line';
      entry.appendChild(dropLine);

      const isRemoved = k.planned === 0;
      const allInst = k.installed >= k.planned && k.planned > 0;
      const kp = pct(k.installed, k.planned);

      const node = document.createElement('div');
      node.className = 'kiosk-node' + (isRemoved ? ' kiosk-removed' : '') + (allInst ? ' all-installed' : '');
      node.innerHTML = `
        <div class="kiosk-header">
          <span class="kiosk-id">${{k.kiosk}}</span>
          <div class="kiosk-bar-wrap">
            <div class="kiosk-mini-bar">
              <div class="kiosk-mini-fill" style="width:${{kp}}%; background:${{barColor(kp)}}"></div>
            </div>
          </div>
          <span class="kiosk-counts">${{k.installed}}/${{k.planned}}</span>
          ${{!isRemoved ? '<span class="kiosk-chevron" id="chev-' + k.kiosk + '">▾</span>' : ''}}
        </div>
        ${{k.installed > 0 ? `<div class="kiosk-amr-line">
          <span>AMR:</span>
          <span class="amr-ok-count">✓ ${{k.amr_count}} done</span>
          ${{(k.installed - k.amr_count) > 0 ? '<span class="amr-miss-count">⚠ ' + (k.installed - k.amr_count) + ' pending</span>' : ''}}
        </div>` : ''}}
        ${{isRemoved ? '<div style="font-size:9px;color:#5a4040;margin-top:3px;">Kiosk removed</div>' : ''}}
      `;

      const detail = document.createElement('div');
      detail.className = 'kiosk-detail';
      detail.id = 'detail-' + k.kiosk;

      if (!isRemoved) {{
        const installedSet = new Set(k.installed_stands);
        const amrSet = new Set(k.amr_stands);
        const serialMap = k.stand_serials || {{}};
        const chipsHtml = k.stands.map(s => {{
          const inst = installedSet.has(s);
          const amr = amrSet.has(s);
          const serial = serialMap[s] || '';
          const isHighlight = highlightStands.has(s);
          let cls = 'stand-chip pending';
          if (inst && amr) cls = 'stand-chip';
          else if (inst) cls = 'stand-chip no-amr';
          if (isHighlight) cls += ' highlight';
          const dot = inst
            ? `<span class="amr-dot ${{amr ? 'ok' : 'missing'}}"></span>`
            : `<span class="amr-dot" style="background:#334d6e"></span>`;
          const title = inst
            ? `Stand ${{s}} · Serial: ${{serial || 'not recorded'}} · AMR: ${{amr ? '✓' : 'pending'}}`
            : `Stand ${{s}} · Meter not yet installed`;
          return `<span class="${{cls}}" data-stand="${{s}}" data-serial="${{serial}}" data-inst="${{inst}}" data-amr="${{amr}}" title="${{title}}">${{dot}}${{s}}</span>`;
        }}).join('');
        const amrMissing = k.installed - k.amr_count;
        detail.innerHTML = `
          <div style="font-size:9px;color:#5B86B3;margin-bottom:4px;">
            STANDS (${{k.stands.length}} in sheet · ${{k.planned}} planned)
            &nbsp;·&nbsp; <span class="amr-ok-count">AMR done: ${{k.amr_count}}</span>
            ${{amrMissing > 0 ? '&nbsp;·&nbsp; <span class="amr-miss-count">AMR pending: ' + amrMissing + '</span>' : ''}}
          </div>
          <div class="stand-grid">${{chipsHtml}}</div>
          ${{k.comment && k.comment !== 'nan' ? '<div class="comment-tag">📌 ' + k.comment + '</div>' : ''}}
        `;

        // Wire up chip click → serial popup
        detail.querySelectorAll('.stand-chip').forEach(chip => {{
          chip.style.cursor = 'pointer';
          chip.addEventListener('click', function(e) {{
            showPopup(
              e,
              chip.dataset.stand,
              chip.dataset.serial,
              chip.dataset.inst === 'True',
              chip.dataset.amr === 'True'
            );
          }});
        }});

        node.addEventListener('click', function() {{
          const d = document.getElementById('detail-' + k.kiosk);
          const chev = document.getElementById('chev-' + k.kiosk);
          const open = d.classList.toggle('open');
          node.classList.toggle('expanded', open);
          if (chev) chev.classList.toggle('open', open);
        }});
      }}

      entry.appendChild(node);
      entry.appendChild(detail);
      grid.appendChild(entry);
    }});

    col.appendChild(grid);
    root.appendChild(col);
  }});
}}

buildDiagram();

// If there are highlighted stands, auto-expand their kiosks
if (highlightStands.size > 0) {{
  data.forEach(ms => {{
    ms.kiosks.forEach(k => {{
      const hasMatch = k.stands && k.stands.some(s => highlightStands.has(s));
      if (hasMatch) {{
        const detail = document.getElementById('detail-' + k.kiosk);
        const node = detail && detail.previousElementSibling;
        const chev = document.getElementById('chev-' + k.kiosk);
        if (detail) detail.classList.add('open');
        if (node) node.classList.add('expanded');
        if (chev) chev.classList.add('open');
      }}
    }});
  }});
}}
</script>
</body>
</html>
"""

    components.html(html, height=900, scrolling=True)