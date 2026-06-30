import streamlit as st
import pandas as pd
import calendar as cal
from datetime import datetime, date
import glob
import os

st.set_page_config(
    page_title="Sitari Evergreen — Meter Commissioning",
    page_icon="🔧",
    layout="wide",
)

# ---------- Styling ----------
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; max-width: 1300px;}
.stMetric {background: #FBF9F3 !important; border: 1px solid #DCD6C4; border-radius: 10px; padding: 10px 14px;}
div[data-testid="stMetric"] {background: #FBF9F3 !important; border: 1px solid #DCD6C4; border-radius: 10px; padding: 10px 14px;}
div[data-testid="stMetricLabel"] {font-size: 12px; text-transform: uppercase; letter-spacing: .05em; color: #3E5066 !important;}
div[data-testid="stMetricLabel"] p {color: #3E5066 !important;}
div[data-testid="stMetricValue"] {color: #152B45 !important;}
div[data-testid="stMetricValue"] div {color: #152B45 !important;}
div[data-testid="stMetricDelta"] {color: #3F7D5C !important;}
div[data-testid="stMetricDelta"] svg {fill: #3F7D5C !important;}
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

    if water_name:
        wdf = xls.parse(water_name)
        wdf.columns = [str(c).strip() for c in wdf.columns]
        out = pd.DataFrame()
        out["stand"] = coalesce_col(wdf, ["Stand Number"]).astype(str).str.strip()
        out["unit_type"] = coalesce_col(wdf, ["Section"])
        out["wbho_section"] = coalesce_col(wdf, ["WBHO Subsection"])
        out["manufacturer"] = coalesce_col(wdf, ["Manufacturer"])
        out["model"] = coalesce_col(wdf, ["Meter Model"])
        out["serial"] = coalesce_col(wdf, ["Meter serial number"])
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
        out["serial"] = coalesce_col(edf, ["Meter Serial"])
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
tab_outstanding, tab_upcoming, tab_overdue, tab_installed, tab_calendar, tab_sections = st.tabs(
    ["🟦 Outstanding", "🟧 Upcoming", "🟥 Overdue", "🟩 Installed", "📅 Calendar", "📊 Sections"]
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
    filtered = status_filters(installed_full, "installed")
    summary_counters(filtered, installed_full, "installed")
    show_table(
        filtered,
        ["stand", "meter_type", "unit_type", "wbho_section", "commission_date", "deadline", "status", "amr"],
        {"stand": "Stand", "meter_type": "Type", "unit_type": "Unit type", "wbho_section": "Section", "commission_date": "Commissioned", "deadline": "Deadline", "status": "Status", "amr": "AMR done"},
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