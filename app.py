import streamlit as st
import pandas as pd
from datetime import datetime, date
import io

st.set_page_config(
    page_title="Sitari Evergreen — Outstanding Meters",
    page_icon="🔧",
    layout="wide",
)

# ---------- Styling ----------
st.markdown("""
<style>
.block-container {padding-top: 1.4rem;}
.stMetric {background: #FBF9F3; border: 1px solid #DCD6C4; border-radius: 10px; padding: 10px 14px;}
div[data-testid="stMetricLabel"] {font-size: 12px; text-transform: uppercase; letter-spacing: .05em; color: #3E5066;}
h1, h2, h3 {color: #152B45;}
.badge {display:inline-block; padding:2px 9px; border-radius:14px; font-size:11px; font-weight:600; font-family: monospace;}
.badge-overdue {background:#F5E2DB; color:#BD4B2C;}
.badge-upcoming {background:#FCEFDD; color:#B96E1E;}
.badge-ontrack {background:#E7EEF5; color:#1F3F66;}
</style>
""", unsafe_allow_html=True)

WATER_SHEET_CANDIDATES = ["Water meters", "Water Meters"]
ELEC_SHEET_CANDIDATES = ["Elec Meters", "Electrical Meters"]

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

@st.cache_data(show_spinner=False)
def load_data(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
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
        out["snag1"] = pd.to_datetime(coalesce_col(wdf, ["Snag Date 1"]), errors="coerce")
        out["snag2"] = pd.to_datetime(coalesce_col(wdf, ["Snag Date 2"]), errors="coerce")
        out["snag3"] = pd.to_datetime(coalesce_col(wdf, ["Snag Date 3"]), errors="coerce")
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
        out["snag1"] = pd.to_datetime(coalesce_col(edf, ["Snag Date 1"]), errors="coerce")
        out["snag2"] = pd.to_datetime(coalesce_col(edf, ["Snag Date 2"]), errors="coerce")
        out["snag3"] = pd.to_datetime(coalesce_col(edf, ["Snag Date 3"]), errors="coerce")
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


def badge(status):
    cls = {
        "Overdue": "badge-overdue",
        "Due soon": "badge-upcoming",
        "On track": "badge-ontrack",
        "Installed": "badge-ontrack",
        "Installed late": "badge-overdue",
    }.get(status, "badge-ontrack")
    return f'<span class="badge {cls}">{status}</span>'


# ---------- Header ----------
st.title("🔧 Sitari Evergreen — Outstanding Meters")
st.caption("Erf 1186 Sitari · Lifestyle Retirement Village · meter commissioning tracker for site technicians")

uploaded = st.file_uploader(
    "Upload the latest meter commissioning spreadsheet",
    type=["xlsx"],
    help="Upload EVG_SIT_FS_Meter_Commissioning_*.xlsx — the most recently updated version."
)

if uploaded is None:
    st.info("Upload the spreadsheet above to see outstanding installations. Ask the project office for the latest file if you don't have it.")
    st.stop()

df = load_data(uploaded.getvalue())

if df.empty:
    st.error("Couldn't find a 'Water meters' or 'Elec Meters' tab in this file. Check the sheet names and try again.")
    st.stop()

st.success(f"Loaded {len(df)} meter points · last refreshed {datetime.now().strftime('%d %b %Y, %H:%M')}")

# ---------- Filters (sidebar, technician-friendly) ----------
st.sidebar.header("Filter the list")
meter_types = st.sidebar.multiselect("Meter type", sorted(df["meter_type"].unique()), default=list(df["meter_type"].unique()))
sections = st.sidebar.multiselect("Section (WBHO)", sorted(df["wbho_section"].dropna().unique()))
unit_types = st.sidebar.multiselect("Unit type", sorted(df["unit_type"].dropna().unique()))
status_filter = st.sidebar.multiselect(
    "Status",
    ["Overdue", "Due soon", "On track", "Installed", "Installed late"],
    default=["Overdue", "Due soon", "On track"]
)
search_stand = st.sidebar.text_input("Search stand number")

filtered = df[df["meter_type"].isin(meter_types)]
if sections:
    filtered = filtered[filtered["wbho_section"].isin(sections)]
if unit_types:
    filtered = filtered[filtered["unit_type"].isin(unit_types)]
if status_filter:
    filtered = filtered[filtered["status"].isin(status_filter)]
if search_stand:
    filtered = filtered[filtered["stand"].str.contains(search_stand, case=False, na=False)]

# ---------- KPIs ----------
outstanding = df[~df["installed"]]
overdue = df[df["status"] == "Overdue"]
due_soon = df[df["status"] == "Due soon"]

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total meter points", len(df))
c2.metric("Installed", int(df["installed"].sum()), f"{round(df['installed'].mean()*100)}% complete")
c3.metric("Outstanding", len(outstanding))
c4.metric("Due within 14 days", len(due_soon))
c5.metric("Overdue", len(overdue), delta_color="inverse")

st.divider()

# ---------- Outstanding work list ----------
st.subheader("Outstanding installations (matches filters above)")

outstanding_view = filtered[~filtered["installed"]].sort_values("deadline")

if outstanding_view.empty:
    st.success("Nothing outstanding for this filter — all matching meters are installed. 🎉")
else:
    display_df = outstanding_view[[
        "stand", "meter_type", "unit_type", "wbho_section", "deadline", "status", "manufacturer", "model"
    ]].rename(columns={
        "stand": "Stand",
        "meter_type": "Type",
        "unit_type": "Unit type",
        "wbho_section": "Section",
        "deadline": "Deadline (Snag 4)",
        "status": "Status",
        "manufacturer": "Manufacturer",
        "model": "Model",
    })
    display_df["Deadline (Snag 4)"] = display_df["Deadline (Snag 4)"].dt.strftime("%d %b %Y")

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Status": st.column_config.TextColumn("Status"),
        },
    )

    csv = display_df.to_csv(index=False).encode("utf-8")
    st.download_button("Download this list as CSV", csv, file_name="outstanding_meters.csv", mime="text/csv")

st.divider()

# ---------- Upcoming by week ----------
st.subheader("Upcoming by week")
upcoming = df[(~df["installed"]) & (df["deadline"].notna()) & (df["deadline"] >= pd.Timestamp(date.today()))].copy()
if upcoming.empty:
    st.caption("No upcoming deadlines — everything outstanding is either overdue or has no deadline set.")
else:
    upcoming["week_start"] = upcoming["deadline"].dt.to_period("W-SUN").apply(lambda p: p.start_time)
    for week_start, group in sorted(upcoming.groupby("week_start")):
        week_end = week_start + pd.Timedelta(days=6)
        with st.expander(f"Week of {week_start.strftime('%d %b')} – {week_end.strftime('%d %b %Y')}  ·  {len(group)} meters", expanded=False):
            show = group[["stand", "meter_type", "unit_type", "wbho_section", "deadline"]].sort_values("deadline")
            show["deadline"] = show["deadline"].dt.strftime("%d %b %Y")
            show.columns = ["Stand", "Type", "Unit type", "Section", "Due date"]
            st.dataframe(show, use_container_width=True, hide_index=True)

st.divider()

# ---------- Section summary ----------
st.subheader("Section progress")
section_summary = df.groupby("wbho_section").agg(
    total=("stand", "count"),
    installed=("installed", "sum"),
    deadline=("deadline", "max"),
    overdue=("status", lambda s: (s == "Overdue").sum())
).reset_index()
section_summary["progress"] = (section_summary["installed"] / section_summary["total"] * 100).round(0)
section_summary = section_summary.sort_values("wbho_section", key=lambda s: s.str.extract(r"(\d+)").fillna(0).astype(int)[0])

st.dataframe(
    section_summary.rename(columns={
        "wbho_section": "Section", "total": "Total", "installed": "Installed",
        "deadline": "Deadline", "overdue": "Overdue", "progress": "% complete"
    }),
    use_container_width=True,
    hide_index=True,
    column_config={
        "% complete": st.column_config.ProgressColumn("% complete", min_value=0, max_value=100, format="%d%%"),
        "Deadline": st.column_config.DatetimeColumn("Deadline", format="DD MMM YYYY"),
    }
)

st.caption("Built from your uploaded commissioning sheet. Refresh by uploading a newer file — nothing is saved on the server between sessions.")