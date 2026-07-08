import streamlit as st
import pandas as pd
import calendar as cal
from datetime import datetime, date, timedelta
import glob
import os
import json
import re
import io
import xml.etree.ElementTree as ET
import streamlit.components.v1 as components
import folium
from streamlit_folium import st_folium
try:
    import paramiko
    PARAMIKO_AVAILABLE = True
except ImportError:
    PARAMIKO_AVAILABLE = False

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

def find_kml_file():
    """Auto-detect the KML file in the same folder as app.py."""
    matches = glob.glob(os.path.join(os.path.dirname(__file__), "*.kml"))
    if not matches:
        return None
    return max(matches, key=os.path.getmtime)


# ── AMR / SFTP helpers ────────────────────────────────────────────────────────

AMR_CACHE_FILE = os.path.join(os.path.dirname(__file__), "amr_cache.json")
AMR_DB_FILE    = os.path.join(os.path.dirname(__file__), "amr_history.db")
CSV_FILENAME_RE = re.compile(r"^[A-Z]+_(\d{8})_(\d{6})\.csv$", re.IGNORECASE)

# ── SQLite history database ───────────────────────────────────────────────────
import sqlite3

def _db_conn():
    conn = sqlite3.connect(AMR_DB_FILE, check_same_thread=False)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS amr_readings (
            serial       TEXT NOT NULL,
            reading_date TEXT NOT NULL,
            reading_value REAL,
            low_battery  INTEGER DEFAULT 0,
            file_ts      TEXT,
            PRIMARY KEY (serial, reading_date)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_serial ON amr_readings(serial)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_date   ON amr_readings(reading_date)")
    conn.commit()
    return conn

def db_upsert_readings(readings_dict, file_ts_str):
    """Insert new readings into the history DB. Existing (serial, reading_date) pairs are ignored."""
    conn = _db_conn()
    rows = [
        (serial, v["reading_date"], v["reading_value"], v["low_battery"], file_ts_str)
        for serial, v in readings_dict.items()
        if v.get("reading_date")
    ]
    conn.executemany(
        "INSERT OR IGNORE INTO amr_readings (serial, reading_date, reading_value, low_battery, file_ts) VALUES (?,?,?,?,?)",
        rows
    )
    conn.commit()
    conn.close()
    return len(rows)

def db_get_history(serial):
    """Return DataFrame of all readings for a serial, sorted by date ascending."""
    conn = _db_conn()
    df = pd.read_sql_query(
        "SELECT reading_date, reading_value, low_battery FROM amr_readings WHERE serial=? ORDER BY reading_date ASC",
        conn, params=(serial,)
    )
    conn.close()
    if not df.empty:
        df["reading_date"] = pd.to_datetime(df["reading_date"], errors="coerce")
    return df

@st.cache_data(ttl=30, show_spinner=False)
def db_stats():
    """Return (total_rows, distinct_serials, min_date, max_date) from DB."""
    try:
        conn = _db_conn()
        row = conn.execute(
            "SELECT COUNT(*), COUNT(DISTINCT serial), MIN(reading_date), MAX(reading_date) FROM amr_readings"
        ).fetchone()
        conn.close()
        return row
    except Exception:
        return (0, 0, None, None)


def _parse_csv_filename_ts(name):
    """Extract datetime from SIT_YYYYMMDD_HHMMSS.csv → datetime or None."""
    m = CSV_FILENAME_RE.match(os.path.basename(name))
    if m:
        try:
            return datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S")
        except ValueError:
            pass
    return None

def parse_amr_csv(csv_bytes_or_str, file_ts=None):
    """
    Parse one AMR CSV file. Returns a dict:
      {base_serial: {reading_date, reading_value, low_battery, file_ts}}
    Skips NR (reverse channel) rows.
    """
    if isinstance(csv_bytes_or_str, (bytes, bytearray)):
        text = csv_bytes_or_str.decode("utf-8", errors="replace")
    else:
        text = csv_bytes_or_str

    df = pd.read_csv(io.StringIO(text))
    df.columns = [c.strip() for c in df.columns]

    df["_addr"] = df["METER_ADDRESS"].astype(str).str.strip()
    df = df[~df["_addr"].str.endswith("NR")].copy()

    # Parse reading date — format "DD/MM/YYYY HH:MM:SS GMT+2"
    df["_reading_dt"] = pd.to_datetime(
        df["READING_DATE"].astype(str).str.replace(r"\s*GMT[+-]\d+", "", regex=True),
        dayfirst=True, errors="coerce"
    )

    result = {}
    for _, row in df.iterrows():
        serial = row["_addr"]
        if not serial or serial == "nan":
            continue
        result[serial] = {
            "reading_date": row["_reading_dt"].isoformat() if pd.notna(row["_reading_dt"]) else None,
            "reading_value": float(row["READING_VALUE"]) if pd.notna(row["READING_VALUE"]) else None,
            "low_battery": int(row.get("LOW_BATTERY", 0) or 0),
            "file_ts": file_ts.isoformat() if file_ts else None,
        }
    return result


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_amr_from_sftp(host, port, username, password, directory, _cache_bust=0):
    """
    Connect to SFTP, find the most recent CSV in `directory`,
    download and parse it. Returns (readings_dict, filename, file_ts, error_str).
    Cached for 1 hour — set _cache_bust=int(time.time()//3600) to force hourly refresh.
    """
    if not PARAMIKO_AVAILABLE:
        return {}, None, None, "paramiko not installed"
    try:
        transport = paramiko.Transport((host, int(port)))
        transport.connect(username=username, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)

        # List all CSV files in directory, pick the newest by filename timestamp
        try:
            files = sftp.listdir(directory)
        except FileNotFoundError:
            transport.close()
            return {}, None, None, f"Directory not found: {directory}"

        csv_files = [f for f in files if f.lower().endswith(".csv")]
        if not csv_files:
            transport.close()
            return {}, None, None, "No CSV files found in directory"

        # Sort by embedded timestamp in filename; fall back to mtime
        def file_sort_key(name):
            ts = _parse_csv_filename_ts(name)
            return ts if ts else datetime.min

        best = max(csv_files, key=file_sort_key)
        file_ts = _parse_csv_filename_ts(best)

        remote_path = directory.rstrip("/") + "/" + best
        buf = io.BytesIO()
        sftp.getfo(remote_path, buf)
        transport.close()

        readings = parse_amr_csv(buf.getvalue(), file_ts)
        # Persist to history DB so we accumulate data over time
        try:
            db_upsert_readings(readings, file_ts.isoformat() if file_ts else None)
        except Exception:
            pass
        return readings, best, file_ts, None

    except Exception as e:
        return {}, None, None, str(e)


def fetch_amr_bulk_history(host, port, username, password, directory, hours=24, progress_cb=None):
    """
    Connect to SFTP, download ALL CSV files whose filename timestamp falls within
    the last `hours` hours, parse each, and upsert into the history DB.
    Returns (files_processed, new_readings_inserted, latest_readings_dict, error_str).
    `progress_cb(current, total, filename)` is called for each file if provided.
    Not cached — intended for one-off bulk historical loads.
    """
    if not PARAMIKO_AVAILABLE:
        return 0, 0, {}, "paramiko not installed"
    try:
        transport = paramiko.Transport((host, int(port)))
        transport.connect(username=username, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)

        try:
            files = sftp.listdir(directory)
        except FileNotFoundError:
            transport.close()
            return 0, 0, {}, f"Directory not found: {directory}"

        csv_files = [f for f in files if f.lower().endswith(".csv")]
        if not csv_files:
            transport.close()
            return 0, 0, {}, "No CSV files found in directory"

        cutoff = datetime.now() - timedelta(hours=hours)

        def ts_key(name):
            ts = _parse_csv_filename_ts(name)
            return ts if ts else datetime.min

        # Filter to files within the window, newest first
        in_window = sorted(
            [f for f in csv_files if ts_key(f) >= cutoff],
            key=ts_key, reverse=True
        )

        if not in_window:
            transport.close()
            return 0, 0, {}, f"No CSV files found within the last {hours} hours"

        files_done = 0
        total_new  = 0
        latest_readings = {}

        for i, fname in enumerate(in_window):
            if progress_cb:
                progress_cb(i + 1, len(in_window), fname)

            file_ts = _parse_csv_filename_ts(fname)
            remote_path = directory.rstrip("/") + "/" + fname
            try:
                buf = io.BytesIO()
                sftp.getfo(remote_path, buf)
                readings = parse_amr_csv(buf.getvalue(), file_ts)
                n = db_upsert_readings(readings, file_ts.isoformat() if file_ts else None)
                total_new += n
                files_done += 1
                # Keep the readings from the most recent file as the "current" snapshot
                if i == 0:
                    latest_readings = readings
            except Exception:
                continue  # skip corrupt / unreadable files silently

        transport.close()
        return files_done, total_new, latest_readings, None

    except Exception as e:
        return 0, 0, {}, str(e)


def load_amr_cache():
    """Load persisted AMR readings from local JSON cache file."""
    if os.path.exists(AMR_CACHE_FILE):
        try:
            with open(AMR_CACHE_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_amr_cache(data):
    """Persist AMR readings to local JSON cache file."""
    try:
        with open(AMR_CACHE_FILE, "w") as f:
            json.dump(data, f)
    except Exception:
        pass

def amr_status_info(reading_date_iso):
    """
    Return (label, color_hex, badge_class) based on hours since last reading.
    Handles None, NaN, empty string, and non-ISO strings gracefully.
    """
    if not reading_date_iso or str(reading_date_iso) in ("nan", "None", "NaT", ""):
        return "No reading", "#607080", "amr-never"
    try:
        dt = datetime.fromisoformat(str(reading_date_iso))
    except (ValueError, TypeError):
        return "No reading", "#607080", "amr-never"
    hours_ago = (datetime.now() - dt).total_seconds() / 3600
    if hours_ago <= 24:
        return f"Last {int(hours_ago)}h ago", "#2E7D52", "amr-green"
    if hours_ago <= 72:
        days = hours_ago / 24
        return f"{days:.1f}d ago", "#D4AC0D", "amr-yellow"
    if hours_ago <= 168:
        days = hours_ago / 24
        return f"{days:.0f}d ago", "#E67E22", "amr-orange"
    days = hours_ago / 24
    return f"{days:.0f}d ago", "#BD4B2C", "amr-red"

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
        out["faulty"] = coalesce_col(wdf, ["Faulty Meter"]).fillna(False).astype(bool)
        out["faulty_replaced"] = coalesce_col(wdf, ["Faulty Replaced"]).fillna(False).astype(bool)
        out["replacement_date"] = pd.to_datetime(coalesce_col(wdf, ["Replacement Date"]), errors="coerce")
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
        out["faulty"] = coalesce_col(edf, ["Faulty Meter"]).fillna(False).astype(bool)
        out["faulty_replaced"] = coalesce_col(edf, ["Faulty Replaced"]).fillna(False).astype(bool)
        out["replacement_date"] = pd.to_datetime(coalesce_col(edf, ["Replacement Date"]), errors="coerce")
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
    edf["deadline"] = pd.to_datetime(
        coalesce_col(edf, ["Snag Date 4"]), errors="coerce"
    )

    def _fmt_serial(v):
        if pd.isna(v): return ""
        try: return str(int(float(v)))
        except: return str(v).strip()

    edf["serial_str"] = edf["Meter Serial"].apply(_fmt_serial)

    today = pd.Timestamp(date.today())
    edf["days_to_deadline"] = (edf["deadline"] - today).dt.days

    def _stand_status(row):
        if row["installed"]:
            return "installed"
        if pd.notna(row["deadline"]) and row["deadline"] < today:
            return "overdue"
        if pd.notna(row["deadline"]) and row["days_to_deadline"] <= 14:
            return "due_soon"
        return "on_track"

    edf["stand_status"] = edf.apply(_stand_status, axis=1)

    # Aggregate per kiosk
    def agg_stands(x):
        return sorted(x.tolist())

    def agg_installed_stands(x):
        return sorted(edf.loc[x.index][edf.loc[x.index, "installed"]]["stand_str"].tolist())

    def agg_amr_stands(x):
        return sorted(edf.loc[x.index][edf.loc[x.index, "amr_done"]]["stand_str"].tolist())

    def agg_stand_serials(x):
        subset = edf.loc[x.index][edf.loc[x.index, "serial_str"] != ""]
        return {row["stand_str"]: row["serial_str"] for _, row in subset.iterrows()}

    def agg_overdue_stands(x):
        return sorted(edf.loc[x.index][edf.loc[x.index, "stand_status"] == "overdue"]["stand_str"].tolist())

    def agg_due_soon_stands(x):
        return sorted(edf.loc[x.index][edf.loc[x.index, "stand_status"] == "due_soon"]["stand_str"].tolist())

    def agg_stand_deadlines(x):
        """Return dict of stand → {deadline_str, days_to, status} for uninstalled stands."""
        subset = edf.loc[x.index][~edf.loc[x.index, "installed"] & edf.loc[x.index, "deadline"].notna()]
        return {
            row["stand_str"]: {
                "deadline": row["deadline"].strftime("%d %b %Y"),
                "days_to": int(row["days_to_deadline"]) if pd.notna(row["days_to_deadline"]) else 999,
                "status": row["stand_status"],
            }
            for _, row in subset.iterrows()
        }

    kiosk_agg = edf.groupby("Kiosk Number").agg(
        installed_count=("installed", "sum"),
        amr_count=("amr_done", "sum"),
        total_count=("Stand Number", "count"),
        stands=("stand_str", agg_stands),
        installed_stands=("stand_str", agg_installed_stands),
        amr_stands=("stand_str", agg_amr_stands),
        stand_serials=("stand_str", agg_stand_serials),
        overdue_stands=("stand_str", agg_overdue_stands),
        due_soon_stands=("stand_str", agg_due_soon_stands),
        stand_deadlines=("stand_str", agg_stand_deadlines),
    ).reset_index()
    kiosk_agg.columns = [
        "kiosk", "installed", "amr_count", "total",
        "stands", "installed_stands", "amr_stands", "stand_serials",
        "overdue_stands", "due_soon_stands", "stand_deadlines",
    ]

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
    merged["overdue_stands"] = merged["overdue_stands"].apply(lambda x: x if isinstance(x, list) else [])
    merged["due_soon_stands"] = merged["due_soon_stands"].apply(lambda x: x if isinstance(x, list) else [])
    merged["stand_deadlines"] = merged["stand_deadlines"].apply(lambda x: x if isinstance(x, dict) else {})

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
            "overdue_stands": row["overdue_stands"],
            "due_soon_stands": row["due_soon_stands"],
            "stand_deadlines": row["stand_deadlines"],
            "comment": str(row.get("Comments", "") or ""),
        })

    # Sort kiosks within each minisub
    for ms_id in hierarchy:
        hierarchy[ms_id]["kiosks"].sort(key=lambda k: sort_kiosk_key(k["kiosk"]))

    return hierarchy


# ── KML helpers ──────────────────────────────────────────────────────────────

KML_NS = "{http://www.opengis.net/kml/2.2}"

def _kml_text(el, tag):
    child = el.find(f"{KML_NS}{tag}")
    return child.text.strip() if child is not None and child.text else ""

def _parse_coords(coord_str):
    """Parse KML coordinate string into list of [lat, lon] pairs."""
    pairs = []
    for token in coord_str.strip().split():
        parts = token.split(",")
        if len(parts) >= 2:
            try:
                lon, lat = float(parts[0]), float(parts[1])
                pairs.append([lat, lon])
            except ValueError:
                pass
    return pairs

def parse_kml(kml_bytes):
    """
    Parse KML exported from Google Earth. Recognises three folder types:
      - FS Units / units  → stand polygons
      - Kiosks            → kiosk point placemarks
      - Minisubs          → minisub point placemarks
    Returns: polygons, kiosks, minisubs
    """
    import re as _re
    root = ET.fromstring(kml_bytes)
    polygons, kiosks, minisubs = [], [], []

    def _pm_name(pm):
        n = pm.find(f"{KML_NS}name")
        return (n.text or "").strip() if n is not None else ""

    def _extract_point(pm):
        pt = pm.find(f".//{KML_NS}Point/{KML_NS}coordinates")
        if pt is not None and pt.text:
            parts = pt.text.strip().split(",")
            if len(parts) >= 2:
                try:
                    return float(parts[1]), float(parts[0])  # lat, lon
                except ValueError:
                    pass
        return None, None

    def _extract_polygon(pm):
        outer = pm.find(
            f".//{KML_NS}Polygon/{KML_NS}outerBoundaryIs/{KML_NS}LinearRing/{KML_NS}coordinates"
        )
        if outer is not None and outer.text:
            return _parse_coords(outer.text)
        return []

    for folder in root.iter(f"{KML_NS}Folder"):
        fn = folder.find(f"{KML_NS}name")
        fname = (fn.text or "").strip().lower() if fn is not None else ""
        is_units   = any(k in fname for k in ("unit", "fs unit", "house", "stand"))
        is_kiosk   = "kiosk" in fname
        is_minisub = any(k in fname for k in ("minisub", "mini sub", "mini-sub"))

        for pm in folder.findall(f"{KML_NS}Placemark"):
            name = _pm_name(pm)
            has_poly  = pm.find(f".//{KML_NS}Polygon") is not None
            has_point = pm.find(f".//{KML_NS}Point")   is not None

            if has_poly and (is_units or (not is_kiosk and not is_minisub)):
                coords = _extract_polygon(pm)
                if coords:
                    polygons.append({"name": name, "coords": coords})
            elif has_point and is_minisub:
                lat, lon = _extract_point(pm)
                if lat is not None:
                    minisubs.append({"name": name, "lat": lat, "lon": lon})
            elif has_point and is_kiosk:
                lat, lon = _extract_point(pm)
                if lat is not None:
                    kiosks.append({"name": name, "lat": lat, "lon": lon})
            elif has_point and not is_units:
                lat, lon = _extract_point(pm)
                if lat is not None:
                    if _re.match(r"\d+EVE\d+", name, _re.IGNORECASE):
                        kiosks.append({"name": name, "lat": lat, "lon": lon})
                    elif _re.match(r"MS\d*", name, _re.IGNORECASE):
                        minisubs.append({"name": name, "lat": lat, "lon": lon})

    # Last-resort fallback if no folder names matched
    if not polygons and not kiosks and not minisubs:
        import re as _re
        for pm in root.iter(f"{KML_NS}Placemark"):
            name = _pm_name(pm)
            if pm.find(f".//{KML_NS}Polygon") is not None:
                coords = _extract_polygon(pm)
                if coords:
                    polygons.append({"name": name, "coords": coords})
            elif pm.find(f".//{KML_NS}Point") is not None:
                lat, lon = _extract_point(pm)
                if lat is not None:
                    if _re.match(r"\d+EVE\d+", name, _re.IGNORECASE):
                        kiosks.append({"name": name, "lat": lat, "lon": lon})
                    elif _re.match(r"MS\d*", name, _re.IGNORECASE):
                        minisubs.append({"name": name, "lat": lat, "lon": lon})

    return polygons, kiosks, minisubs


def stand_map_color(stand_str, df, faulty_mode=False):
    """Return (fill_color, opacity, status_label) for a stand."""
    rows = df[df["stand"] == stand_str]
    if rows.empty:
        return "#607080", 0.35, "No data"

    # Faulty overlay mode — colour by fault status instead of install status
    if faulty_mode:
        faulty_rows = rows[rows["faulty"]] if "faulty" in rows.columns else pd.DataFrame()
        if not faulty_rows.empty:
            replaced = faulty_rows["faulty_replaced"].any() if "faulty_replaced" in faulty_rows.columns else False
            if replaced:
                return "#8B5CF6", 0.85, "Faulty — replaced"        # purple
            else:
                return "#EF4444", 0.90, "Faulty — awaiting replacement"  # bright red
        return "#1E3A2F", 0.30, "No fault recorded"   # very dark, recede into background

    elec  = rows[rows["meter_type"] == "Electrical"]
    water = rows[rows["meter_type"] == "Water"]
    elec_inst  = bool(elec["installed"].any())  if not elec.empty  else False
    water_inst = bool(water["installed"].any()) if not water.empty else False
    elec_amr   = bool(elec["amr"].any())        if not elec.empty  else False
    water_amr  = bool(water["amr"].any())       if not water.empty else False
    today = pd.Timestamp(date.today())
    overdue  = rows[~rows["installed"] & rows["deadline"].notna() & (rows["deadline"] < today)]
    due_soon = rows[~rows["installed"] & rows["deadline"].notna() &
                    ((rows["deadline"] - today).dt.days.between(0, 14))]
    if elec_inst and water_inst and elec_amr and water_amr:
        return "#2E7D52", 0.80, "Meter \u2713 \u00b7 AMR \u2713"
    if elec_inst and water_inst:
        return "#E69138", 0.80, "Meters \u2713 \u00b7 AMR pending"
    if elec_inst or water_inst:
        return "#5B86B3", 0.75, "Partially installed"
    if not overdue.empty:
        return "#BD4B2C", 0.85, "Overdue"
    if not due_soon.empty:
        return "#D4AC0D", 0.80, "Due soon"
    return "#3A5068", 0.55, "On track"


# ── Cached map HTML renderers ──────────────────────────────────────────────────
# These wrap the folium builders and cache the rendered HTML string so that
# filter changes don't rebuild hundreds of polygons from scratch on every rerun.

def _poly_hash(polygons, kiosks, minisubs=()):
    """Stable string key for KML geometry (only changes when KML file changes)."""
    import hashlib, json
    data = {"p": [(p["name"], p["coords"][0]) for p in polygons[:5]],
            "k": [k["name"] for k in kiosks],
            "n": len(polygons)}
    return hashlib.md5(json.dumps(data, default=str).encode()).hexdigest()[:16]

def _df_hash(df):
    """Fast hash of a DataFrame's content."""
    import hashlib
    return hashlib.md5(pd.util.hash_pandas_object(df, index=False).values.tobytes()).hexdigest()[:16]

def _amr_hash(amr_readings):
    """Hash of the AMR readings dict (keyed on serial → date pairs)."""
    import hashlib, json
    snapshot = sorted((k, v.get("reading_date","")) for k, v in amr_readings.items())
    return hashlib.md5(json.dumps(snapshot[:50]).encode()).hexdigest()[:16]


@st.cache_data(show_spinner=False, ttl=600)
def cached_estate_map_html(geom_hash, df_hash_val, center, show_labels, faulty_mode,
                            _polygons, _kiosks, _minisubs, _df):
    """Build estate map and return HTML. Cache key excludes underscore-prefixed args."""
    m = build_estate_map(_polygons, _kiosks, _minisubs, _df, center, show_labels, faulty_mode)
    return m.get_root().render()


@st.cache_data(show_spinner=False, ttl=120)
def cached_amr_map_html(geom_hash, df_hash_val, amr_hash_val, faulty_hash_val, center,
                         _polygons, _kiosks, _df, _amr_readings, _faulty_df):
    """Build AMR map and return HTML. Cache key excludes underscore-prefixed args."""
    m = build_amr_map(_polygons, _kiosks, _df, _amr_readings, _faulty_df, center)
    return m.get_root().render()


def render_cached_map(html_str, height=640):
    """Render a cached folium HTML string via st.components."""
    components.html(html_str, height=height, scrolling=False)


def build_estate_map(polygons, kiosks, minisubs, df, center, show_labels=True, faulty_mode=False):
    m = folium.Map(location=center, zoom_start=19, tiles=None,
                   max_zoom=22, zoom_control=True)
    folium.TileLayer(
        tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        attr="Esri World Imagery", name="Satellite", overlay=False, control=True,
        max_zoom=22, max_native_zoom=19,
    ).add_to(m)
    folium.TileLayer(
        tiles="https://{s}.basemaps.cartocdn.com/dark_only_labels/{z}/{x}/{y}{r}.png",
        attr="(c) OpenStreetMap (c) CARTO", name="Labels",
        overlay=True, control=True, opacity=0.7,
    ).add_to(m)

    # ── Unit polygons ─────────────────────────────────────────────────
    for poly in polygons:
        stand = poly["name"].strip()
        color, opacity, status_label = stand_map_color(stand, df, faulty_mode)
        rows = df[df["stand"] == stand]
        elec  = rows[rows["meter_type"] == "Electrical"]
        water = rows[rows["meter_type"] == "Water"]

        # Faulty badge for popup
        faulty_rows = rows[rows["faulty"]] if "faulty" in rows.columns else pd.DataFrame()
        faulty_badge = ""
        if not faulty_rows.empty:
            replaced = faulty_rows["faulty_replaced"].any()
            repl_date = faulty_rows["replacement_date"].dropna()
            repl_str = repl_date.iloc[0].strftime("%d %b %Y") if not repl_date.empty else ""
            if replaced:
                faulty_badge = (f"<div style='margin:4px 0 6px;padding:3px 10px;border-radius:6px;"
                                f"background:#8B5CF622;color:#8B5CF6;border:1px solid #8B5CF688;"
                                f"font-size:11px;font-weight:600'>✅ Faulty — replaced"
                                f"{(' on ' + repl_str) if repl_str else ''}</div>")
            else:
                faulty_badge = (f"<div style='margin:4px 0 6px;padding:3px 10px;border-radius:6px;"
                                f"background:#EF444422;color:#EF4444;border:1px solid #EF444488;"
                                f"font-size:11px;font-weight:600'>⚠️ Faulty — awaiting replacement</div>")

        def _row(r_df, label):
            if r_df.empty:
                return f"<tr><td colspan='3' style='color:#aaa;padding:4px 0'><b>{label}:</b> not in sheet</td></tr>"
            r = r_df.iloc[0]
            serial = r["serial"] if r["serial"] else "\u2014"
            inst   = ("\u2705 " + r["commission_date"].strftime("%d %b %Y")
                      if r["installed"] and pd.notna(r["commission_date"]) else "\u23f3 Pending")
            amr_s  = "\u2705" if r["amr"] else "\u23f3"
            dl     = r["deadline"].strftime("%d %b %Y") if pd.notna(r["deadline"]) else "\u2014"
            return (
                f"<tr><td style='padding:3px 6px 1px;font-weight:700'>{label}</td>"
                f"<td style='padding:3px 6px 1px'>{inst}</td>"
                f"<td style='padding:3px 6px 1px'>AMR {amr_s}</td></tr>"
                f"<tr><td style='padding:1px 6px 4px;font-size:10px;color:#777'>Serial</td>"
                f"<td style='padding:1px 6px 4px;font-family:monospace;font-size:10px' colspan='2'>{serial}</td></tr>"
                f"<tr><td style='padding:1px 6px 6px;font-size:10px;color:#777'>Deadline</td>"
                f"<td style='padding:1px 6px 6px;font-size:10px' colspan='2'>{dl}</td></tr>"
            )

        popup_html = (
            f"<div style='font-family:sans-serif;min-width:240px;max-width:300px'>"
            f"<div style='font-size:15px;font-weight:700;color:#152B45;margin-bottom:6px'>Stand {stand}</div>"
            f"<div style='display:inline-block;padding:2px 10px;border-radius:10px;font-size:11px;"
            f"font-weight:600;background:{color}22;color:{color};border:1px solid {color}88;"
            f"margin-bottom:6px'>{status_label}</div>"
            f"{faulty_badge}"
            f"<table style='width:100%;font-size:12px;border-collapse:collapse;border-top:1px solid #eee'>"
            f"{_row(elec,'Electrical')}{_row(water,'Water')}"
            f"</table></div>"
        )

        folium.Polygon(
            locations=poly["coords"],
            color=color, weight=1.5,
            fill=True, fill_color=color, fill_opacity=opacity,
            tooltip=folium.Tooltip(f"<b>Stand {stand}</b><br>{status_label}", sticky=True),
            popup=folium.Popup(popup_html, max_width=300),
        ).add_to(m)

        if show_labels and poly["coords"]:
            clat = sum(c[0] for c in poly["coords"]) / len(poly["coords"])
            clon = sum(c[1] for c in poly["coords"]) / len(poly["coords"])
            folium.Marker(
                [clat, clon],
                icon=folium.DivIcon(
                    html=f"<div style='font-size:10px;font-weight:700;color:#fff;"
                         f"text-shadow:0 0 3px #000,0 0 3px #000;line-height:1'>{stand}</div>",
                    icon_size=(36, 14), icon_anchor=(18, 7),
                ),
            ).add_to(m)

    # ── Build kiosk → stands lookup from elec sheet ───────────────────
    elec_df = df[df["meter_type"] == "Electrical"].copy()
    # We need kiosk number per stand — load from hierarchy if available,
    # but df doesn't carry kiosk_number. We'll re-derive from stand_map_color
    # by using the elec_df directly; kiosk info comes via load_kiosk_data
    # which is separate. For the map we attach stands via a pre-built lookup
    # passed in as kiosk_stands (built in the tab).
    # kiosks list items may carry a "stands" key injected by the tab.

    # ── Kiosk markers — small dot + label ────────────────────────────
    for k in kiosks:
        kname = k["name"].strip()
        stands_list = k.get("stands", [])
        total = len(stands_list)

        # Build compact unit list for popup
        inst_set = set(df[df["installed"] & (df["meter_type"] == "Electrical")]["stand"].tolist())
        amr_set  = set(df[df["amr"]       & (df["meter_type"] == "Electrical")]["stand"].tolist())

        chips = ""
        for s in stands_list:
            inst = s in inst_set
            amr  = s in amr_set
            if inst and amr:
                bg, fc = "#2E7D5222", "#2E7D52"
            elif inst:
                bg, fc = "#E6913822", "#B96E1E"
            else:
                bg, fc = "#3A506822", "#7a9ec4"
            chips += (f"<span style='display:inline-block;margin:2px;padding:1px 5px;"
                      f"border-radius:4px;font-size:10px;font-weight:600;"
                      f"background:{bg};color:{fc};border:1px solid {fc}55'>{s}</span>")

        installed_count = sum(1 for s in stands_list if s in inst_set)
        amr_count       = sum(1 for s in stands_list if s in amr_set)

        popup_html = (
            f"<div style='font-family:sans-serif;min-width:200px;max-width:320px'>"
            f"<div style='font-size:14px;font-weight:700;color:#152B45;margin-bottom:4px'>"
            f"\u26a1 Kiosk {kname}</div>"
            f"<div style='font-size:11px;color:#555;margin-bottom:8px'>"
            f"Meters: <b>{installed_count}/{total}</b> installed &nbsp;|&nbsp; "
            f"AMR: <b>{amr_count}/{total}</b></div>"
            f"<div style='font-size:9px;color:#888;margin-bottom:4px;text-transform:uppercase;"
            f"letter-spacing:.05em'>Units fed by this kiosk</div>"
            f"<div style='line-height:1.8'>{chips if chips else '<span style=\"color:#aaa\">No stands mapped yet</span>'}</div>"
            f"<div style='font-size:9px;color:#aaa;margin-top:6px'>"
            f"\u2705 installed + AMR &nbsp; \U0001f7e0 installed, AMR pending &nbsp; \U0001f535 not yet installed</div>"
            f"</div>"
        )

        # Small circle marker + tiny label — much less cluttered than a big div
        folium.CircleMarker(
            location=[k["lat"], k["lon"]],
            radius=10,
            color="#B96E1E",
            fill=True,
            fill_color="#E69138",
            fill_opacity=0.95,
            weight=2,
            tooltip=folium.Tooltip(
                f"<b>\u26a1 {kname}</b> &nbsp; {installed_count}/{total} installed",
                sticky=True
            ),
            popup=folium.Popup(popup_html, max_width=340),
        ).add_to(m)

        # Tiny text label above the dot
        folium.Marker(
            location=[k["lat"], k["lon"]],
            icon=folium.DivIcon(
                html=(f"<div style='font-size:10px;font-weight:700;color:#E69138;"
                      f"text-shadow:0 0 3px #000,0 0 3px #000;white-space:nowrap;"
                      f"margin-top:-22px;margin-left:12px'>{kname}</div>"),
                icon_size=(65, 16), icon_anchor=(0, 16),
            ),
        ).add_to(m)

    # ── Minisub markers — small square icon ───────────────────────────
    ms_serials = {"MS1": "82929702", "MS2": "82929684", "MS3": "71205556"}
    for ms in minisubs:
        msname = ms["name"].strip()
        serial_str = ms_serials.get(msname.upper(), "")
        folium.CircleMarker(
            location=[ms["lat"], ms["lon"]],
            radius=9,
            color="#5B86B3",
            fill=True,
            fill_color="#1F3F66",
            fill_opacity=0.95,
            weight=2.5,
            tooltip=folium.Tooltip(
                f"<b>\U0001f50c {msname}</b>{(' &nbsp; ' + serial_str) if serial_str else ''}",
                sticky=True
            ),
            popup=folium.Popup(
                f"<b>{msname}</b><br>Serial: {serial_str or '\u2014'}",
                max_width=180
            ),
        ).add_to(m)

        folium.Marker(
            location=[ms["lat"], ms["lon"]],
            icon=folium.DivIcon(
                html=(f"<div style='font-size:9px;font-weight:700;color:#5B86B3;"
                      f"text-shadow:0 0 3px #000,0 0 3px #000;white-space:nowrap;"
                      f"margin-top:-20px;margin-left:12px'>{msname}</div>"),
                icon_size=(40, 14), icon_anchor=(0, 14),
            ),
        ).add_to(m)

    folium.LayerControl().add_to(m)
    return m


def kml_center(polygons, kiosks, minisubs):
    """Compute map center from all coordinates."""
    all_lats, all_lons = [], []
    for p in polygons:
        for lat, lon in p["coords"]:
            all_lats.append(lat); all_lons.append(lon)
    for items in (kiosks, minisubs):
        for p in items:
            all_lats.append(p["lat"]); all_lons.append(p["lon"])
    if all_lats:
        return [sum(all_lats)/len(all_lats), sum(all_lons)/len(all_lons)]
    return [-34.054, 18.778]   # Sitari estate fallback


def build_amr_map(polygons, kiosks, df, amr_readings, faulty_df, center):
    """
    Folium map coloured by AMR import status.
    amr_readings: dict of serial → {reading_date, reading_value, ...}
    faulty_df:    DataFrame of faulty-but-not-replaced meters (stand column)
    """
    from datetime import datetime as _dt
    m = folium.Map(location=center, zoom_start=19, tiles=None, max_zoom=22)
    folium.TileLayer(
        tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        attr="Esri World Imagery", name="Satellite", overlay=False, control=True,
        max_zoom=22, max_native_zoom=19,
    ).add_to(m)
    folium.TileLayer(
        tiles="https://{s}.basemaps.cartocdn.com/dark_only_labels/{z}/{x}/{y}{r}.png",
        attr="(c) OpenStreetMap (c) CARTO", name="Labels",
        overlay=True, control=True, opacity=0.7,
    ).add_to(m)

    faulty_pending_stands = set(faulty_df["stand"].tolist()) if not faulty_df.empty else set()

    amr_colors = {
        "amr-green":  "#2E7D52",
        "amr-yellow": "#D4AC0D",
        "amr-orange": "#E67E22",
        "amr-red":    "#BD4B2C",
        "amr-never":  "#607080",
    }

    for poly in polygons:
        stand = poly["name"].strip()
        rows  = df[df["stand"] == stand]

        # Find the best AMR status for this stand (prefer worst = most attention needed)
        # Priority order: green (best/most recent) → yellow → orange → red → never (worst/no data)
        # We track the BEST status across all meters on this stand so that
        # any meter importing in the last 24h colours the stand green.
        badge_priority = ["amr-green", "amr-yellow", "amr-orange", "amr-red", "amr-never"]
        best_badge     = "amr-never"   # start at worst; upgrade toward green as we find readings
        last_reading_str = "—"
        last_value_str   = "—"
        last_serial      = "—"

        for _, r in rows.iterrows():
            serial = r.get("serial", "")
            if not serial:
                continue
            reading = amr_readings.get(serial)
            rd_iso  = reading["reading_date"] if reading else None
            _, _, badge = amr_status_info(rd_iso)
            if badge_priority.index(badge) < badge_priority.index(best_badge):
                best_badge = badge
            if rd_iso and str(rd_iso) not in ("nan","None",""):
                last_reading_str = str(rd_iso)[:16].replace("T"," ")
            if reading and reading.get("reading_value") is not None:
                mtype = r.get("meter_type","")
                val   = reading["reading_value"]
                last_value_str = f"{val:,.1f} {'L' if mtype=='Water' else 'kWh'}"
            last_serial = serial

        color   = amr_colors.get(best_badge, "#607080")
        is_faulty_pending = stand in faulty_pending_stands
        border_color      = "#EF4444" if is_faulty_pending else color
        border_weight     = 3 if is_faulty_pending else 1.5
        opacity = 0.85

        faulty_note = ""
        if is_faulty_pending:
            faulty_note = "<div style='margin:4px 0;padding:3px 8px;background:#EF444422;color:#EF4444;border:1px solid #EF444466;border-radius:5px;font-size:11px;font-weight:700'>⚠️ Faulty meter — not yet replaced</div>"

        popup_html = (
            f"<div style='font-family:sans-serif;min-width:200px'>"
            f"<div style='font-size:14px;font-weight:700;color:#152B45;margin-bottom:4px'>Stand {stand}</div>"
            f"<div style='display:inline-block;padding:2px 9px;border-radius:8px;font-size:11px;"
            f"font-weight:600;background:{color}22;color:{color};border:1px solid {color}88;margin-bottom:6px'>"
            f"{amr_status_info(amr_readings.get(last_serial,{}).get('reading_date') if last_serial != '—' else None)[0]}</div>"
            f"{faulty_note}"
            f"<table style='font-size:11px;width:100%;border-collapse:collapse'>"
            f"<tr><td style='color:#777;padding:2px 4px'>Last reading</td><td style='padding:2px 4px'>{last_reading_str}</td></tr>"
            f"<tr><td style='color:#777;padding:2px 4px'>Value</td><td style='padding:2px 4px'>{last_value_str}</td></tr>"
            f"<tr><td style='color:#777;padding:2px 4px'>Serial</td><td style='padding:2px 4px;font-family:monospace'>{last_serial}</td></tr>"
            f"</table>"
            f"<div style='font-size:9px;color:#aaa;margin-top:6px'>Stand ID: {stand}</div>"
            f"</div>"
        )

        folium.Polygon(
            locations=poly["coords"],
            color=border_color, weight=border_weight,
            fill=True, fill_color=color, fill_opacity=opacity,
            tooltip=folium.Tooltip(f"<b>Stand {stand}</b>", sticky=True),
            popup=folium.Popup(popup_html, max_width=280),
        ).add_to(m)

        # Stand label
        if poly["coords"]:
            clat = sum(c[0] for c in poly["coords"]) / len(poly["coords"])
            clon = sum(c[1] for c in poly["coords"]) / len(poly["coords"])
            folium.Marker(
                [clat, clon],
                icon=folium.DivIcon(
                    html=f"<div style='font-size:10px;font-weight:700;color:#fff;"
                         f"text-shadow:0 0 3px #000,0 0 3px #000;line-height:1'>{stand}</div>",
                    icon_size=(36, 14), icon_anchor=(18, 7),
                ),
            ).add_to(m)

    # Kiosk markers (lightweight)
    for k in kiosks:
        kname = k["name"].strip()
        folium.CircleMarker(
            location=[k["lat"], k["lon"]], radius=8,
            color="#B96E1E", fill=True, fill_color="#E69138", fill_opacity=0.95, weight=2,
            tooltip=folium.Tooltip(f"<b>\u26a1 {kname}</b>", sticky=True),
        ).add_to(m)
        folium.Marker(
            location=[k["lat"], k["lon"]],
            icon=folium.DivIcon(
                html=f"<div style='font-size:9px;font-weight:700;color:#E69138;"
                     f"text-shadow:0 0 3px #000;white-space:nowrap;margin-top:-20px;margin-left:12px'>{kname}</div>",
                icon_size=(60, 14), icon_anchor=(0, 14),
            ),
        ).add_to(m)

    folium.LayerControl().add_to(m)
    return m



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
    # Use a hash of the CSV content to guarantee a unique key even when the
    # same table function is called multiple times with identical row counts.
    import hashlib
    key_hash = hashlib.md5(csv).hexdigest()[:10]
    st.download_button(
        "Download as CSV", csv,
        file_name="meters_export.csv", mime="text/csv",
        key=f"dl_{columns[0]}_{key_hash}"
    )


def summary_counters(view_df, full_df, label):
    """Shows total counts for this category, plus a water/electrical split,
    and how many remain after the filters above are applied."""
    total_in_category = len(full_df)
    water_n = int((full_df["meter_type"] == "Water").sum())
    elec_n  = int((full_df["meter_type"] == "Electrical").sum())
    shown_n = len(view_df)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"Total {label}", total_in_category)
    c2.metric("Water", water_n)
    c3.metric("Electrical", elec_n)
    c4.metric("Matching current filters", shown_n)


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

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Total meter points", total)
c2.metric("Installed", installed_n, f"{round(installed_n/total*100)}% complete")
c3.metric("Outstanding", outstanding_n)
c4.metric("Due within 14 days", due_soon_n)
c5.metric("Overdue", overdue_n, delta_color="inverse")
faulty_n = int(df["faulty"].sum()) if "faulty" in df.columns else 0
faulty_pending_n = int(df[df["faulty"] & ~df["faulty_replaced"]]["stand"].count()) if "faulty" in df.columns else 0
c6.metric("Faulty meters", faulty_n, f"{faulty_pending_n} awaiting replacement", delta_color="inverse")

st.divider()

# ---------- Tabs ----------
tab_outstanding, tab_upcoming, tab_overdue, tab_installed, tab_calendar, tab_sections, tab_retic, tab_map, tab_faulty, tab_amr = st.tabs(
    ["🟦 Outstanding", "🟧 Upcoming", "🟥 Overdue", "🟩 Installed", "📅 Calendar", "📊 Sections", "⚡ Reticulation", "🗺️ Estate Map", "⚠️ Faulty Meters", "📡 AMR Live"]
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

    # Build faulty stand lookup for reticulation diagram
    faulty_stands_set = set(df[df["faulty"]]["stand"].tolist()) if "faulty" in df.columns else set()
    faulty_replaced_set = set(df[df["faulty"] & df["faulty_replaced"]]["stand"].tolist()) if "faulty" in df.columns else set()
    faulty_json = json.dumps(list(faulty_stands_set))
    faulty_replaced_json = json.dumps(list(faulty_replaced_set))

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
  .stand-chip.faulty {{
    background: #EF444422; color: #EF4444; border: 2px solid #EF444488;
    animation: faulty-pulse 1.8s ease-in-out infinite;
  }}
  .stand-chip.faulty-replaced {{
    background: #8B5CF622; color: #A78BFA; border: 1px solid #8B5CF688;
  }}
  @keyframes faulty-pulse {{
    0%, 100% {{ box-shadow: 0 0 0 0 #EF444433; }}
    50% {{ box-shadow: 0 0 0 3px #EF444422; }}
  }}
  .faulty-tag {{
    display: inline-block; font-size: 8px; padding: 0 4px; border-radius: 3px;
    font-weight: 700; margin-left: 2px; vertical-align: middle;
    background: #EF444422; color: #EF4444; border: 1px solid #EF444466;
  }}
  .faulty-tag.replaced {{
    background: #8B5CF622; color: #A78BFA; border: 1px solid #8B5CF666;
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
  /* Deadline badges on chips */
  .dl-badge {{
    display: inline-block; font-size: 8px; padding: 0 4px; border-radius: 3px;
    font-weight: 700; vertical-align: middle; margin-left: 2px; letter-spacing: .02em;
  }}
  .dl-overdue {{ background: #BD4B2C44; color: #e07060; border: 1px solid #BD4B2C66; }}
  .dl-due-soon {{ background: #E6913844; color: #d4902a; border: 1px solid #E6913866; }}
  .kiosk-deadline-row {{ display: flex; gap: 6px; margin-top: 4px; flex-wrap: wrap; }}
  .kiosk-dl-tag {{
    font-size: 9px; padding: 1px 7px; border-radius: 10px; font-weight: 600;
  }}
  .kiosk-dl-tag.overdue {{ background: #BD4B2C33; color: #e07060; border: 1px solid #BD4B2C55; }}
  .kiosk-dl-tag.due-soon {{ background: #E6913833; color: #d4902a; border: 1px solid #E6913855; }}
</style>
</head>
<body>

<div class="serial-popup" id="serialPopup"></div>

<div class="legend-bar">
  <strong style="color:#9FB0C2;font-size:10px;letter-spacing:.06em;">STAND STATUS:</strong>
  <div class="legend-item"><span class="legend-swatch" style="background:#3F7D5C33;border:1px solid #3F7D5C66"></span><span style="color:#6eb88a">Meter ✓ · AMR ✓</span></div>
  <div class="legend-item"><span class="legend-swatch" style="background:#E6913833;border:1px solid #E6913866"></span><span style="color:#d4902a">Meter ✓ · AMR pending</span></div>
  <div class="legend-item"><span class="legend-swatch" style="background:#1F3F6633;border:1px solid #334d6e"></span><span style="color:#7a9ec4">Meter not yet installed</span></div>
  <div class="legend-item"><span class="dl-badge dl-overdue" style="display:inline-block">OVR</span><span style="color:#e07060;margin-left:4px">Past Snag Date 4</span></div>
  <div class="legend-item"><span class="dl-badge dl-due-soon" style="display:inline-block">DUE</span><span style="color:#d4902a;margin-left:4px">Due within 14 days</span></div>
  <div class="legend-item"><span class="faulty-tag" style="display:inline-block">FLT</span><span style="color:#EF4444;margin-left:4px">Faulty — awaiting replacement</span></div>
  <div class="legend-item"><span class="faulty-tag replaced" style="display:inline-block">RPLC</span><span style="color:#A78BFA;margin-left:4px">Faulty — replaced</span></div>
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
const faultyStands = new Set({faulty_json});
const faultyReplacedStands = new Set({faulty_replaced_json});
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

function showPopup(e, stand, serial, inst, amr, deadline, days, dlstatus) {{
  const amrHtml = !inst
    ? `<span class="sp-pending">Meter not yet installed</span>`
    : amr
      ? `<span class="sp-amr-ok">✓ Commissioned</span>`
      : `<span class="sp-amr-miss">⚠ Pending</span>`;
  const serialHtml = serial
    ? `<span class="sp-val">${{serial}}</span>`
    : `<span class="sp-pending">No serial recorded</span>`;
  const deadlineHtml = deadline
    ? (dlstatus === 'overdue'
        ? `<span class="sp-amr-miss">${{deadline}} (${{Math.abs(days)}}d overdue)</span>`
        : dlstatus === 'due_soon'
          ? `<span style="color:#d4902a">${{deadline}} (${{days}}d remaining)</span>`
          : `<span class="sp-val">${{deadline}} (${{days}}d remaining)</span>`)
    : `<span class="sp-pending">—</span>`;

  popup.innerHTML = `
    <div class="sp-stand">Stand ${{stand}}</div>
    <div class="sp-row"><span class="sp-label">Meter serial</span>${{serialHtml}}</div>
    <div class="sp-row"><span class="sp-label">Status</span><span class="sp-val">${{inst ? 'Installed' : 'Not installed'}}</span></div>
    <div class="sp-row"><span class="sp-label">AMR</span>${{amrHtml}}</div>
    ${{!inst ? '<div class="sp-row"><span class="sp-label">Deadline</span>' + deadlineHtml + '</div>' : ''}}
  `;
  popup.classList.add('visible');
  const x = Math.min(e.clientX + 12, window.innerWidth - 230);
  const y = Math.min(e.clientY + 12, window.innerHeight - 160);
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
        ${{(k.overdue_stands && k.overdue_stands.length > 0) || (k.due_soon_stands && k.due_soon_stands.length > 0) ? `
        <div class="kiosk-deadline-row">
          ${{k.overdue_stands && k.overdue_stands.length > 0 ? '<span class="kiosk-dl-tag overdue">🟥 ' + k.overdue_stands.length + ' overdue</span>' : ''}}
          ${{k.due_soon_stands && k.due_soon_stands.length > 0 ? '<span class="kiosk-dl-tag due-soon">🟧 ' + k.due_soon_stands.length + ' due soon</span>' : ''}}
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
        const overduSet = new Set(k.overdue_stands || []);
        const dueSoonSet = new Set(k.due_soon_stands || []);
        const deadlineMap = k.stand_deadlines || {{}};
        const chipsHtml = k.stands.map(s => {{
          const inst = installedSet.has(s);
          const amr = amrSet.has(s);
          const serial = serialMap[s] || '';
          const isHighlight = highlightStands.has(s);
          const isOverdue = overduSet.has(s);
          const isDueSoon = dueSoonSet.has(s);
          const dlInfo = deadlineMap[s] || null;
          const isFaulty = faultyStands.has(s);
          const isFaultyReplaced = faultyReplacedStands.has(s);

          let cls = 'stand-chip pending';
          if (isFaulty && !isFaultyReplaced) cls = 'stand-chip faulty';
          else if (isFaulty && isFaultyReplaced) cls = 'stand-chip faulty-replaced';
          else if (inst && amr) cls = 'stand-chip';
          else if (inst) cls = 'stand-chip no-amr';
          if (isHighlight) cls += ' highlight';

          const dot = isFaulty
            ? (isFaultyReplaced
                ? `<span class="amr-dot" style="background:#8B5CF6"></span>`
                : `<span class="amr-dot" style="background:#EF4444"></span>`)
            : inst
              ? `<span class="amr-dot ${{amr ? 'ok' : 'missing'}}"></span>`
              : `<span class="amr-dot" style="background:#334d6e"></span>`;

          const faultyBadge = isFaulty
            ? (isFaultyReplaced
                ? `<span class="faulty-tag replaced">RPLC</span>`
                : `<span class="faulty-tag">FLT</span>`)
            : '';

          const dlBadge = !inst && isOverdue
            ? `<span class="dl-badge dl-overdue">OVR</span>`
            : !inst && isDueSoon
              ? `<span class="dl-badge dl-due-soon">DUE</span>`
              : '';

          const dlTitle = dlInfo ? ` · Deadline: ${{dlInfo.deadline}} (${{dlInfo.days_to < 0 ? Math.abs(dlInfo.days_to) + 'd overdue' : dlInfo.days_to + 'd remaining'}})` : '';
          const faultyTitle = isFaulty ? ` ⚠️ FAULTY${{isFaultyReplaced ? ' (replaced)' : ' - awaiting replacement'}}` : '';
          const title = inst
            ? `Stand ${{s}} · Serial: ${{serial || 'not recorded'}} · AMR: ${{amr ? '✓' : 'pending'}}${{faultyTitle}}`
            : `Stand ${{s}} · Not installed${{dlTitle}}${{faultyTitle}}`;

          return `<span class="${{cls}}" data-stand="${{s}}" data-serial="${{serial}}" data-inst="${{inst}}" data-amr="${{amr}}" data-deadline="${{dlInfo ? dlInfo.deadline : ''}}" data-days="${{dlInfo ? dlInfo.days_to : ''}}" data-dlstatus="${{dlInfo ? dlInfo.status : ''}}" title="${{title}}">${{dot}}${{s}}${{faultyBadge}}${{dlBadge}}</span>`;
        }}).join('');
        const amrMissing = k.installed - k.amr_count;
        const faultyInKiosk = k.stands ? k.stands.filter(s => faultyStands.has(s)).length : 0;
        const faultyReplacedInKiosk = k.stands ? k.stands.filter(s => faultyReplacedStands.has(s)).length : 0;
        detail.innerHTML = `
          <div style="font-size:9px;color:#5B86B3;margin-bottom:4px;">
            STANDS (${{k.stands.length}} in sheet · ${{k.planned}} planned)
            &nbsp;·&nbsp; <span class="amr-ok-count">AMR done: ${{k.amr_count}}</span>
            ${{amrMissing > 0 ? '&nbsp;·&nbsp; <span class="amr-miss-count">AMR pending: ' + amrMissing + '</span>' : ''}}
            ${{(k.overdue_stands||[]).length > 0 ? '&nbsp;·&nbsp; <span style="color:#e07060">🟥 ' + k.overdue_stands.length + ' overdue</span>' : ''}}
            ${{(k.due_soon_stands||[]).length > 0 ? '&nbsp;·&nbsp; <span style="color:#d4902a">🟧 ' + k.due_soon_stands.length + ' due soon</span>' : ''}}
            ${{faultyInKiosk > 0 ? '&nbsp;·&nbsp; <span style="color:#EF4444">⚠️ ' + faultyInKiosk + ' faulty (' + faultyReplacedInKiosk + ' replaced)</span>' : ''}}
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
              chip.dataset.amr === 'True',
              chip.dataset.deadline || '',
              chip.dataset.days !== '' ? parseInt(chip.dataset.days) : null,
              chip.dataset.dlstatus || ''
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

# =====================================================================
# ESTATE MAP TAB
# =====================================================================
with tab_map:
    st.subheader("\U0001f5fa\ufe0f Estate Map \u2014 Installation Status")

    # Auto-load KML from repo folder, fall back to upload
    auto_kml_path = find_kml_file()
    saved_kml, saved_kml_name = None, None

    if auto_kml_path:
        with open(auto_kml_path, "rb") as _f:
            saved_kml = _f.read()
        saved_kml_name = os.path.basename(auto_kml_path)

    uploaded_kml = st.file_uploader(
        "Upload a different KML file (optional — auto-loaded from repo if present)",
        type=["kml"], key="kml_uploader"
    )
    if uploaded_kml is not None:
        saved_kml = uploaded_kml.read()
        saved_kml_name = uploaded_kml.name

    if saved_kml is None:
        st.info(
            "No KML found. Commit your `EVG_Sitari.kml` into the repo folder alongside `app.py` "
            "and it will load automatically — or upload it using the button above."
        )
    else:
        st.caption(f"\U0001f4c2 Using **{saved_kml_name}**")

        try:
            polygons, kiosks, minisubs = parse_kml(saved_kml)
        except Exception as e:
            st.error(f"Could not parse KML: {e}")
            polygons, kiosks, minisubs = [], [], []

        if not polygons and not kiosks and not minisubs:
            st.warning("Nothing found in this KML — check folder names contain \'unit\', \'kiosk\', or \'minisub\'.")
        else:
            st.caption(
                f"\U0001f3e0 **{len(polygons)}** unit polygons \u00b7 "
                f"\u26a1 **{len(kiosks)}** kiosks \u00b7 "
                f"\U0001f50c **{len(minisubs)}** minisubs"
            )

            # ── Controls ────────────────────────────────────────────────
            col_a, col_b, col_c, col_d = st.columns(4)
            with col_a:
                map_meter_type = st.radio(
                    "Colour polygons by", ["Both meters", "Electrical only", "Water only"],
                    horizontal=True, key="map_meter_type"
                )
            with col_b:
                show_kiosks_map  = st.checkbox("Show kiosk labels", value=True, key="map_show_kiosks")
                show_minisubs    = st.checkbox("Show minisub markers", value=True, key="map_show_ms")
            with col_c:
                show_stand_labels = st.checkbox("Show stand number labels", value=True, key="map_stand_labels")
            with col_d:
                faulty_mode = st.checkbox(
                    "⚠️ Highlight faulty meters", value=False, key="map_faulty_mode",
                    help="Switch polygon colours to show fault status: red = faulty awaiting replacement, purple = replaced, dark = no fault."
                )

            map_df = df.copy()
            if map_meter_type == "Electrical only":
                map_df = df[df["meter_type"] == "Electrical"]
            elif map_meter_type == "Water only":
                map_df = df[df["meter_type"] == "Water"]

            # ── Legend ──────────────────────────────────────────────────
            if faulty_mode:
                legend_items = [
                    ("#EF4444", "Faulty \u2014 awaiting replacement"),
                    ("#8B5CF6", "Faulty \u2014 replaced"),
                    ("#1E3A2F", "No fault recorded"),
                ]
            else:
                legend_items = [
                    ("#2E7D52", "Meter \u2713 \u00b7 AMR \u2713"),
                    ("#E69138", "Meters \u2713 \u00b7 AMR pending"),
                    ("#5B86B3", "Partially installed"),
                    ("#BD4B2C", "Overdue"),
                    ("#D4AC0D", "Due soon"),
                    ("#3A5068", "On track"),
                    ("#607080", "No data"),
                ]
            lg = "<div style=\'display:flex;flex-wrap:wrap;gap:10px;padding:6px 0;font-size:12px;align-items:center;\'>"
            for color, label in legend_items:
                lg += (f"<div style=\'display:flex;align-items:center;gap:5px;\'>"
                       f"<span style=\'display:inline-block;width:13px;height:13px;border-radius:3px;"
                       f"background:{color};border:1px solid {color}99\'></span><span>{label}</span></div>")
            lg += "</div>"
            st.markdown(lg, unsafe_allow_html=True)

            # ── KPIs ────────────────────────────────────────────────────
            mapped_stands = {p["name"].strip() for p in polygons}
            matched = df[df["stand"].isin(mapped_stands)]
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Polygons on map", len(polygons))
            k2.metric("Matched to data", len(matched["stand"].unique()))
            k3.metric("Installed (matched)", int(matched["installed"].sum()))
            k4.metric("AMR done (matched)", int(matched["amr"].sum()))
            k5.metric("Kiosks / Minisubs", f"{len(kiosks)} / {len(minisubs)}")

            # ── Render map ──────────────────────────────────────────────
            center = kml_center(polygons, kiosks, minisubs)

            # Build kiosk → list of stands from elec sheet and attach to
            # each kiosk dict so the popup can show which units it feeds
            elec_sheet = df[df["meter_type"] == "Electrical"]
            # Re-load kiosk number per stand from the Excel file directly
            try:
                _xls = pd.ExcelFile(data_path)
                _edf = _xls.parse("Elec Meters")
                _edf.columns = [str(c).strip() for c in _edf.columns]
                _edf["stand_str"] = _edf["Stand Number"].astype(str).str.strip()
                kiosk_stands_map = (
                    _edf.groupby("Kiosk Number")["stand_str"]
                    .apply(lambda x: sorted(x.tolist()))
                    .to_dict()
                )
            except Exception:
                kiosk_stands_map = {}

            kiosks_with_stands = [
                {**k, "stands": kiosk_stands_map.get(k["name"].strip(), [])}
                for k in kiosks
            ]

            _kw = kiosks_with_stands if show_kiosks_map else []
            _ms = minisubs if show_minisubs else []
            _gh = _poly_hash(polygons, _kw, _ms)
            _dh = _df_hash(map_df)
            _fh = str(faulty_mode) + str(show_stand_labels)

            with st.spinner("Building map…"):
                map_html = cached_estate_map_html(
                    _gh, _dh + _fh, center, show_stand_labels, faulty_mode,
                    polygons, _kw, _ms, map_df
                )
                render_cached_map(map_html, height=660)

            st.caption(
                "Satellite imagery: Esri World Imagery. "
                "Click any house polygon to see meter and AMR status. "
                "Commit an updated KML to the repo to refresh the map."
            )

# =====================================================================
# FAULTY METERS TAB
# =====================================================================
with tab_faulty:
    st.subheader("⚠️ Faulty Meter Log")
    st.caption(
        "All meters flagged as faulty in the spreadsheet. "
        "Update **Faulty Meter**, **Faulty Replaced**, and **Replacement Date** "
        "columns in the Excel sheet and push to refresh."
    )

    faulty_df = df[df["faulty"]].copy() if "faulty" in df.columns else pd.DataFrame()

    if faulty_df.empty:
        st.success("No meters currently flagged as faulty.")
    else:
        # KPIs
        total_faulty   = len(faulty_df)
        replaced       = int(faulty_df["faulty_replaced"].sum())
        pending_repl   = total_faulty - replaced
        water_faulty   = int((faulty_df["meter_type"] == "Water").sum())
        elec_faulty    = int((faulty_df["meter_type"] == "Electrical").sum())

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Total faulty", total_faulty)
        k2.metric("Awaiting replacement", pending_repl)
        k3.metric("Already replaced", replaced)
        k4.metric("Water meters", water_faulty)
        k5.metric("Electrical meters", elec_faulty)

        st.divider()

        # Filters
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            f_type = st.multiselect(
                "Meter type", ["Water", "Electrical"],
                default=list(faulty_df["meter_type"].unique()),
                key="faulty_type"
            )
        with fc2:
            f_status = st.multiselect(
                "Replacement status",
                ["Awaiting replacement", "Replaced"],
                default=["Awaiting replacement", "Replaced"],
                key="faulty_status"
            )
        with fc3:
            f_section = st.multiselect(
                "Section (WBHO)",
                sorted(faulty_df["wbho_section"].dropna().unique()),
                key="faulty_section"
            )

        view = faulty_df[faulty_df["meter_type"].isin(f_type)].copy()

        status_map = []
        if "Awaiting replacement" in f_status:
            status_map.append(~view["faulty_replaced"])
        if "Replaced" in f_status:
            status_map.append(view["faulty_replaced"])
        if status_map:
            combined = status_map[0]
            for m in status_map[1:]:
                combined = combined | m
            view = view[combined]

        if f_section:
            view = view[view["wbho_section"].isin(f_section)]

        if view.empty:
            st.info("No faulty meters match the current filters.")
        else:
            # Build display table
            display = view[[
                "stand", "meter_type", "unit_type", "wbho_section",
                "serial", "commission_date", "faulty_replaced", "replacement_date", "manufacturer", "model"
            ]].copy()

            display["faulty_replaced"] = display["faulty_replaced"].map(
                {True: "✅ Replaced", False: "⏳ Awaiting replacement"}
            )
            display["commission_date"] = pd.to_datetime(display["commission_date"]).dt.strftime("%d %b %Y")
            display["replacement_date"] = pd.to_datetime(display["replacement_date"]).dt.strftime("%d %b %Y").fillna("—")

            display.columns = [
                "Stand", "Type", "Unit type", "Section",
                "Original serial", "Commissioned", "Replacement status", "Replaced on",
                "Manufacturer", "Model"
            ]
            display = display.sort_values(["Replacement status", "Stand"])

            st.dataframe(display, use_container_width=True, hide_index=True,
                         column_config={
                             "Replacement status": st.column_config.TextColumn("Status", width="medium"),
                         })

            import hashlib
            csv = display.to_csv(index=False).encode("utf-8")
            st.download_button(
                "Download faulty meter list as CSV", csv,
                file_name="faulty_meters.csv", mime="text/csv",
                key=f"dl_faulty_{hashlib.md5(csv).hexdigest()[:8]}"
            )

        st.divider()
        st.markdown("##### Replacement history timeline")
        replaced_with_date = faulty_df[faulty_df["faulty_replaced"] & faulty_df["replacement_date"].notna()]
        if replaced_with_date.empty:
            st.caption("No replacement dates recorded yet.")
        else:
            for _, row in replaced_with_date.sort_values("replacement_date").iterrows():
                st.markdown(
                    f"**Stand {row['stand']}** ({row['meter_type']}) — "
                    f"original serial `{row['serial'] or '—'}` — "
                    f"replaced {row['replacement_date'].strftime('%d %b %Y')}"
                )


# =====================================================================
# AMR LIVE TAB
# =====================================================================
with tab_amr:
    st.subheader("📡 AMR Live — Meter Reading Status")
    st.caption(
        "Matches meter serials from your commissioning spreadsheet against the latest "
        "hourly CSV file on the SFTP server. Auto-refreshes every hour on the next page load."
    )

    # ── SFTP Configuration ────────────────────────────────────────────
    with st.expander("⚙️ SFTP Configuration", expanded="sftp_configured" not in st.session_state):
        st.markdown(
            "Enter your SFTP credentials below, or add them to `.streamlit/secrets.toml` "
            "on Streamlit Cloud so they're not stored in plain text.\n\n"
            "```toml\n[sftp]\nhost = \"sftp.example.com\"\nport = 22\n"
            "username = \"your_user\"\npassword = \"your_pass\"\n"
            "directory = \"/path/to/csv/files\"\n```"
        )

        # Try to load from st.secrets first
        try:
            sftp_defaults = st.secrets.get("sftp", {})
        except Exception:
            sftp_defaults = {}

        sc1, sc2 = st.columns(2)
        with sc1:
            sftp_host = st.text_input("SFTP Host", value=sftp_defaults.get("host", ""), key="sftp_host")
            sftp_user = st.text_input("Username",  value=sftp_defaults.get("username", ""), key="sftp_user")
            sftp_dir  = st.text_input("Directory", value=sftp_defaults.get("directory", ""), placeholder="/data/readings", key="sftp_dir")
        with sc2:
            sftp_port = st.number_input("Port", value=int(sftp_defaults.get("port", 22)), min_value=1, max_value=65535, key="sftp_port")
            sftp_pass = st.text_input("Password", value=sftp_defaults.get("password", ""), type="password", key="sftp_pass")

        sftp_ready = all([sftp_host, sftp_user, sftp_pass, sftp_dir])
        if sftp_ready:
            st.session_state["sftp_configured"] = True

    # ── Load readings: SFTP or uploaded CSV ──────────────────────────
    amr_readings = {}
    amr_source   = None
    amr_file_ts  = None
    amr_error    = None

    # Manual CSV upload as fallback / testing
    uploaded_amr = st.file_uploader(
        "Or upload a CSV file directly (for testing without SFTP)",
        type=["csv"], key="amr_csv_upload"
    )

    col_fetch, col_bulk, col_hours, col_last = st.columns([1, 1.2, 0.8, 2])
    with col_fetch:
        fetch_clicked = st.button("🔄 Fetch latest from SFTP", disabled=not sftp_ready, key="amr_fetch_btn")
    with col_hours:
        bulk_hours = st.number_input("Hours back", min_value=1, max_value=168, value=24,
                                     key="amr_bulk_hours", label_visibility="collapsed",
                                     help="How many hours back to fetch for history population")
    with col_bulk:
        bulk_clicked = st.button(f"📥 Fetch last {bulk_hours}h (history)", disabled=not sftp_ready,
                                 key="amr_bulk_btn",
                                 help="Downloads ALL CSV files in the SFTP directory within the selected window and stores them in the history database.")

    if uploaded_amr is not None:
        raw = uploaded_amr.read()
        fname_ts = _parse_csv_filename_ts(uploaded_amr.name)
        amr_readings = parse_amr_csv(raw, fname_ts)
        amr_source   = f"Uploaded: {uploaded_amr.name}"
        amr_file_ts  = fname_ts
        try:
            db_upsert_readings(amr_readings, fname_ts.isoformat() if fname_ts else None)
        except Exception:
            pass
        save_amr_cache({"readings": amr_readings, "source": amr_source,
                        "file_ts": fname_ts.isoformat() if fname_ts else None})

    elif bulk_clicked and sftp_ready:
        progress_bar  = st.progress(0, text="Connecting to SFTP…")
        status_text   = st.empty()

        def _progress(current, total, fname):
            pct = int(current / total * 100)
            progress_bar.progress(pct, text=f"Fetching file {current}/{total}: {fname}")
            status_text.caption(f"Processing: **{fname}**")

        files_done, new_rows, latest_rdgs, bulk_err = fetch_amr_bulk_history(
            sftp_host, int(sftp_port), sftp_user, sftp_pass, sftp_dir,
            hours=int(bulk_hours), progress_cb=_progress
        )
        progress_bar.empty()
        status_text.empty()

        if bulk_err:
            st.error(f"SFTP error: {bulk_err}")
        else:
            total_rows, distinct_serials, mn, mx = db_stats()
            st.success(
                f"✅ Fetched **{files_done}** files from the last {bulk_hours}h · "
                f"**{new_rows}** new readings added to history DB · "
                f"DB now holds **{total_rows:,}** readings across **{distinct_serials}** serials "
                f"({mn[:10] if mn else '—'} → {mx[:10] if mx else '—'})"
            )
            if latest_rdgs:
                amr_readings = latest_rdgs
                amr_source   = f"SFTP bulk ({bulk_hours}h)"
                amr_file_ts  = None
                save_amr_cache({"readings": amr_readings, "source": amr_source, "file_ts": None})

    elif fetch_clicked and sftp_ready:
        import time
        cache_bust = int(time.time() // 3600)
        with st.spinner("Connecting to SFTP…"):
            amr_readings, fname, amr_file_ts, amr_error = fetch_amr_from_sftp(
                sftp_host, int(sftp_port), sftp_user, sftp_pass, sftp_dir, cache_bust
            )
        if amr_error:
            st.error(f"SFTP error: {amr_error}")
        else:
            amr_source = f"SFTP: {fname}"
            save_amr_cache({"readings": amr_readings, "source": amr_source,
                            "file_ts": amr_file_ts.isoformat() if amr_file_ts else None})
            st.success(f"Fetched **{fname}** — {len(amr_readings)} meter readings loaded.")

    else:
        # Load from local cache (persists across restarts on self-hosted)
        cached = load_amr_cache()
        if cached and cached.get("readings"):
            amr_readings = cached["readings"]
            amr_source   = cached.get("source", "Cached data")
            raw_ts        = cached.get("file_ts")
            amr_file_ts  = datetime.fromisoformat(raw_ts) if raw_ts else None

    # Auto-fetch on page load if SFTP configured and no data yet
    if not amr_readings and sftp_ready and not fetch_clicked and not bulk_clicked:
        import time
        cache_bust = int(time.time() // 3600)
        with st.spinner("Auto-fetching latest readings from SFTP…"):
            amr_readings, fname, amr_file_ts, amr_error = fetch_amr_from_sftp(
                sftp_host, int(sftp_port), sftp_user, sftp_pass, sftp_dir, cache_bust
            )
        if not amr_error and amr_readings:
            amr_source = f"SFTP: {fname}"
            save_amr_cache({"readings": amr_readings, "source": amr_source,
                            "file_ts": amr_file_ts.isoformat() if amr_file_ts else None})

    with col_last:
        if amr_source:
            ts_str = amr_file_ts.strftime("%d %b %Y %H:%M") if amr_file_ts else "unknown time"
            st.caption(f"📂 {amr_source} · file timestamp: **{ts_str}** · {len(amr_readings)} serials in CSV")
        elif not sftp_ready:
            st.caption("Configure SFTP above or upload a CSV to see readings.")

    if not amr_readings:
        st.info("No AMR data loaded yet. Configure SFTP and click 'Fetch latest', or upload a CSV file.")
    else:
        st.divider()

        # ── Build enriched AMR table from df + readings ───────────────
        amr_df = df[df["installed"] & df["serial"].str.len().gt(0)].copy()
        amr_df = amr_df.drop_duplicates(subset=["stand", "meter_type"])

        rows = []
        for _, r in amr_df.iterrows():
            serial = r["serial"]
            reading = amr_readings.get(serial)
            rd_iso   = reading["reading_date"]   if reading else None
            rd_value = reading["reading_value"]  if reading else None
            low_bat  = int(reading["low_battery"]) if reading else 0
            label, color, badge = amr_status_info(rd_iso)
            rows.append({
                "stand":        r["stand"],
                "meter_type":   r["meter_type"],
                "unit_type":    r["unit_type"],
                "wbho_section": r["wbho_section"],
                "serial":       serial,
                "amr_ok":       r["amr"],
                "last_reading": rd_iso,
                "reading_value":rd_value,
                "low_battery":  low_bat,
                "status_label": label,
                "status_color": color,
                "status_badge": badge,
            })

        amr_table = pd.DataFrame(rows)

        # Status buckets
        green  = amr_table[amr_table["status_badge"] == "amr-green"]
        yellow = amr_table[amr_table["status_badge"] == "amr-yellow"]
        orange = amr_table[amr_table["status_badge"] == "amr-orange"]
        red    = amr_table[amr_table["status_badge"] == "amr-red"]
        never  = amr_table[amr_table["status_badge"] == "amr-never"]
        total_amr = len(amr_table)

        # ── KPIs ──────────────────────────────────────────────────────
        a1, a2, a3, a4, a5, a6 = st.columns(6)
        a1.metric("AMR-enabled meters", total_amr)
        a2.metric("🟢 Last 24h",   len(green),  f"{round(len(green)/total_amr*100) if total_amr else 0}%")
        a3.metric("🟡 1–3 days",   len(yellow), f"{round(len(yellow)/total_amr*100) if total_amr else 0}%")
        a4.metric("🟠 4–7 days",   len(orange), f"{round(len(orange)/total_amr*100) if total_amr else 0}%")
        a5.metric("🔴 7+ days",    len(red),    f"{round(len(red)/total_amr*100) if total_amr else 0}%")
        a6.metric("⚫ No reading", len(never),  f"{round(len(never)/total_amr*100) if total_amr else 0}%")

        # Low battery alert
        low_bat_meters = amr_table[amr_table["low_battery"] == 1]
        if not low_bat_meters.empty:
            st.warning(
                f"🔋 **{len(low_bat_meters)} meters reporting low battery** — "
                f"{', '.join(low_bat_meters['stand'].head(8).tolist())}"
                + (" and more…" if len(low_bat_meters) > 8 else "")
            )

        st.divider()

        # ── Filters ───────────────────────────────────────────────────
        af1, af2, af3, af4 = st.columns(4)
        with af1:
            a_type = st.multiselect("Meter type", ["Water", "Electrical"],
                                    default=["Water", "Electrical"], key="amr_type")
        with af2:
            a_status = st.multiselect(
                "Status", ["🟢 Last 24h", "🟡 1–3 days", "🟠 4–7 days", "🔴 7+ days", "⚫ No reading"],
                default=["🟢 Last 24h", "🟡 1–3 days", "🟠 4–7 days", "🔴 7+ days", "⚫ No reading"],
                key="amr_status_filter"
            )
        with af3:
            a_section = st.multiselect("Section (WBHO)",
                                       sorted(amr_table["wbho_section"].dropna().unique()),
                                       key="amr_section")
        with af4:
            a_serial_search = st.text_input("Search serial", key="amr_serial_search")

        status_badge_map = {
            "🟢 Last 24h":   "amr-green",
            "🟡 1–3 days":   "amr-yellow",
            "🟠 4–7 days":   "amr-orange",
            "🔴 7+ days":    "amr-red",
            "⚫ No reading": "amr-never",
        }
        selected_badges = [status_badge_map[s] for s in a_status]

        view = amr_table[
            amr_table["meter_type"].isin(a_type) &
            amr_table["status_badge"].isin(selected_badges)
        ]
        if a_section:
            view = view[view["wbho_section"].isin(a_section)]
        if a_serial_search:
            view = view[view["serial"].str.contains(a_serial_search.strip(), case=False, na=False)]

        # ── Visual status grid (HTML component) ───────────────────────
        # ── Faulty cross-reference ────────────────────────────────────
        faulty_pending_df = df[df["faulty"] & ~df["faulty_replaced"]] if "faulty" in df.columns else pd.DataFrame()
        faulty_pending_stands = set(faulty_pending_df["stand"].tolist())

        # Flag stale/missing meters that are also in the faulty list
        stale_and_faulty = amr_table[
            amr_table["status_badge"].isin(["amr-orange","amr-red","amr-never"]) &
            amr_table["stand"].isin(faulty_pending_stands)
        ]
        if not stale_and_faulty.empty:
            st.warning(
                f"⚠️ **{len(stale_and_faulty)} stale/missing meters are also flagged as faulty (not yet replaced)** — "
                f"likely cause of missing readings: "
                + ", ".join(f"Stand {r['stand']}" for _, r in stale_and_faulty.iterrows())
            )

        st.divider()

        # ── Site map coloured by AMR status ───────────────────────────
        st.markdown("##### AMR status — site map")

        # Rebuild kiosk/minisub/polygon lookup (reuse KML already loaded if available)
        _kml_path = find_kml_file()
        _amr_polygons, _amr_kiosks = [], []
        if _kml_path:
            try:
                with open(_kml_path, "rb") as _f:
                    _kml_data = _f.read()
                _amr_polygons, _amr_kiosks, _ = parse_kml(_kml_data)
            except Exception:
                pass

        if not _amr_polygons:
            st.info("No KML found — push your `EVG_Sitari.kml` to the repo to see the site map here. Showing summary grid instead.")
            # Fallback chip grid
            color_map_fb = {"amr-green":"#2E7D52","amr-yellow":"#D4AC0D","amr-orange":"#E67E22","amr-red":"#BD4B2C","amr-never":"#607080"}
            fg = "<div style='display:flex;flex-wrap:wrap;gap:5px;padding:8px 0'>"
            for _, r in amr_table.sort_values(["meter_type","stand"]).iterrows():
                c = color_map_fb.get(r["status_badge"],"#607080")
                in_f = r["meter_type"] in a_type and r["status_badge"] in selected_badges
                op = "1.0" if in_f else "0.15"
                stand_fb = r["stand"]
                status_fb = r["status_label"]
                fg += (f"<span title='Stand {stand_fb} · {status_fb}' "
                       f"style='display:inline-block;width:36px;height:24px;border-radius:4px;"
                       f"background:{c};opacity:{op};font-size:8px;color:#fff;font-weight:700;"
                       f"text-align:center;line-height:24px'>{stand_fb}</span>")
            fg += "</div>"
            st.markdown(fg, unsafe_allow_html=True)
        else:
            _amr_map_df = df.copy()
            if a_type and len(a_type) < 2:
                _amr_map_df = df[df["meter_type"].isin(a_type)]

            _center = kml_center(_amr_polygons, _amr_kiosks, [])
            _gh  = _poly_hash(_amr_polygons, _amr_kiosks)
            _dh  = _df_hash(_amr_map_df)
            _ah  = _amr_hash(amr_readings)
            _fph = str(sorted(faulty_pending_df["stand"].tolist())) if not faulty_pending_df.empty else ""

            with st.spinner("Building AMR map…"):
                amr_map_html = cached_amr_map_html(
                    _gh, _dh, _ah, _fph, _center,
                    _amr_polygons, _amr_kiosks, _amr_map_df, amr_readings, faulty_pending_df
                )
                render_cached_map(amr_map_html, height=620)

            # Stand selection via dropdown only (no st_folium click overhead)
            auto_stand = None

            # AMR map legend
            leg_items = [
                ("#2E7D52","🟢 Last 24h"),("#D4AC0D","🟡 1–3 days"),
                ("#E67E22","🟠 4–7 days"),("#BD4B2C","🔴 7+ days"),
                ("#607080","⚫ No reading"),("#EF4444","⚠️ Faulty pending"),
            ]
            leg = "<div style='display:flex;flex-wrap:wrap;gap:12px;font-size:11px;margin-top:4px'>"
            for col, lbl in leg_items:
                leg += f"<span><span style='display:inline-block;width:12px;height:12px;border-radius:3px;background:{col};vertical-align:middle;margin-right:4px'></span>{lbl}</span>"
            leg += "</div>"
            st.markdown(leg, unsafe_allow_html=True)

            if auto_stand:
                st.session_state["amr_selected_stand"] = auto_stand

        st.divider()

        # ── Stand history ──────────────────────────────────────────────
        st.markdown("##### Stand history")

        # Build stand options (only installed meters with a serial)
        stand_options = sorted(amr_table["stand"].unique().tolist())
        default_stand = st.session_state.get("amr_selected_stand", stand_options[0] if stand_options else None)
        if default_stand not in stand_options:
            default_stand = stand_options[0] if stand_options else None

        sel_col, info_col = st.columns([1, 2])
        with sel_col:
            selected_stand = st.selectbox(
                "Select a stand to view its reading history",
                options=stand_options,
                index=stand_options.index(default_stand) if default_stand in stand_options else 0,
                key="amr_stand_select"
            )
            if selected_stand:
                st.session_state["amr_selected_stand"] = selected_stand

        if selected_stand:
            stand_rows  = amr_table[amr_table["stand"] == selected_stand]
            serials     = stand_rows["serial"].dropna().unique().tolist()
            serials     = [s for s in serials if s and s not in ("nan","None","")]

            with info_col:
                if not stand_rows.empty:
                    r0 = stand_rows.iloc[0]
                    _, color_h, badge_h = amr_status_info(r0["last_reading"])
                    st.markdown(
                        f"**Stand {selected_stand}** · {r0['unit_type']} · {r0['wbho_section']}  \n"
                        f"Serial: `{r0['serial']}` · "
                        f"<span style='background:{color_h}22;color:{color_h};padding:1px 7px;"
                        f"border-radius:8px;font-size:12px;font-weight:600'>{r0['status_label']}</span>"
                        + (" &nbsp; ⚠️ **Faulty**" if selected_stand in faulty_pending_stands else ""),
                        unsafe_allow_html=True
                    )

            if not serials:
                st.info(f"Stand {selected_stand} has no meter serial recorded — can't look up history.")
            else:
                import plotly.graph_objects as go
                import plotly.express as px

                # DB stats
                total_rows, distinct_serials, min_dt, max_dt = db_stats()
                st.caption(
                    f"History database: **{total_rows:,}** readings · **{distinct_serials}** serials · "
                    f"from {min_dt[:10] if min_dt else '—'} to {max_dt[:10] if max_dt else '—'}"
                )

                all_hist = []
                for serial in serials:
                    hist_df = db_get_history(serial)
                    if not hist_df.empty:
                        hist_df["serial"] = serial
                        hist_df["meter_type"] = stand_rows[stand_rows["serial"] == serial]["meter_type"].iloc[0] if len(stand_rows[stand_rows["serial"] == serial]) else "Unknown"
                        all_hist.append(hist_df)

                if not all_hist:
                    st.info(f"No historical readings in the database yet for stand {selected_stand}. History builds up as readings are fetched from SFTP.")
                else:
                    hist_full = pd.concat(all_hist, ignore_index=True).sort_values("reading_date")

                    for hist_sub in all_hist:
                        mtype   = hist_sub["meter_type"].iloc[0]
                        serial  = hist_sub["serial"].iloc[0]
                        unit    = "Litres" if mtype == "Water" else "kWh"
                        n_reads = len(hist_sub)
                        n_low   = int(hist_sub["low_battery"].sum())

                        # Calculate consumption (delta between readings)
                        # Pre-compute derived columns before the tab split
                        # so both tabs can access them
                        hist_sub = hist_sub.copy()
                        hist_sub["consumption"] = hist_sub["reading_value"].diff().clip(lower=0)
                        hist_sub["gap_hours"]   = hist_sub["reading_date"].diff().dt.total_seconds() / 3600

                        tab_chart, tab_raw = st.tabs([f"📈 {mtype} chart", f"📋 Raw readings ({n_reads})"])

                        with tab_chart:
                            if len(hist_sub) < 2:
                                st.info("Need at least 2 readings to draw a chart. Keep fetching data hourly.")
                            else:
                                fig = go.Figure()
                                # Cumulative meter reading
                                fig.add_trace(go.Scatter(
                                    x=hist_sub["reading_date"],
                                    y=hist_sub["reading_value"],
                                    mode="lines+markers",
                                    name=f"Meter reading ({unit})",
                                    line=dict(color="#2E7D52" if mtype == "Water" else "#5B86B3", width=2),
                                    marker=dict(size=4),
                                ))
                                # Consumption bars
                                fig.add_trace(go.Bar(
                                    x=hist_sub["reading_date"],
                                    y=hist_sub["consumption"],
                                    name=f"Consumption per interval ({unit})",
                                    marker_color="#E69138",
                                    opacity=0.5,
                                    yaxis="y2",
                                ))
                                # Mark gaps (missed readings) — points where gap > 3 hours
                                gaps = hist_sub[hist_sub["gap_hours"] > 3]
                                if not gaps.empty:
                                    fig.add_trace(go.Scatter(
                                        x=gaps["reading_date"],
                                        y=gaps["reading_value"],
                                        mode="markers",
                                        name="Gap in readings",
                                        marker=dict(color="#BD4B2C", size=10, symbol="x"),
                                    ))
                                # Low battery markers
                                lbat = hist_sub[hist_sub["low_battery"] == 1]
                                if not lbat.empty:
                                    fig.add_trace(go.Scatter(
                                        x=lbat["reading_date"],
                                        y=lbat["reading_value"],
                                        mode="markers",
                                        name="Low battery",
                                        marker=dict(color="#FFF176", size=8, symbol="triangle-down"),
                                    ))

                                fig.update_layout(
                                    title=f"Stand {selected_stand} · {mtype} · Serial {serial}",
                                    xaxis_title="Date / Time",
                                    yaxis_title=f"Cumulative reading ({unit})",
                                    yaxis2=dict(title=f"Consumption ({unit})", overlaying="y", side="right", showgrid=False),
                                    legend=dict(orientation="h", yanchor="bottom", y=1.02),
                                    plot_bgcolor="#0E1117", paper_bgcolor="#0E1117",
                                    font=dict(color="#E0E0E0"),
                                    xaxis=dict(gridcolor="#1E2D3D"),
                                    yaxis=dict(gridcolor="#1E2D3D"),
                                    height=420,
                                )
                                st.plotly_chart(fig, use_container_width=True)
                                st.caption(
                                    f"🔴 X markers = gap >3h between readings · "
                                    f"🟡 ▽ = low battery reported · "
                                    f"Orange bars = consumption per interval"
                                )

                        with tab_raw:
                            raw_disp = hist_sub[["reading_date","reading_value","low_battery","gap_hours"]].copy()
                            raw_disp.columns = ["Date / Time", f"Reading ({unit})", "Low battery", "Gap since prev (h)"]
                            raw_disp["Date / Time"] = raw_disp["Date / Time"].dt.strftime("%d %b %Y %H:%M")
                            raw_disp["Low battery"] = raw_disp["Low battery"].map({1:"🔋", 0:""})
                            raw_disp["Gap since prev (h)"] = raw_disp["Gap since prev (h)"].fillna(0).round(1)
                            st.dataframe(raw_disp.sort_values("Date / Time", ascending=False),
                                         use_container_width=True, hide_index=True)
                            import hashlib
                            rc = raw_disp.to_csv(index=False).encode("utf-8")
                            st.download_button(f"Download {mtype} history", rc,
                                               file_name=f"stand_{selected_stand}_{mtype.lower()}_history.csv",
                                               key=f"dl_hist_{selected_stand}_{mtype}_{hashlib.md5(rc).hexdigest()[:6]}")

        st.divider()

        # ── Detail table ──────────────────────────────────────────────
        st.markdown(f"##### Meter detail ({len(view)} meters)")

        def fmt_reading(v, mtype):
            if v is None:
                return "—"
            if mtype == "Water":
                return f"{v:,.1f} L"
            return f"{v:,.3f} kWh"

        display_rows = []
        for _, r in view.sort_values(["status_badge", "stand"]).iterrows():
            display_rows.append({
                "Stand":        r["stand"],
                "Type":         r["meter_type"],
                "Unit type":    r["unit_type"],
                "Section":      r["wbho_section"],
                "Serial":       r["serial"],
                "Last reading": str(r["last_reading"])[:16].replace("T", " ") if r["last_reading"] and str(r["last_reading"]) not in ("nan", "None", "") else "—",
                "Reading":      fmt_reading(r["reading_value"], r["meter_type"]),
                "Status":       r["status_label"],
                "Faulty":       "⚠️" if r["stand"] in faulty_pending_stands else "",
                "Low battery":  "🔋" if r["low_battery"] else "",
            })

        disp_df = pd.DataFrame(display_rows)
        st.dataframe(disp_df, use_container_width=True, hide_index=True,
                     column_config={
                         "Status": st.column_config.TextColumn("Status", width="small"),
                         "Faulty": st.column_config.TextColumn("Fault", width="small"),
                         "Low battery": st.column_config.TextColumn("Bat", width="small"),
                     })

        import hashlib
        csv_out = disp_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download AMR status as CSV", csv_out,
            file_name="amr_status.csv", mime="text/csv",
            key=f"dl_amr_{hashlib.md5(csv_out).hexdigest()[:8]}"
        )

        st.divider()

        # ── Per-type breakdown ────────────────────────────────────────
        st.markdown("##### Breakdown by meter type")
        for mtype in ["Water", "Electrical"]:
            subset = amr_table[amr_table["meter_type"] == mtype]
            if subset.empty:
                continue
            counts = subset["status_badge"].value_counts()
            total  = len(subset)
            parts  = []
            for badge, emoji, label in [
                ("amr-green",  "🟢", "Last 24h"),
                ("amr-yellow", "🟡", "1–3 days"),
                ("amr-orange", "🟠", "4–7 days"),
                ("amr-red",    "🔴", "7+ days"),
                ("amr-never",  "⚫", "No reading"),
            ]:
                n = counts.get(badge, 0)
                if n:
                    parts.append(f"{emoji} **{n}** {label}")
            low = len(subset[subset["low_battery"] == 1])
            bat_str = f" · 🔋 **{low}** low battery" if low else ""
            faulty_stale = len(subset[subset["stand"].isin(faulty_pending_stands) & subset["status_badge"].isin(["amr-orange","amr-red","amr-never"])])
            fault_str = f" · ⚠️ **{faulty_stale}** stale+faulty" if faulty_stale else ""
            importing_pct = round(counts.get("amr-green", 0) / total * 100) if total else 0
            st.markdown(
                f"**{mtype}** ({total} meters · {importing_pct}% importing) &nbsp; "
                + " &ensp; ".join(parts) + bat_str + fault_str
            )

        st.caption(
            f"Auto-refreshes every hour on the next page load (st.cache_data TTL=3600s). "
            f"Click '🔄 Fetch latest from SFTP' to force an immediate refresh. "
            f"History database stores all readings and persists between sessions."
        )