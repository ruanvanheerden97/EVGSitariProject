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

# ---------- Mobile detection ----------
def _detect_mobile():
    """Best-effort mobile detection from the browser User-Agent."""
    try:
        ua = st.context.headers.get("User-Agent", "")
        return bool(re.search(r"Mobi|Android|iPhone|iPad", ua, re.IGNORECASE))
    except Exception:
        return False

if "is_mobile" not in st.session_state:
    st.session_state["is_mobile"] = _detect_mobile()
IS_MOBILE = st.session_state["is_mobile"]


def metric_row(items, desktop_cols=None):
    """
    Render a row of st.metric tiles that adapts to screen size:
    max 4 per row on desktop (wrapping into extra rows), 2 per row on mobile.
    The cap keeps tiles wide enough that labels and values never truncate.
    items: list of tuples (label, value) or (label, value, delta) or
           (label, value, delta, delta_color).
    """
    per_row = 2 if IS_MOBILE else min(desktop_cols or len(items), 4)
    for i in range(0, len(items), per_row):
        chunk = items[i:i + per_row]
        cols = st.columns(per_row)
        for col, item in zip(cols, chunk):
            label, value = item[0], item[1]
            delta = item[2] if len(item) > 2 else None
            dcol  = item[3] if len(item) > 3 else "normal"
            col.metric(label, value, delta, delta_color=dcol)


# ---------- Styling ----------
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; max-width: 1300px;}
.stMetric, div[data-testid="stMetric"] {
  background: #FBF9F3 !important;
  border: 1px solid #DCD6C4;
  border-radius: 10px;
  padding: 7px 10px;
}
div[data-testid="stMetric"] * {
  color: #152B45 !important;
}
div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] * {
  font-size: 11px !important;
  text-transform: uppercase;
  letter-spacing: .03em;
  color: #3E5066 !important;
  white-space: normal !important;      /* wrap instead of truncating */
  overflow: visible !important;
  text-overflow: clip !important;
  line-height: 1.25 !important;
}
div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] * {
  color: #152B45 !important;
  font-size: 1.45rem !important;       /* compact value — fits narrow tiles */
  line-height: 1.15 !important;
  white-space: normal !important;
  overflow: visible !important;
}
div[data-testid="stMetricDelta"], div[data-testid="stMetricDelta"] * {
  color: #3F7D5C !important;
  fill: #3F7D5C !important;
  font-size: 11px !important;
  white-space: normal !important;      /* long delta notes wrap too */
  overflow: visible !important;
  text-overflow: clip !important;
  line-height: 1.2 !important;
}
h1, h2, h3 {color: #152B45;}

/* ── Mobile responsiveness ─────────────────────────────────────── */
@media (max-width: 740px) {
  .block-container {padding: 0.6rem 0.7rem 2rem 0.7rem !important;}
  h1 {font-size: 1.35rem !important;}
  h2 {font-size: 1.1rem !important;}
  h3 {font-size: 1.0rem !important;}
  div[data-testid="stMetric"] {padding: 6px 8px !important;}
  div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] * {
    font-size: 1.15rem !important;
  }
  div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] * {
    font-size: 10px !important;
  }
  /* keep tab bar on one scrollable line */
  div[data-testid="stTabs"] button {padding: 6px 8px !important; font-size: 12px !important;}
  div[role="tablist"] {overflow-x: auto !important; flex-wrap: nowrap !important;}
  /* header logos smaller */
  .app-header-logo img {max-height: 46px !important;}
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>
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

WATER_SHEET_CANDIDATES    = ["FS Water", "Water meters", "Water Meters"]
ELEC_SHEET_CANDIDATES     = ["FS Elec", "Elec Meters", "Electrical Meters"]
APRT_ELEC_CANDIDATES      = ["Aprt Elec", "Apartment Elec"]
APRT_WATER_CANDIDATES     = ["Aprt Water", "Apartment Water"]
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
CSV_FILENAME_RE = re.compile(r"^[A-Z]+_(\d{8})_(\d{6})\.csv$", re.IGNORECASE)

# ── Supabase / PostgreSQL history database ────────────────────────────────────
try:
    import psycopg2
    import psycopg2.extras
    PSYCOPG2_AVAILABLE = True
except ImportError:
    PSYCOPG2_AVAILABLE = False


def _get_db_url():
    """Read the Supabase connection URL from st.secrets or env."""
    try:
        return st.secrets["db"]["url"]
    except Exception:
        return os.environ.get("DB_URL")


def _db_conn():
    """Open a Supabase PostgreSQL connection and ensure the table exists."""
    if not PSYCOPG2_AVAILABLE:
        raise RuntimeError("psycopg2 not installed — add psycopg2-binary to requirements.txt")
    url = _get_db_url()
    if not url:
        raise RuntimeError("No database URL. Add [db] url to .streamlit/secrets.toml")
    # The ! in passwords needs encoding, so pass params explicitly
    # Also try appending sslmode for Supabase
    try:
        conn = psycopg2.connect(url, sslmode="require", connect_timeout=10)
    except Exception:
        conn = psycopg2.connect(url, connect_timeout=10)
    # Ensure table exists (fast no-op if already present)
    if "amr_table_created" not in st.session_state:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS amr_readings (
                    serial        TEXT    NOT NULL,
                    reading_date  TEXT    NOT NULL,
                    reading_value DOUBLE PRECISION,
                    low_battery   INTEGER DEFAULT 0,
                    file_ts       TEXT,
                    PRIMARY KEY (serial, reading_date)
                )
            """)
            cur.execute("CREATE INDEX IF NOT EXISTS idx_amr_serial  ON amr_readings(serial)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_amr_date    ON amr_readings(reading_date)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_amr_file_ts ON amr_readings(file_ts)")
        conn.commit()
        st.session_state["amr_table_created"] = True
    return conn


def db_upsert_readings(readings_dict, file_ts_str):
    """Insert new readings. Existing (serial, reading_date) pairs are silently skipped."""
    if not PSYCOPG2_AVAILABLE or not _get_db_url():
        return 0
    rows = [
        (serial, v["reading_date"], v["reading_value"], v["low_battery"], file_ts_str)
        for serial, v in readings_dict.items()
        if v.get("reading_date")
    ]
    if not rows:
        return 0
    conn = _db_conn()
    with conn.cursor() as cur:
        psycopg2.extras.execute_values(
            cur,
            """INSERT INTO amr_readings (serial, reading_date, reading_value, low_battery, file_ts)
               VALUES %s
               ON CONFLICT (serial, reading_date) DO NOTHING""",
            rows,
            page_size=500,
        )
    conn.commit()
    conn.close()
    return len(rows)


def db_get_history(serial):
    """Return DataFrame of all readings for a serial, sorted ascending by date."""
    if not PSYCOPG2_AVAILABLE or not _get_db_url():
        return pd.DataFrame(columns=["reading_date", "reading_value", "low_battery"])
    try:
        conn = _db_conn()
        df = pd.read_sql_query(
            "SELECT reading_date, reading_value, low_battery FROM amr_readings "
            "WHERE serial = %s ORDER BY reading_date ASC",
            conn, params=(serial,)
        )
        conn.close()
        if not df.empty:
            df["reading_date"] = pd.to_datetime(df["reading_date"], errors="coerce")
        return df
    except Exception:
        return pd.DataFrame(columns=["reading_date", "reading_value", "low_battery"])


@st.cache_data(ttl=30, show_spinner=False)
def db_stats():
    """Return (total_rows, distinct_serials, min_date, max_date) from DB."""
    if not PSYCOPG2_AVAILABLE or not _get_db_url():
        return (0, 0, None, None)
    try:
        conn = _db_conn()
        with conn.cursor() as cur:
            cur.execute(
                "SELECT COUNT(*), COUNT(DISTINCT serial), "
                "MIN(reading_date), MAX(reading_date) FROM amr_readings"
            )
            row = cur.fetchone()
        conn.close()
        return row
    except Exception:
        return (0, 0, None, None)


def get_latest_file_ts():
    """Return the most recent file_ts in the DB as a datetime, or None if empty."""
    if not PSYCOPG2_AVAILABLE or not _get_db_url():
        return None
    try:
        conn = _db_conn()
        with conn.cursor() as cur:
            cur.execute(
                "SELECT MAX(file_ts) FROM amr_readings WHERE file_ts IS NOT NULL"
            )
            row = cur.fetchone()
        conn.close()
        if row and row[0]:
            return datetime.fromisoformat(row[0])
    except Exception:
        pass
    return None


def load_latest_from_db():
    """
    Reconstruct amr_readings dict (serial → {reading_date, reading_value, low_battery})
    from the DB by picking the most recent reading per serial.
    Used to restore the live display after a redeploy without re-fetching SFTP.
    """
    if not PSYCOPG2_AVAILABLE or not _get_db_url():
        return {}
    try:
        conn = _db_conn()
        df = pd.read_sql_query(
            """SELECT DISTINCT ON (serial)
                   serial, reading_date, reading_value, low_battery, file_ts
               FROM amr_readings
               ORDER BY serial, reading_date DESC""",
            conn
        )
        conn.close()
        result = {}
        for _, row in df.iterrows():
            result[row["serial"]] = {
                "reading_date":  row["reading_date"],
                "reading_value": row["reading_value"],
                "low_battery":   int(row["low_battery"] or 0),
                "file_ts":       row["file_ts"],
            }
        return result
    except Exception:
        return {}


# ── Stock / inventory (freestanding project) ─────────────────────────────────
STOCK_SEED = [
    # item_key, label, qty, reorder_level
    ("water_box",         "Water meter boxes (only)",              15, 5),
    ("water_meter",       "Water meters · 20mm Axioma",            10, 5),
    ("elec_programmed",   "Hexing electricity meters · programmed", 42, 10),
    ("elec_unprogrammed", "Hexing electricity meters · unprogrammed", 30, 10),
]


def _db_stock_init(conn):
    """Ensure stock tables exist and seed initial counts if empty."""
    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS stock_items (
                item_key      TEXT PRIMARY KEY,
                label         TEXT,
                baseline_qty  INTEGER NOT NULL,
                baseline_date TEXT NOT NULL,
                reorder_level INTEGER DEFAULT 5
            )""")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS stock_adjustments (
                id         SERIAL PRIMARY KEY,
                item_key   TEXT NOT NULL,
                qty_change INTEGER NOT NULL,
                note       TEXT,
                created_at TEXT NOT NULL
            )""")
        cur.execute("SELECT COUNT(*) FROM stock_items")
        if cur.fetchone()[0] == 0:
            now_iso = datetime.now().isoformat()
            for key, label, qty, reorder in STOCK_SEED:
                cur.execute(
                    "INSERT INTO stock_items (item_key, label, baseline_qty, baseline_date, reorder_level) "
                    "VALUES (%s,%s,%s,%s,%s) ON CONFLICT (item_key) DO NOTHING",
                    (key, label, qty, now_iso, reorder))
    conn.commit()


def db_stock_get():
    """Return {item_key: {label, baseline_qty, baseline_date, reorder_level, adjustments}}."""
    if not PSYCOPG2_AVAILABLE or not _get_db_url():
        return {}
    try:
        conn = _db_conn()
        _db_stock_init(conn)
        out = {}
        with conn.cursor() as cur:
            cur.execute("SELECT item_key, label, baseline_qty, baseline_date, reorder_level FROM stock_items")
            for key, label, bqty, bdate, reorder in cur.fetchall():
                cur2 = conn.cursor()
                cur2.execute(
                    "SELECT COALESCE(SUM(qty_change),0) FROM stock_adjustments "
                    "WHERE item_key=%s AND created_at > %s", (key, bdate))
                adj = cur2.fetchone()[0] or 0
                cur2.close()
                out[key] = {"label": label, "baseline_qty": bqty,
                            "baseline_date": bdate, "reorder_level": reorder,
                            "adjustments": int(adj)}
        conn.close()
        return out
    except Exception:
        return {}


def db_stock_adjust(item_key, qty_change, note=""):
    """Record a manual stock movement (+ received, − issued/correction)."""
    if not PSYCOPG2_AVAILABLE or not _get_db_url():
        return False
    try:
        conn = _db_conn()
        _db_stock_init(conn)
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO stock_adjustments (item_key, qty_change, note, created_at) VALUES (%s,%s,%s,%s)",
                (item_key, int(qty_change), note, datetime.now().isoformat()))
        conn.commit(); conn.close()
        return True
    except Exception:
        return False


def db_stock_set_baseline(item_key, qty, reorder_level=None):
    """Stocktake: set a new counted quantity as of now (resets auto-deduction from today)."""
    if not PSYCOPG2_AVAILABLE or not _get_db_url():
        return False
    try:
        conn = _db_conn()
        _db_stock_init(conn)
        with conn.cursor() as cur:
            if reorder_level is not None:
                cur.execute(
                    "UPDATE stock_items SET baseline_qty=%s, baseline_date=%s, reorder_level=%s WHERE item_key=%s",
                    (int(qty), datetime.now().isoformat(), int(reorder_level), item_key))
            else:
                cur.execute(
                    "UPDATE stock_items SET baseline_qty=%s, baseline_date=%s WHERE item_key=%s",
                    (int(qty), datetime.now().isoformat(), item_key))
        conn.commit(); conn.close()
        return True
    except Exception:
        return False


@st.cache_data(ttl=300, show_spinner=False)
def db_get_consumption(start_iso, end_iso, tol_hours=3):
    """
    Timestamp-ALIGNED consumption per serial, for fair parent-vs-children
    comparison. Every meter's usage is taken between its readings closest
    to the SAME two anchor times (window start and window end), within
    ±tol_hours. A meter without a reading near both anchors is treated as
    having no usable data for the window.

    Returns {serial: {delta, first, last, t_first, t_last, span_hours, n}}.
    """
    if not PSYCOPG2_AVAILABLE or not _get_db_url():
        return {}
    try:
        t0 = pd.Timestamp(start_iso)
        t1 = pd.Timestamp(end_iso)
        tol = pd.Timedelta(hours=tol_hours)

        conn = _db_conn()
        df = pd.read_sql_query(
            """SELECT serial, reading_date, reading_value
               FROM amr_readings
               WHERE reading_date >= %s AND reading_date <= %s
                 AND reading_value IS NOT NULL
               ORDER BY serial, reading_date""",
            conn, params=((t0 - tol).isoformat(), (t1 + tol).isoformat())
        )
        conn.close()
        if df.empty:
            return {}
        df["reading_dt"] = pd.to_datetime(df["reading_date"], errors="coerce")
        df = df[df["reading_dt"].notna()]

        result = {}
        for serial, grp in df.groupby("serial"):
            grp = grp.sort_values("reading_dt").reset_index(drop=True)
            # Reading closest to each anchor, within tolerance
            d0 = (grp["reading_dt"] - t0).abs()
            d1 = (grp["reading_dt"] - t1).abs()
            i0 = int(d0.idxmin())
            i1 = int(d1.idxmin())
            if d0.loc[i0] > tol or d1.loc[i1] > tol or i1 <= i0:
                continue
            v0 = float(grp["reading_value"].iloc[i0])
            v1 = float(grp["reading_value"].iloc[i1])
            span = (grp["reading_dt"].iloc[i1] - grp["reading_dt"].iloc[i0]).total_seconds() / 3600
            result[str(serial)] = {
                "delta": max(v1 - v0, 0.0),
                "first": v0, "last": v1,
                "t_first": grp["reading_dt"].iloc[i0].isoformat(),
                "t_last":  grp["reading_dt"].iloc[i1].isoformat(),
                "span_hours": round(span, 1),
                "n": len(grp),
            }
        return result
    except Exception:
        return {}


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


def fetch_amr_bulk_history(host, port, username, password, directory,
                           hours=24, since_dt=None, progress_cb=None):
    """
    Connect to SFTP, download CSV files within the time window, parse, and upsert into DB.
    Returns (files_processed, new_readings_inserted, latest_readings_dict, error_str).
    DB failures are logged but don't prevent file counting or amr_readings capture.
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

        cutoff = since_dt if since_dt else (datetime.now() - timedelta(hours=hours))

        def ts_key(name):
            ts = _parse_csv_filename_ts(name)
            return ts if ts else datetime.min

        in_window = sorted(
            [f for f in csv_files if ts_key(f) >= cutoff],
            key=ts_key, reverse=True
        )

        if not in_window:
            # No files in window — return the most recent file anyway as current snapshot
            all_sorted = sorted(csv_files, key=ts_key, reverse=True)
            if all_sorted:
                fname = all_sorted[0]
                file_ts = ts_key(fname)
                buf = io.BytesIO()
                sftp.getfo(directory.rstrip("/") + "/" + fname, buf)
                transport.close()
                readings = parse_amr_csv(buf.getvalue(), file_ts)
                try:
                    db_upsert_readings(readings, file_ts.isoformat() if file_ts else None)
                except Exception:
                    pass
                return 1, 0, readings, None
            transport.close()
            return 0, 0, {}, f"No CSV files found since {cutoff.strftime('%d %b %Y %H:%M')}"

        files_done  = 0
        total_new   = 0
        latest_readings = {}
        db_errors   = []

        for i, fname in enumerate(in_window):
            if progress_cb:
                progress_cb(i + 1, len(in_window), fname)

            file_ts = ts_key(fname)
            remote_path = directory.rstrip("/") + "/" + fname
            try:
                buf = io.BytesIO()
                sftp.getfo(remote_path, buf)
                readings = parse_amr_csv(buf.getvalue(), file_ts)
                files_done += 1
                # Always capture the most recent file's readings for the live display
                if i == 0:
                    latest_readings = readings
                # Try to persist to DB — failure here doesn't discard the file
                try:
                    n = db_upsert_readings(readings, file_ts.isoformat() if file_ts else None)
                    total_new += n
                except Exception as db_err:
                    db_errors.append(str(db_err))
            except Exception:
                continue

        transport.close()

        error_msg = None
        if db_errors and files_done > 0:
            error_msg = f"DB write errors ({len(db_errors)}): {db_errors[0]}"

        return files_done, total_new, latest_readings, error_msg

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

# =====================================================================
# SHARED: stand history renderer (graphs + raw readings for a stand)
# =====================================================================
# ── Balancing visual helpers ──────────────────────────────────────────────────
BAL_COLORS = {
    "parent":  "#5B86B3",   # steel blue
    "child":   "#2E7D52",   # green — measured children
    "est":     "#8B5CF6",   # purple — estimated / unmetered
    "loss_ok": "#3F7D5C",   # small technical loss
    "loss_hi": "#BD4B2C",   # high loss — investigate
    "anom":    "#E67E22",   # children exceed parent
}


def build_balance_sankey(parent_label, parent_val, children, est_val=0.0, height=320):
    """
    Energy-flow Sankey: parent → children (+ estimated + loss).
    children: list of (label, value, kind) where kind ∈ {'child','est'}.
    parent_val may be None → virtual parent (sum of children), labelled as such.
    Returns a plotly Figure.
    """
    import plotly.graph_objects as go

    kids = [(l, v, k) for l, v, k in children if v and v > 0.01]
    child_sum = sum(v for _, v, _ in kids)
    virtual = parent_val is None
    p = child_sum + (est_val or 0) if virtual else parent_val

    labels = [f"{parent_label}" + (" (virtual)" if virtual else "") + f"  {p:,.1f} kWh"]
    node_colors = [BAL_COLORS["parent"]]
    sources, targets, values, link_colors = [], [], [], []

    for lbl, val, kind in kids:
        labels.append(f"{lbl}  {val:,.1f}")
        node_colors.append(BAL_COLORS["est"] if kind == "est" else BAL_COLORS["child"])
        sources.append(0); targets.append(len(labels) - 1); values.append(val)
        link_colors.append("rgba(139,92,246,.45)" if kind == "est" else "rgba(46,125,82,.45)")

    if est_val and est_val > 0.01:
        labels.append(f"Unmetered (est.)  {est_val:,.1f}")
        node_colors.append(BAL_COLORS["est"])
        sources.append(0); targets.append(len(labels) - 1); values.append(est_val)
        link_colors.append("rgba(139,92,246,.55)")

    if not virtual and p is not None:
        residual = p - child_sum - (est_val or 0)
        if residual > 0.01:
            pct = residual / p * 100 if p else 0
            lc = BAL_COLORS["loss_hi"] if pct > 10 else BAL_COLORS["loss_ok"]
            labels.append(f"Loss  {residual:,.1f} ({pct:.1f}%)")
            node_colors.append(lc)
            sources.append(0); targets.append(len(labels) - 1); values.append(residual)
            link_colors.append("rgba(189,75,44,.5)" if pct > 10 else "rgba(63,125,92,.35)")
        elif residual < -0.01:
            labels.append(f"⚠ Children exceed parent by {abs(residual):,.1f}")
            node_colors.append(BAL_COLORS["anom"])
            sources.append(0); targets.append(len(labels) - 1); values.append(0.01)
            link_colors.append("rgba(230,126,34,.5)")

    fig = go.Figure(go.Sankey(
        arrangement="snap",
        node=dict(label=labels, color=node_colors, pad=12, thickness=16,
                  line=dict(color="#0e1117", width=0.5)),
        link=dict(source=sources, target=targets, value=values, color=link_colors),
    ))
    fig.update_layout(
        height=height, margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="#0E1117", font=dict(color="#E0E0E0", size=11, family="monospace"),
    )
    return fig


def build_db_bullet_chart(rows, unit="kWh", height_per_row=34):
    """
    Horizontal stacked bars: measured children (green) + estimate (purple) per DB,
    with a diamond marker at the parent check-meter value.
    rows: list of dicts {label, parent (float|None), measured, est, missing:int}
    """
    import plotly.graph_objects as go

    labels   = [r["label"] for r in rows][::-1]
    measured = [r["measured"] for r in rows][::-1]
    ests     = [r["est"] for r in rows][::-1]
    parents  = [r["parent"] for r in rows][::-1]
    missing  = [r["missing"] for r in rows][::-1]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=labels, x=measured, orientation="h", name="Children measured",
        marker_color=BAL_COLORS["child"],
        hovertemplate="%{y}<br>Measured: %{x:,.1f} " + unit + "<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        y=labels, x=ests, orientation="h", name="Estimated (unmetered)",
        marker_color=BAL_COLORS["est"],
        customdata=missing,
        hovertemplate="%{y}<br>Est.: %{x:,.1f} " + unit + " across %{customdata} meter(s)<extra></extra>",
    ))
    p_y = [l for l, p in zip(labels, parents) if p is not None]
    p_x = [p for p in parents if p is not None]
    if p_x:
        fig.add_trace(go.Scatter(
            y=p_y, x=p_x, mode="markers", name="Check meter (parent)",
            marker=dict(symbol="diamond", size=11, color=BAL_COLORS["parent"],
                        line=dict(color="#fff", width=1)),
            hovertemplate="%{y}<br>Check meter: %{x:,.1f} " + unit + "<extra></extra>",
        ))
    fig.update_layout(
        barmode="stack",
        height=max(180, height_per_row * len(rows) + 90),
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="#0E1117", paper_bgcolor="#0E1117",
        font=dict(color="#E0E0E0", size=11),
        xaxis=dict(title=unit, gridcolor="#1E2D3D"),
        yaxis=dict(gridcolor="#1E2D3D"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    return fig


def render_stand_history(stand, df_all, key_prefix="hist"):
    """Show usage charts + raw readings for every meter (all types) on a stand."""
    import plotly.graph_objects as go
    import hashlib as _hl

    stand_meters = df_all[(df_all["stand"] == stand) & df_all["serial"].str.len().gt(0)]
    stand_meters = stand_meters.drop_duplicates(subset=["meter_type"])
    if stand_meters.empty:
        st.info(f"No meters with serials found for {stand}.")
        return

    for _, r in stand_meters.iterrows():
        serial = r["serial"]
        mtype  = r["meter_type"]
        unit   = "kWh" if mtype == "Electrical" else "Litres"
        old_serial = str(r.get("old_serial", "") or "")
        if old_serial and old_serial not in ("nan", "None"):
            st.caption(f"🔁 **{mtype}** meter was replaced — showing new meter `{serial}` "
                       f"(previous meter `{old_serial}` below).")
        hist_df = db_get_history(serial)
        if hist_df.empty:
            st.caption(f"**{mtype}** · serial `{serial}` — no readings in history yet"
                       + (" (new meter — history builds up from its first import)." if old_serial and old_serial not in ("nan","None") else "."))
            if old_serial and old_serial not in ("nan", "None"):
                _old_hist = db_get_history(old_serial)
                if not _old_hist.empty:
                    with st.expander(f"📜 Previous {mtype} meter `{old_serial}` — {len(_old_hist)} readings"):
                        _od = _old_hist.copy()
                        _od["reading_date"] = _od["reading_date"].dt.strftime("%d %b %Y %H:%M")
                        st.dataframe(_od.sort_values("reading_date", ascending=False),
                                     width='stretch', hide_index=True, height=240)
            continue
        hist_df = hist_df.copy()
        hist_df["consumption"] = hist_df["reading_value"].diff().clip(lower=0)
        hist_df["gap_hours"]   = hist_df["reading_date"].diff().dt.total_seconds() / 3600

        t_chart, t_raw = st.tabs([f"📈 {mtype} ({len(hist_df)} readings)", f"📋 {mtype} raw"])
        with t_chart:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=hist_df["reading_date"], y=hist_df["reading_value"],
                mode="lines+markers", name=f"Reading ({unit})",
                line=dict(color="#5B86B3" if mtype == "Electrical" else "#2E7D52", width=2),
                marker=dict(size=4)))
            fig.add_trace(go.Bar(x=hist_df["reading_date"], y=hist_df["consumption"],
                name="Usage/interval", marker_color="#E69138", opacity=0.5, yaxis="y2"))
            gaps = hist_df[hist_df["gap_hours"] > 3]
            if not gaps.empty:
                fig.add_trace(go.Scatter(x=gaps["reading_date"], y=gaps["reading_value"],
                    mode="markers", name="Gap >3h",
                    marker=dict(color="#BD4B2C", size=10, symbol="x")))
            lbat = hist_df[hist_df["low_battery"] == 1]
            if not lbat.empty:
                fig.add_trace(go.Scatter(x=lbat["reading_date"], y=lbat["reading_value"],
                    mode="markers", name="Low battery",
                    marker=dict(color="#FFF176", size=8, symbol="triangle-down")))
            fig.update_layout(
                title=f"{stand} · {mtype} · {serial}",
                yaxis_title=unit,
                yaxis2=dict(title=f"Usage ({unit})", overlaying="y", side="right", showgrid=False),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                plot_bgcolor="#0E1117", paper_bgcolor="#0E1117",
                font=dict(color="#E0E0E0"), xaxis=dict(gridcolor="#1E2D3D"),
                yaxis=dict(gridcolor="#1E2D3D"), height=340, margin=dict(t=60, b=20))
            st.plotly_chart(fig, width='stretch',
                            key=f"{key_prefix}_fig_{stand}_{mtype}")
        with t_raw:
            raw = hist_df[["reading_date","reading_value","low_battery","gap_hours"]].copy()
            raw.columns = ["Date/Time", f"Reading ({unit})", "Low bat", "Gap (h)"]
            raw["Date/Time"] = raw["Date/Time"].dt.strftime("%d %b %Y %H:%M")
            raw["Low bat"]   = raw["Low bat"].map({1: "🔋", 0: ""})
            raw["Gap (h)"]   = raw["Gap (h)"].fillna(0).round(1)
            st.dataframe(raw.sort_values("Date/Time", ascending=False),
                         width='stretch', hide_index=True, height=260)
            rc = raw.to_csv(index=False).encode("utf-8")
            st.download_button(f"⬇️ {mtype} history", rc,
                file_name=f"{stand}_{mtype.lower().replace(' ','_')}_history.csv",
                key=f"{key_prefix}_dl_{stand}_{mtype}_{_hl.md5(rc).hexdigest()[:6]}")

        # Previous meter (replaced) — its readings remain in the DB under the old serial
        if old_serial and old_serial not in ("nan", "None"):
            _old_hist = db_get_history(old_serial)
            if not _old_hist.empty:
                with st.expander(f"📜 Previous {mtype} meter `{old_serial}` — {len(_old_hist)} readings (before replacement)"):
                    _od = _old_hist.copy()
                    _od["reading_date"] = _od["reading_date"].dt.strftime("%d %b %Y %H:%M")
                    st.dataframe(_od.sort_values("reading_date", ascending=False),
                                 width='stretch', hide_index=True, height=240)


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
def load_data(file_path, _mtime, site_type="freestanding"):
    """Load and normalise meter data. site_type: 'freestanding' or 'apartments'."""
    xls = pd.ExcelFile(file_path)

    if site_type == "apartments":
        return _load_apartment_data(xls)

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

        # Replacement meter serial: flexible column match — any header containing
        # 'serial' plus 'replacement'/'new'. When Faulty Replaced is ticked and a
        # replacement serial is present, it becomes the ACTIVE serial (used for
        # AMR matching, maps, history); the original moves to old_serial.
        _repl_col = next(
            (c for c in wdf.columns
             if "serial" in c.lower() and ("replace" in c.lower() or "new" in c.lower())),
            None)
        out["old_serial"] = ""
        if _repl_col:
            _repl = fmt_serial(wdf[_repl_col])
            _use  = out["faulty_replaced"] & (_repl.str.len() > 0)
            out.loc[_use, "old_serial"] = out.loc[_use, "serial"]
            out.loc[_use, "serial"]     = _repl[_use]

        out["eb_port"] = ""
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
        # EB Port: AMR device port allocated to this meter. A port number with
        # AMR Installed=False means the AMR is wired but the meter is still OFF
        # (building site, power not on) — no import expected yet.
        _eb_col = next((c for c in edf.columns
                        if "eb" in c.lower() and "port" in c.lower()), None)
        if _eb_col:
            def _fmt_port(v):
                if pd.isna(v):
                    return ""
                s = str(v).strip()
                if s in ("", "nan", "None"):
                    return ""
                try:
                    return str(int(float(s)))   # 3.0 → "3"
                except (ValueError, TypeError):
                    return s                     # e.g. "P3" stays as-is
            out["eb_port"] = edf[_eb_col].apply(_fmt_port)
        else:
            out["eb_port"] = ""
        _repl_col_e = next(
            (c for c in edf.columns
             if "serial" in c.lower() and ("replace" in c.lower() or "new" in c.lower())),
            None)
        out["old_serial"] = ""
        if _repl_col_e:
            _repl_e = fmt_serial(edf[_repl_col_e])
            _use_e  = out["faulty_replaced"] & (_repl_e.str.len() > 0)
            out.loc[_use_e, "old_serial"] = out.loc[_use_e, "serial"]
            out.loc[_use_e, "serial"]     = _repl_e[_use_e]
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


def _load_apartment_data(xls):
    """
    Normalise Aprt Elec + Aprt Water sheets into the same schema as freestanding data.
    Key differences:
    - No commission date → installed = serial present
    - No snag dates → no deadline tracking
    - Apartment Water has two serials per stand (cold + hot)
    - Hierarchy uses Apartment Block + Parent DB instead of WBHO Subsection + Kiosk
    """
    records = []
    today = pd.Timestamp("today")

    # ── Apartment Electrical ──────────────────────────────────────────
    elec_name = find_sheet(xls, APRT_ELEC_CANDIDATES)
    if elec_name:
        edf = xls.parse(elec_name)
        edf.columns = [str(c).strip() for c in edf.columns]
        out = pd.DataFrame()
        out["stand"]          = edf["Stand"].astype(str).str.strip()
        out["unit_type"]      = edf.get("Apartment Block", pd.Series(["Apartment"]*len(edf))).astype(str).str.strip()
        out["wbho_section"]   = edf.get("Parent DB",       pd.Series(["Unknown"  ]*len(edf))).astype(str).str.strip()
        out["manufacturer"]   = ""
        out["model"]          = ""
        out["serial"]         = edf["Elect meter serial"].astype(str).str.strip().replace("nan","")
        out["commission_date"]= pd.NaT
        out["amr"]            = edf["AMR installed"].fillna(False).astype(bool)
        out["deadline"]       = pd.NaT
        out["faulty"]         = False
        out["faulty_replaced"]= False
        out["old_serial"]     = ""
        out["replacement_date"]= pd.NaT
        _eb_col_a = next((c for c in edf.columns
                          if "eb" in c.lower() and "port" in c.lower()), None)
        if _eb_col_a:
            def _fmt_port_a(v):
                if pd.isna(v):
                    return ""
                s = str(v).strip()
                if s in ("", "nan", "None"):
                    return ""
                try:
                    return str(int(float(s)))
                except (ValueError, TypeError):
                    return s
            out["eb_port"] = edf[_eb_col_a].apply(_fmt_port_a)
        else:
            out["eb_port"] = ""
        out["meter_type"]     = "Electrical"
        out["installed"]      = out["serial"].str.len().gt(0) & out["serial"].ne("nan")
        out = out[out["stand"].notna() & (out["stand"] != "") & (out["stand"].str.lower() != "nan")]
        records.append(out)

    # ── Apartment Water (cold + hot — two rows per stand) ─────────────
    water_name = find_sheet(xls, APRT_WATER_CANDIDATES)
    if water_name:
        wdf = xls.parse(water_name)
        wdf.columns = [str(c).strip() for c in wdf.columns]
        block_col = wdf.get("Apartment Block", pd.Series(["Apartment"]*len(wdf))).astype(str).str.strip()

        for meter_subtype, serial_col in [("Water (Cold)", "Cold water Serial"),
                                          ("Water (Hot)",  "Hot water serial")]:
            if serial_col not in wdf.columns:
                continue
            out = pd.DataFrame()
            out["stand"]           = wdf["Stand"].astype(str).str.strip()
            out["unit_type"]       = block_col
            out["wbho_section"]    = block_col          # no DB info for water
            out["manufacturer"]    = ""
            out["model"]           = ""
            out["serial"]          = wdf[serial_col].astype(str).str.strip().replace("nan","")
            out["commission_date"] = pd.NaT
            out["amr"]             = wdf["AMR installed"].fillna(False).astype(bool)
            out["deadline"]        = pd.NaT
            out["faulty"]          = False
            out["faulty_replaced"] = False
            out["old_serial"]      = ""
            out["replacement_date"]= pd.NaT
            out["eb_port"]         = ""
            out["meter_type"]      = meter_subtype
            out["installed"]       = out["serial"].str.len().gt(0) & out["serial"].ne("nan")
            out = out[out["stand"].notna() & (out["stand"] != "") & (out["stand"].str.lower() != "nan")]
            records.append(out)

    if not records:
        return pd.DataFrame()

    df = pd.concat(records, ignore_index=True)

    # Status: apartments have no deadlines, so just installed/not
    def status_row(r):
        return "Installed" if r["installed"] else "On track"

    df["status"]           = df.apply(status_row, axis=1)
    df["days_to_deadline"] = None
    return df


@st.cache_data(show_spinner=False)
def load_aprt_reticulation(file_path, _mtime):
    """
    Build the 3-level apartment hierarchy from the Aprt Elec sheet:
        Minisub  →  Block Bulk Meter  →  DB check meters  →  apartment meters
    Row conventions in the sheet:
      - Parent DB in ('MS1','MS2')           → that row IS the block bulk meter
      - Parent DB == 'Block X Bulk Meter'    → that row is a DB check meter (stand = DB name)
      - Parent DB == 'DB-...'                → apartment meter fed from that DB
    Returns {block: {minisub, bulk, checks, dbs}}.
    """
    xls = pd.ExcelFile(file_path)
    elec_name = find_sheet(xls, APRT_ELEC_CANDIDATES)
    if not elec_name:
        return {}

    edf = xls.parse(elec_name)
    edf.columns = [str(c).strip() for c in edf.columns]
    edf["serial_str"]   = edf["Elect meter serial"].astype(str).str.strip()
    edf["stand_str"]    = edf["Stand"].astype(str).str.strip()
    edf["block"]        = edf["Apartment Block"].astype(str).str.strip()
    edf["db"]           = edf["Parent DB"].astype(str).str.strip()
    edf["amr"]          = edf["AMR installed"].fillna(False).astype(bool)
    edf["parent_meter"] = edf["Parent Meter"].astype(str).str.strip()

    MS_NAMES = {"MS1", "MS2", "MS3"}
    hierarchy = {}
    for block, bgrp in edf.groupby("block"):
        entry = {"minisub": None, "bulk": None, "checks": {}, "dbs": {}}

        # Bulk meter row: Parent DB is a minisub
        bulk_rows = bgrp[bgrp["db"].isin(MS_NAMES)]
        if not bulk_rows.empty:
            br = bulk_rows.iloc[0]
            entry["bulk"] = {"stand": br["stand_str"], "serial": br["serial_str"],
                             "amr": bool(br["amr"])}
            entry["minisub"] = {"name": br["db"], "serial": br["parent_meter"]}

        bulk_group_name = entry["bulk"]["stand"] if entry["bulk"] else None

        # DB check meters: rows whose Parent DB is the bulk meter stand name
        if bulk_group_name:
            for _, cr in bgrp[bgrp["db"] == bulk_group_name].iterrows():
                entry["checks"][cr["stand_str"]] = {
                    "serial": cr["serial_str"], "amr": bool(cr["amr"]),
                }

        # Apartment DB groups: everything else
        skip = MS_NAMES | ({bulk_group_name} if bulk_group_name else set())
        for db, dgrp in bgrp[~bgrp["db"].isin(skip)].groupby("db"):
            entry["dbs"][db] = {
                "parent_meter": dgrp["parent_meter"].iloc[0],
                "check_serial": entry["checks"].get(db, {}).get("serial", dgrp["parent_meter"].iloc[0]),
                "check_amr":    entry["checks"].get(db, {}).get("amr", False),
                "meters":       dgrp["stand_str"].tolist(),
                "serials":      dict(zip(dgrp["stand_str"], dgrp["serial_str"])),
                "amr_meters":   dgrp[dgrp["amr"]]["stand_str"].tolist(),
                "total":        len(dgrp),
                "amr_count":    int(dgrp["amr"].sum()),
            }
        hierarchy[block] = entry
    return hierarchy


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
    elec_name = find_sheet(xls, ELEC_SHEET_CANDIDATES)
    if not elec_name:
        return {}
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


def build_amr_map(polygons, kiosks, df, amr_readings, faulty_df, center, kiosk_meters=None):
    """
    Folium map coloured by AMR import status.
    amr_readings: dict of serial → {reading_date, reading_value, ...}
    faulty_df:    DataFrame of faulty-but-not-replaced meters (stand column)
    kiosk_meters: optional dict kiosk_name → list of {stand, serial, eb_port, amr}
                  — when given, kiosk markers get a clickable popup listing every
                  connected meter with its AMR state.
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
        "meter-off":  "#2F4B7C",   # AMR wired (EB port), meter switched off
        "no-amr":     "#3A4454",   # no AMR device fitted
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
        any_amr_live     = False
        any_eb_port      = False

        for _, r in rows.iterrows():
            serial = r.get("serial", "")
            if not serial:
                continue
            if bool(r.get("amr", False)):
                any_amr_live = True
            if str(r.get("eb_port", "") or "").strip() not in ("", "nan", "None"):
                any_eb_port = True
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

        # If no meter on this stand is AMR-live and nothing has ever imported,
        # distinguish "wired but off" (EB port assigned) from "AMR not fitted".
        meter_state_note = ""
        if best_badge == "amr-never" and not any_amr_live:
            if any_eb_port:
                best_badge = "meter-off"
                meter_state_note = ("<div style='margin:4px 0;padding:3px 8px;background:#2F4B7C22;"
                                    "color:#7FA3D8;border:1px solid #2F4B7C88;border-radius:5px;"
                                    "font-size:11px;font-weight:700'>🔌 Meter off — AMR wired, power not on</div>")
            else:
                best_badge = "no-amr"
                meter_state_note = ("<div style='margin:4px 0;padding:3px 8px;background:#3A445422;"
                                    "color:#8A97A8;border:1px solid #3A445488;border-radius:5px;"
                                    "font-size:11px;font-weight:700'>🔧 AMR not fitted yet</div>")

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
            f"{meter_state_note}"
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
        _meters = (kiosk_meters or {}).get(kname, [])

        if _meters:
            n_live = n_off = n_none = 0
            rows_html = ""
            for mt in _meters:
                _amr  = bool(mt.get("amr"))
                _ebp  = str(mt.get("eb_port", "") or "").strip()
                _ser  = mt.get("serial", "")
                _rd   = amr_readings.get(_ser)
                if _amr:
                    n_live += 1
                    _lbl, _col, _ = amr_status_info(_rd.get("reading_date") if _rd else None)
                    _state = _lbl
                elif _ebp and _ebp not in ("nan", "None"):
                    n_off += 1
                    _col, _state = "#2F4B7C", "Meter off — AMR wired"
                else:
                    n_none += 1
                    _col, _state = "#3A4454", "AMR not fitted"
                _port_s = f" · EB {_ebp}" if _ebp and _ebp not in ("nan", "None") else ""
                rows_html += (
                    f"<tr>"
                    f"<td style='padding:2px 4px'><span style='display:inline-block;width:9px;height:9px;"
                    f"border-radius:50%;background:{_col}'></span></td>"
                    f"<td style='padding:2px 4px;font-weight:600'>{mt.get('stand','')}</td>"
                    f"<td style='padding:2px 4px;font-family:monospace;font-size:10px'>{_ser}</td>"
                    f"<td style='padding:2px 4px;color:#555;font-size:10px'>{_state}{_port_s}</td>"
                    f"</tr>")
            popup_html = (
                f"<div style='font-family:sans-serif;min-width:250px'>"
                f"<div style='font-size:13px;font-weight:700;color:#152B45'>⚡ Kiosk {kname}</div>"
                f"<div style='font-size:10px;color:#666;margin:3px 0 6px'>"
                f"{len(_meters)} meters · 🟢 {n_live} live · 🔌 {n_off} off · 🔧 {n_none} no AMR</div>"
                f"<div style='max-height:230px;overflow-y:auto'>"
                f"<table style='font-size:11px;border-collapse:collapse;width:100%'>{rows_html}</table>"
                f"</div></div>")
            popup = folium.Popup(popup_html, max_width=330)
        else:
            popup = None

        folium.CircleMarker(
            location=[k["lat"], k["lon"]], radius=8,
            color="#B96E1E", fill=True, fill_color="#E69138", fill_opacity=0.95, weight=2,
            tooltip=folium.Tooltip(f"<b>\u26a1 {kname}</b>" +
                                   (f"<br><span style='font-size:10px'>{len(_meters)} meters — click for AMR status</span>" if _meters else ""),
                                   sticky=True),
            popup=popup,
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
    st.dataframe(show, width='stretch', hide_index=True)
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


# ---------- Site selector ----------
def _find_logo(*keywords):
    """Find a logo image by keyword, tolerant of renamed files.
    Searches assets/ and the repo root for png/jpg containing any keyword."""
    base = os.path.dirname(__file__)
    for folder in (os.path.join(base, "assets"), base):
        if not os.path.isdir(folder):
            continue
        for f in sorted(os.listdir(folder)):
            if not f.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
                continue
            name = f.lower()
            if any(k in name for k in keywords):
                return os.path.join(folder, f)
    return None

_EVG_LOGO = _find_logo("evergreen", "sitari", "estate")
_VOL_LOGO = _find_logo("voltano")

# Voltano logo pinned in the Streamlit header (top-left)
if _VOL_LOGO:
    try:
        st.logo(_VOL_LOGO, size="large")
    except Exception:
        pass

if IS_MOBILE:
    # Compact stacked header for phones
    if _EVG_LOGO:
        st.image(_EVG_LOGO, width=210)
    st.markdown("### 🔧 Sitari Evergreen — Meter Commissioning")
    st.caption("Erf 1186 Sitari · Lifestyle Retirement Village · managed by **Voltano Metering**")
else:
    hc1, hc2, hc3 = st.columns([1.3, 3, 1.3])
    with hc1:
        if _EVG_LOGO:
            st.image(_EVG_LOGO, width='stretch')
    with hc2:
        st.title("🔧 Sitari Evergreen — Meter Commissioning")
        st.caption("Erf 1186 Sitari · Lifestyle Retirement Village · managed by **Voltano Metering**")
    with hc3:
        if _VOL_LOGO:
            st.image(_VOL_LOGO, width='stretch')

# Handle stand-click links from the apartment floor plan (query params).
# Must run BEFORE the sidebar widgets so we can steer them.
_qp_stand = st.query_params.get("sel_stand")
if _qp_stand:
    st.session_state["fp_selected_stand"] = _qp_stand
    if st.query_params.get("view") == "apartments":
        st.session_state["site_type"] = "🏢 Apartments"
        st.session_state["workspace"] = "📊 Meter Management"
        st.session_state["page_met"]  = "AMR Live"
    st.query_params.clear()

# ---------- Sidebar navigation ----------
WORKSPACES = {
    "🏘️ Estate Management": ["Outstanding", "Upcoming", "Overdue", "Installed",
                              "Calendar", "Sections", "Reticulation",
                              "Faulty Meters", "Stock"],
    "📊 Meter Management":  ["Estate Map", "AMR Live", "Balancing"],
}
_PAGE_ICONS = {
    "Outstanding": "🟦", "Upcoming": "🟧", "Overdue": "🟥", "Installed": "🟩",
    "Calendar": "📅", "Sections": "📊", "Reticulation": "⚡",
    "Faulty Meters": "⚠️", "Stock": "📦",
    "Estate Map": "🗺️", "AMR Live": "📡", "Balancing": "⚖️",
}

with st.sidebar:
    st.markdown("## 🧭 Navigation")
    workspace = st.radio("Workspace", list(WORKSPACES.keys()), key="workspace")

    _page_key = "page_est" if "Estate" in workspace else "page_met"
    _pages = WORKSPACES[workspace]
    _PAGE = st.radio("Page", _pages, key=_page_key,
                     format_func=lambda p: f"{_PAGE_ICONS.get(p,'')} {p}")

    st.divider()
    site_type = st.radio("Scope", ["🏠 Freestanding", "🏢 Apartments", "🌍 All"],
                         key="site_type")
    st.toggle("📱 Compact mobile layout", key="is_mobile",
              help="Auto-detected from your device — flip manually if the layout looks off.")
    st.caption("Voltano Metering · Evergreen Sitari")

is_apartments = "Apartments" in site_type
scope_all     = "All" in site_type
site_key      = "apartments" if is_apartments else "freestanding"

# ---------- Load data ----------
data_path = find_data_file()
if not data_path:
    st.error(f"No file matching `{FILE_PATTERN}` found in this app's folder. Push the latest spreadsheet to the repo to continue.")
    st.stop()

mtime = os.path.getmtime(data_path)
if scope_all:
    _df_fs  = load_data(data_path, mtime, "freestanding")
    _df_apt = load_data(data_path, mtime, "apartments")
    df = pd.concat([_df_fs, _df_apt], ignore_index=True) if not _df_apt.empty else _df_fs
else:
    df = load_data(data_path, mtime, site_key)

if df.empty:
    st.error("Couldn't find the required sheets in this file. Check the sheet names.")
    st.stop()

_scope_label = "🌍 Whole site" if scope_all else ("🏢 Apartments" if is_apartments else "🏠 Freestanding")
st.caption(
    f"📂 Using **{os.path.basename(data_path)}** · last updated {datetime.fromtimestamp(mtime).strftime('%d %b %Y, %H:%M')}"
    f" · {_scope_label}"
)

# ---------- KPI strip (always visible) ----------
total = len(df)
installed_n   = int(df["installed"].sum())
overdue_n     = int((df["status"] == "Overdue").sum())
due_soon_n    = int((df["status"] == "Due soon").sum())
outstanding_n = total - installed_n

if is_apartments:
    metric_row([
        ("Total meter points", total),
        ("Installed (serial present)", installed_n, f"{round(installed_n/total*100) if total else 0}% complete"),
        ("AMR commissioned", int(df["amr"].sum()) if "amr" in df.columns else 0),
        ("Meter types", "Elec · Cold · Hot"),
    ])
else:
    faulty_n         = int(df["faulty"].sum()) if "faulty" in df.columns else 0
    faulty_pending_n = int(df[df["faulty"] & ~df["faulty_replaced"]]["stand"].count()) if "faulty" in df.columns else 0
    metric_row([
        ("Total meter points", total),
        ("Installed", installed_n, f"{round(installed_n/total*100) if total else 0}% complete"),
        ("Outstanding", outstanding_n),
        ("Due within 14 days", due_soon_n),
        ("Overdue", overdue_n, None, "inverse"),
        ("Faulty meters", faulty_n, f"{faulty_pending_n} awaiting replacement", "inverse"),
    ])

st.divider()

# ---------- Page banner (navigation lives in the left sidebar) ----------
_FS_ONLY_PAGES = {"Outstanding", "Upcoming", "Overdue", "Calendar", "Sections",
                  "Faulty Meters", "Stock", "Estate Map"}
st.markdown(f"### {_PAGE_ICONS.get(_PAGE, '')} {_PAGE}")
if is_apartments and _PAGE in _FS_ONLY_PAGES:
    st.info(f"**{_PAGE}** applies to the freestanding project (apartments have no "
            f"{'stock or ' if _PAGE == 'Stock' else ''}snag-date tracking here). "
            "Switch the scope to 🏠 Freestanding or 🌍 All in the sidebar to view it.")

COLS   = ["stand", "meter_type", "unit_type", "wbho_section", "deadline", "status"]
RENAME = {"stand":"Stand","meter_type":"Type","unit_type":"Unit type",
          "wbho_section":"Section","deadline":"Deadline (Snag 4)","status":"Status"}

if _PAGE == "Outstanding" and not is_apartments:
    st.subheader("All outstanding meters")
    outstanding_full = df[~df["installed"]]
    filtered = status_filters(outstanding_full, "outstanding")
    summary_counters(filtered, outstanding_full, "outstanding")
    show_table(filtered, COLS, RENAME, sort_col="deadline")

if _PAGE == "Upcoming" and not is_apartments:
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

if _PAGE == "Overdue" and not is_apartments:
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

if _PAGE == "Installed" and not is_apartments:
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

if _PAGE == "Calendar" and not is_apartments:
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

if _PAGE == "Sections" and not is_apartments:
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
        width='stretch',
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
        width='stretch', hide_index=True,
        column_config={"% complete": st.column_config.ProgressColumn("% complete", min_value=0, max_value=100, format="%d%%")},
    )

st.caption("Reads the spreadsheet pushed into this repo folder. Push an updated copy whenever meters are installed — the app picks up the most recently modified matching file automatically.")

# =====================================================================
# RETICULATION TAB — single-line diagram: Minisubs → Kiosks → Meters
# =====================================================================
if _PAGE == "Reticulation" and not is_apartments:
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

    # Apartment block BULK meters hang off the minisubs (Block A ← MS2, Block B ← MS1)
    aprt_hier_fs = load_aprt_reticulation(data_path, mtime)
    BLOCK_LABEL_FS = {"Block A": "Helderberg Suites", "Block B": "Tafelberg Suites"}
    ms_aprt_blocks = {}
    for _bname, _entry in (aprt_hier_fs or {}).items():
        if not _entry.get("bulk") or not _entry.get("minisub"):
            continue
        try:
            _ms_id = int(str(_entry["minisub"]["name"]).upper().replace("MS", ""))
        except ValueError:
            continue
        ms_aprt_blocks[_ms_id] = {
            "block": _bname.strip(),
            "label": BLOCK_LABEL_FS.get(_bname.strip(), _bname.strip()),
            "bulk_serial": _entry["bulk"]["serial"],
            "bulk_amr": bool(_entry["bulk"]["amr"]),
            "total": sum(d["total"] for d in _entry["dbs"].values()),
            "amr_count": sum(d["amr_count"] for d in _entry["dbs"].values()),
        }

    for ms_id in sorted(hierarchy.keys()):
        ms = hierarchy[ms_id]
        ms_installed = sum(k["installed"] for k in ms["kiosks"])
        ms_planned = sum(k["planned"] for k in ms["kiosks"])
        ms_amr = sum(k["amr_count"] for k in ms["kiosks"])
        try:
            _ms_int = int(float(ms_id))
        except (ValueError, TypeError):
            _ms_int = None
        diagram_data.append({
            "ms_id": ms_id,
            "serial": ms["serial"],
            "ms_installed": ms_installed,
            "ms_planned": ms_planned,
            "ms_amr": ms_amr,
            "kiosks": ms["kiosks"],
            "aprt_block": ms_aprt_blocks.get(_ms_int),
        })

    diagram_json = json.dumps(diagram_data)
    highlight_json = json.dumps(highlight_stands)

    # Build faulty stand lookup for the reticulation diagram.
    # This is the ELECTRICAL reticulation — only electrical meter faults belong
    # here. Water faults (tracked in FS Water) must not mark these stand chips.
    _elec_df = df[df["meter_type"] == "Electrical"] if "meter_type" in df.columns else df
    faulty_stands_set   = set(_elec_df[_elec_df["faulty"]]["stand"].tolist()) if "faulty" in _elec_df.columns else set()
    faulty_replaced_set = set(_elec_df[_elec_df["faulty"] & _elec_df["faulty_replaced"]]["stand"].tolist()) if "faulty" in _elec_df.columns else set()
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

    // ── Apartment block bulk meter — rendered FIRST, at the top ──
    if (ms.aprt_block) {{
      const ab = ms.aprt_block;
      const abEntry = document.createElement('div');
      abEntry.className = 'kiosk-entry';
      const abDrop = document.createElement('div');
      abDrop.className = 'kiosk-drop-line';
      abDrop.style.background = '#8B5CF6';
      abEntry.appendChild(abDrop);

      const abNode = document.createElement('div');
      abNode.className = 'kiosk-node';
      abNode.style.borderColor = '#8B5CF6';
      abNode.style.background = '#181330';
      abNode.style.cursor = 'default';
      abNode.innerHTML = `
        <div class="kiosk-header">
          <span class="kiosk-id" style="color:#A78BFA">🏢 ${{ab.block}} Bulk</span>
          <span class="kiosk-counts" style="font-family:monospace">${{ab.bulk_serial}}</span>
        </div>
        <div class="kiosk-amr-line">
          <span style="color:#A78BFA">${{ab.label}} · ${{ab.total}} apt meters · ${{ab.amr_count}} AMR</span>
          ${{!ab.bulk_amr ? '<span class="amr-miss-count">⚠ bulk AMR pending</span>' : ''}}
        </div>`;
      abEntry.appendChild(abNode);
      grid.appendChild(abEntry);
    }}

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

    components.html(html, height=650 if IS_MOBILE else 900, scrolling=True)

# =====================================================================
# ESTATE MAP TAB
# =====================================================================
if _PAGE == "Estate Map" and not is_apartments:
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
                _elec_sheet_name = find_sheet(_xls, ELEC_SHEET_CANDIDATES)
                _edf = _xls.parse(_elec_sheet_name)
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
                render_cached_map(map_html, height=430 if IS_MOBILE else 660)

            st.caption(
                "Satellite imagery: Esri World Imagery. "
                "Click any house polygon to see meter and AMR status. "
                "Commit an updated KML to the repo to refresh the map."
            )

# =====================================================================
# FAULTY METERS TAB
# =====================================================================
if _PAGE == "Faulty Meters" and not is_apartments:
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

            st.dataframe(display, width='stretch', hide_index=True,
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
if _PAGE == "AMR Live" and not is_apartments:
    st.subheader("📡 AMR Live — Meter Reading Status")
    st.caption(
        "Live meter readings imported hourly from the automated collection system. "
        "Matched against the commissioning spreadsheet per meter serial."
    )

    # ── SFTP settings — read silently from secrets (never shown on screen) ──
    try:
        _sftp_cfg = st.secrets.get("sftp", {})
    except Exception:
        _sftp_cfg = {}
    sftp_host = _sftp_cfg.get("host", "")
    sftp_user = _sftp_cfg.get("username", "")
    sftp_pass = _sftp_cfg.get("password", "")
    sftp_dir  = _sftp_cfg.get("directory", "")
    sftp_port = int(_sftp_cfg.get("port", 22))
    sftp_ready = all([sftp_host, sftp_user, sftp_pass, sftp_dir])

    # ── Load readings: SFTP or uploaded CSV ──────────────────────────
    amr_readings = {}
    amr_source   = None
    amr_file_ts  = None
    amr_error    = None

    # Manual CSV upload — admin fallback, tucked away
    with st.expander("🛠️ Admin: manual CSV import", expanded=False):
        uploaded_amr = st.file_uploader(
            "Upload an AMR CSV file directly (fallback if the automated feed is down)",
            type=["csv"], key="amr_csv_upload"
        )

    # ── Auto-sync on startup ──────────────────────────────────────────
    # Runs once per browser session. Checks what the most recent file in
    # the DB is and fetches any SFTP files newer than that — so the DB
    # always catches up automatically without manual intervention.
    if sftp_ready and "amr_auto_synced" not in st.session_state:
        latest_ts = get_latest_file_ts()
        now = datetime.now()

        if latest_ts is None:
            # DB is empty (fresh deploy) — seed with last 24 hours
            sync_since = now - timedelta(hours=24)
            sync_label = "last 24h (no data in DB)"
        elif (now - latest_ts).total_seconds() > 3600:
            # More than one hour since last fetched file — catch up
            sync_since = latest_ts
            sync_label = f"since {latest_ts.strftime('%d %b %Y %H:%M')}"
        else:
            sync_since = None
            sync_label = None

        if sync_since:
            with st.spinner(f"🔄 Auto-syncing AMR readings ({sync_label})…"):
                _af, _ar, _rdgs, _err = fetch_amr_bulk_history(
                    sftp_host, int(sftp_port), sftp_user, sftp_pass, sftp_dir,
                    since_dt=sync_since
                )
            if _err:
                st.warning(f"Auto-sync warning: {_err}")
            elif _ar > 0:
                st.toast(f"✅ Auto-synced {_af} files · {_ar} new readings added to history DB")
                # Use the latest fetched readings as the current snapshot
                if _rdgs:
                    save_amr_cache({"readings": _rdgs, "source": "Auto-sync",
                                    "file_ts": None})
            else:
                st.toast("✅ AMR history is up to date")

        st.session_state["amr_auto_synced"] = True  # don't repeat this session


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

        if bulk_err and files_done == 0:
            # Complete failure — no files downloaded at all
            st.error(f"SFTP/fetch error: {bulk_err}")
        else:
            # Always save latest readings to cache regardless of DB success
            if latest_rdgs:
                amr_readings = latest_rdgs
                amr_source   = f"SFTP bulk ({bulk_hours}h)"
                amr_file_ts  = None
                save_amr_cache({"readings": amr_readings, "source": amr_source, "file_ts": None})

            total_rows, distinct_serials, mn, mx = db_stats()
            st.success(
                f"✅ Fetched **{files_done}** files · "
                f"**{new_rows}** new readings added to DB · "
                f"DB now holds **{total_rows:,}** readings across **{distinct_serials}** serials "
                f"({mn[:10] if mn else '—'} → {mx[:10] if mx else '—'})"
            )
            # If DB had connection errors, surface them so they're diagnosable
            if bulk_err and files_done > 0:
                st.warning(f"⚠️ Files were fetched but DB writes had errors — readings are in the live display only. DB error: {bulk_err}")
            if files_done == 0 and not latest_rdgs:
                st.info(f"No CSV files found in the selected window ({bulk_hours}h). The most recent file on SFTP may be older — try increasing the hours window.")

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
        # Load from local JSON cache (fast — written after each fetch)
        cached = load_amr_cache()
        if cached and cached.get("readings"):
            amr_readings = cached["readings"]
            amr_source   = cached.get("source", "Cached data")
            raw_ts        = cached.get("file_ts")
            amr_file_ts  = datetime.fromisoformat(raw_ts) if raw_ts else None
        else:
            # Cache is empty (fresh redeploy) — restore snapshot from Supabase DB
            with st.spinner("Restoring latest readings from database…"):
                db_restored = load_latest_from_db()
            if db_restored:
                amr_readings = db_restored
                amr_source   = "Restored from Supabase DB"
                amr_file_ts  = None
                save_amr_cache({"readings": amr_readings, "source": amr_source, "file_ts": None})
                st.toast(f"✅ Restored {len(db_restored)} meter readings from database")

    # Auto-fetch on page load if SFTP configured and no data yet.
    # The auto-sync above will have already populated the DB and cache if it ran,
    # so try loading from cache first before making another SFTP round-trip.
    if not amr_readings and sftp_ready and not fetch_clicked and not bulk_clicked:
        # Re-check cache (auto-sync may have just written to it)
        cached = load_amr_cache()
        if cached and cached.get("readings"):
            amr_readings = cached["readings"]
            amr_source   = cached.get("source", "Auto-sync")
            raw_ts       = cached.get("file_ts")
            amr_file_ts  = datetime.fromisoformat(raw_ts) if raw_ts else None
        else:
            import time
            cache_bust = int(time.time() // 3600)
            with st.spinner("Fetching latest readings from SFTP…"):
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

        # ── Split: AMR installed vs wired-but-off vs not fitted ────────
        # amr=True                      → live, import expected
        # amr=False + EB port assigned  → AMR wired, meter still OFF (no import expected)
        # amr=False + no EB port        → AMR device not fitted yet
        all_installed = df[df["installed"] & df["serial"].str.len().gt(0)].copy()
        all_installed = all_installed.drop_duplicates(subset=["stand", "meter_type"])
        if "eb_port" not in all_installed.columns:
            all_installed["eb_port"] = ""
        all_installed["eb_port"] = all_installed["eb_port"].fillna("").astype(str)
        amr_installed_df = all_installed[all_installed["amr"] == True]
        _not_live        = all_installed[all_installed["amr"] == False]
        meter_off_df     = _not_live[_not_live["eb_port"].str.len() > 0]
        amr_pending_df   = _not_live[_not_live["eb_port"].str.len() == 0]

        rows = []
        for _, r in amr_installed_df.iterrows():
            serial  = r["serial"]
            reading = amr_readings.get(serial)
            rd_iso  = reading["reading_date"]  if reading else None
            rd_val  = reading["reading_value"] if reading else None
            low_bat = int(reading["low_battery"]) if reading else 0
            label, color, badge = amr_status_info(rd_iso)
            rows.append({"stand":r["stand"],"meter_type":r["meter_type"],"unit_type":r["unit_type"],
                         "wbho_section":r["wbho_section"],"serial":serial,"last_reading":rd_iso,
                         "reading_value":rd_val,"low_battery":low_bat,
                         "status_label":label,"status_color":color,"status_badge":badge})

        amr_table = pd.DataFrame(rows) if rows else pd.DataFrame(
            columns=["stand","meter_type","unit_type","wbho_section","serial",
                     "last_reading","reading_value","low_battery","status_label","status_color","status_badge"])

        green  = amr_table[amr_table["status_badge"]=="amr-green"]
        yellow = amr_table[amr_table["status_badge"]=="amr-yellow"]
        orange = amr_table[amr_table["status_badge"]=="amr-orange"]
        red    = amr_table[amr_table["status_badge"]=="amr-red"]
        never  = amr_table[amr_table["status_badge"]=="amr-never"]
        total_amr_inst = len(amr_table)
        importing_pct  = round(len(green)/total_amr_inst*100) if total_amr_inst else 0

        faulty_pending_df    = df[df["faulty"] & ~df["faulty_replaced"]] if "faulty" in df.columns else pd.DataFrame()
        faulty_pending_stands = set(faulty_pending_df["stand"].tolist())

        # ── KPI row ───────────────────────────────────────────────────
        st.markdown("#### AMR commissioned meters — import status")
        metric_row([
            ("AMR commissioned", total_amr_inst),
            ("🟢 <24h",  len(green),  f"{importing_pct}%"),
            ("🟡 1–3d",  len(yellow), f"{round(len(yellow)/total_amr_inst*100) if total_amr_inst else 0}%"),
            ("🟠 4–7d",  len(orange), f"{round(len(orange)/total_amr_inst*100) if total_amr_inst else 0}%"),
            ("🔴 7d+",   len(red),    f"{round(len(red)/total_amr_inst*100) if total_amr_inst else 0}%"),
            ("⚫ Never",  len(never),  f"{round(len(never)/total_amr_inst*100) if total_amr_inst else 0}%"),
            ("🔌 Meter off", len(meter_off_df), "AMR wired, power not on", "off"),
            ("🔧 AMR not fitted", len(amr_pending_df), "no AMR device yet"),
        ])

        # ── Per-type summary rows (Electricity / Water / Hot Water separately) ──
        _type_icons = {"Electrical": "⚡", "Water": "💧", "Water (Cold)": "💧", "Water (Hot)": "♨️"}
        for _tname in sorted(amr_table["meter_type"].unique().tolist()):
            _tsub  = amr_table[amr_table["meter_type"] == _tname]
            _tpend = amr_pending_df[amr_pending_df["meter_type"] == _tname]
            _toff  = meter_off_df[meter_off_df["meter_type"] == _tname]
            _tcnt  = _tsub["status_badge"].value_counts()
            _ttot  = len(_tsub)
            _tg    = int(_tcnt.get("amr-green", 0))
            metric_row([
                (f"{_type_icons.get(_tname,'')} {_tname}", _ttot),
                ("🟢 <24h",  _tg, f"{round(_tg/_ttot*100) if _ttot else 0}%"),
                ("🟡 1–3d",  int(_tcnt.get("amr-yellow", 0))),
                ("🟠 4–7d",  int(_tcnt.get("amr-orange", 0))),
                ("🔴 7d+",   int(_tcnt.get("amr-red", 0))),
                ("⚫ Never", int(_tcnt.get("amr-never", 0))),
                ("🔌 Off", len(_toff)),
                ("🔧 Not fitted", len(_tpend)),
            ])

        low_bat_meters   = amr_table[amr_table["low_battery"]==1]
        stale_and_faulty = amr_table[amr_table["status_badge"].isin(["amr-orange","amr-red","amr-never"]) &
                                     amr_table["stand"].isin(faulty_pending_stands)]
        ac1, ac2 = st.columns(2)
        with ac1:
            if not low_bat_meters.empty:
                st.warning(f"🔋 **{len(low_bat_meters)} low battery** — " +
                           ", ".join(low_bat_meters["stand"].head(6).tolist()) +
                           (" …" if len(low_bat_meters)>6 else ""))
        with ac2:
            if not stale_and_faulty.empty:
                st.warning(f"⚠️ **{len(stale_and_faulty)} stale + faulty** — " +
                           ", ".join(f"Stand {r['stand']}" for _,r in stale_and_faulty.head(5).iterrows()) +
                           (" …" if len(stale_and_faulty)>5 else ""))

        st.divider()

        # ── Two-column layout: filters+table | map ────────────────────
        left_col, right_col = st.columns([1, 1.6])

        with left_col:
            st.markdown("##### Filters & detail")
            af1, af2 = st.columns(2)
            with af1:
                _types_present = sorted(amr_table["meter_type"].unique().tolist()) if not amr_table.empty else sorted(df["meter_type"].unique().tolist())
                a_type = st.multiselect("Type", _types_present,
                                        default=_types_present, key="amr_type")
            with af2:
                a_status = st.multiselect("Status",
                    ["🟢 Last 24h","🟡 1–3 days","🟠 4–7 days","🔴 7+ days","⚫ No reading"],
                    default=["🟢 Last 24h","🟡 1–3 days","🟠 4–7 days","🔴 7+ days","⚫ No reading"],
                    key="amr_status_filter")
            a_section = st.multiselect("Section",sorted(amr_table["wbho_section"].dropna().unique()),key="amr_section")
            a_serial_search = st.text_input("Search serial",key="amr_serial_search")

            status_badge_map = {"🟢 Last 24h":"amr-green","🟡 1–3 days":"amr-yellow",
                                "🟠 4–7 days":"amr-orange","🔴 7+ days":"amr-red","⚫ No reading":"amr-never"}
            selected_badges = [status_badge_map[s] for s in a_status]

            view = amr_table[amr_table["meter_type"].isin(a_type) &
                             amr_table["status_badge"].isin(selected_badges)]
            if a_section:
                view = view[view["wbho_section"].isin(a_section)]
            if a_serial_search:
                view = view[view["serial"].str.contains(a_serial_search.strip(),case=False,na=False)]

            def fmt_reading(v, mtype):
                if v is None: return "—"
                return f"{v:,.1f} L" if mtype=="Water" else f"{v:,.3f} kWh"

            display_rows = []
            for _, r in view.sort_values(["status_badge","stand"]).iterrows():
                display_rows.append({
                    "Stand":r["stand"],"Type":r["meter_type"],"Section":r["wbho_section"],
                    "Serial":r["serial"],
                    "Last read":str(r["last_reading"])[:16].replace("T"," ") if r["last_reading"] and str(r["last_reading"]) not in ("nan","None","") else "—",
                    "Reading":fmt_reading(r["reading_value"],r["meter_type"]),
                    "Status":r["status_label"],
                    "Fault":"\u26a0\ufe0f" if r["stand"] in faulty_pending_stands else "",
                    "Bat":"\U0001f50b" if r["low_battery"] else "",
                })
            disp_df = pd.DataFrame(display_rows)
            st.caption(f"{len(view)} meters matching filters")
            st.dataframe(disp_df,width='stretch',hide_index=True,height=300)

            if not meter_off_df.empty:
                with st.expander(f"🔌 Meter off — AMR wired, power not on ({len(meter_off_df)})"):
                    st.caption("These meters have an EB port allocated on the kiosk AMR device, "
                               "but the meter is still switched off (building site). No import is "
                               "expected until the power goes on — tick **AMR Installed** in the "
                               "sheet once it's live.")
                    st.dataframe(meter_off_df[["stand","meter_type","unit_type","wbho_section","serial","eb_port"]].rename(
                        columns={"stand":"Stand","meter_type":"Type","unit_type":"Unit","wbho_section":"Section","serial":"Serial","eb_port":"EB Port"}),
                        width='stretch',hide_index=True)

            if not amr_pending_df.empty:
                with st.expander(f"🔧 AMR not yet fitted ({len(amr_pending_df)})"):
                    st.dataframe(amr_pending_df[["stand","meter_type","unit_type","wbho_section","serial"]].rename(
                        columns={"stand":"Stand","meter_type":"Type","unit_type":"Unit","wbho_section":"Section","serial":"Serial"}),
                        width='stretch',hide_index=True)

            import hashlib
            csv_out = disp_df.to_csv(index=False).encode("utf-8")
            st.download_button("\u2b07\ufe0f Download table",csv_out,file_name="amr_status.csv",
                               mime="text/csv",key=f"dl_amr_{hashlib.md5(csv_out).hexdigest()[:8]}")

            st.markdown("**Import rate by type**")
            for mtype in ["Water","Electrical"]:
                sub = amr_table[amr_table["meter_type"]==mtype]
                if sub.empty: continue
                cnt = sub["status_badge"].value_counts()
                tot = len(sub)
                pct = round(cnt.get("amr-green",0)/tot*100) if tot else 0
                low = len(sub[sub["low_battery"]==1])
                parts = [f"{e} {cnt.get(b,0)}" for b,e in [("amr-green","\U0001f7e2"),("amr-yellow","\U0001f7e1"),
                         ("amr-orange","\U0001f7e0"),("amr-red","\U0001f534"),("amr-never","\u26ab")] if cnt.get(b,0)]
                st.caption(f"**{mtype}** {tot} · {pct}% importing · " + " ".join(parts) +
                           (f" · \U0001f50b {low}" if low else ""))

        with right_col:
            st.markdown("##### AMR site map")
            _kml_path = find_kml_file()
            _amr_polygons,_amr_kiosks = [],[]
            if _kml_path:
                try:
                    with open(_kml_path,"rb") as _f: _kml_data = _f.read()
                    _amr_polygons,_amr_kiosks,_ = parse_kml(_kml_data)
                except Exception: pass

            # Shared: legend + history section (used by both map paths)
            _amr_legend_html = "<div style='display:flex;flex-wrap:wrap;gap:10px;font-size:11px;margin-top:4px'>"
            for _lc, _ll in [("#2E7D52","🟢 <24h"),("#D4AC0D","🟡 1–3d"),("#E67E22","🟠 4–7d"),
                             ("#BD4B2C","🔴 7d+"),("#607080","⚫ AMR live, never seen"),
                             ("#2F4B7C","🔌 Meter off (AMR wired)"),("#3A4454","🔧 AMR not fitted"),
                             ("#EF4444","⚠️ Faulty")]:
                _amr_legend_html += (f"<span><span style='display:inline-block;width:11px;height:11px;border-radius:2px;"
                                     f"background:{_lc};vertical-align:middle;margin-right:3px'></span>{_ll}</span>")
            _amr_legend_html += "</div>"

            def _render_amr_history():
                st.markdown("##### Stand reading history — water & electricity")
                _hist_pool = df[df["installed"] & df["serial"].str.len().gt(0)]
                stand_options = sorted(_hist_pool["stand"].unique().tolist())
                # A widget key set from a map click must exist in options, else drop it
                if st.session_state.get("amr_stand_select") not in stand_options:
                    st.session_state.pop("amr_stand_select", None)
                default_stand = st.session_state.get("amr_selected_stand",
                                                     stand_options[0] if stand_options else None)
                if default_stand not in stand_options and stand_options:
                    default_stand = stand_options[0]

                sel_col, info_col = st.columns([1, 2.4])
                with sel_col:
                    selected_stand = st.selectbox(
                        "Select stand (auto-set when you click the map)",
                        options=stand_options,
                        index=stand_options.index(default_stand) if default_stand in stand_options else 0,
                        key="amr_stand_select")
                    if selected_stand:
                        st.session_state["amr_selected_stand"] = selected_stand
                with info_col:
                    total_rows_db, _, mn, mx = db_stats()
                    st.caption(f"History DB: **{total_rows_db:,}** readings · "
                               f"{mn[:10] if mn else '—'} → {mx[:10] if mx else '—'}"
                               + (" · ⚠️ Faulty meter on this stand" if selected_stand in faulty_pending_stands else ""))

                if selected_stand:
                    render_stand_history(selected_stand, df, key_prefix="amr")
                st.caption("Auto-refreshes hourly. History stored permanently in Supabase.")

            if not _amr_polygons:
                st.info("No KML found — push `EVG_Sitari.kml` to see the map.")
                st.divider()
                _render_amr_history()
            else:
                _amr_map_df = df.copy() if len(a_type) == len(_types_present) else df[df["meter_type"].isin(a_type)]
                _center = kml_center(_amr_polygons,_amr_kiosks,[])

                if not is_apartments:
                    # Clickable map: st_folium returns the clicked popup so a stand
                    # click auto-opens its history below. Map object cached by resource.
                    # Kiosk → connected elec meters (stand, serial, EB port, AMR state)
                    try:
                        _kx = pd.ExcelFile(data_path)
                        _ken = find_sheet(_kx, ELEC_SHEET_CANDIDATES)
                        _ke = _kx.parse(_ken)
                        _ke.columns = [str(c).strip() for c in _ke.columns]
                        _ke = _ke[_ke["Kiosk Number"].notna()].copy()
                        _ke["kiosk"] = _ke["Kiosk Number"].astype(str).str.strip()
                        _ke["stand"] = _ke["Stand Number"].astype(str).str.strip()
                        _ke["ser"]   = _ke["Meter Serial"].apply(
                            lambda v: str(int(float(v))) if pd.notna(v) and str(v) not in ("nan","") else "")
                        _ke["amrf"]  = _ke["AMR Installed"].fillna(False).astype(bool)
                        _keb = next((c for c in _ke.columns
                                     if "eb" in c.lower() and "port" in c.lower()), None)
                        if _keb:
                            _ke["ebp"] = _ke[_keb].apply(
                                lambda v: "" if pd.isna(v) or str(v).strip() in ("","nan","None")
                                else (str(int(float(v))) if str(v).replace(".","",1).replace("-","",1).isdigit() else str(v).strip()))
                        else:
                            _ke["ebp"] = ""
                        kiosk_meters_map = {
                            kn: [{"stand": r["stand"], "serial": r["ser"],
                                  "eb_port": r["ebp"], "amr": bool(r["amrf"])}
                                 for _, r in grp.iterrows() if r["ser"]]
                            for kn, grp in _ke.groupby("kiosk")
                        }
                    except Exception:
                        kiosk_meters_map = {}

                    @st.cache_resource(show_spinner=False)
                    def _cached_amr_map_obj(geom_h, df_h, amr_h, f_h, km_h):
                        return build_amr_map(_amr_polygons, _amr_kiosks, _amr_map_df,
                                             amr_readings, faulty_pending_df, _center,
                                             kiosk_meters=kiosk_meters_map)
                    _gh = _poly_hash(_amr_polygons,_amr_kiosks)
                    _dh = _df_hash(_amr_map_df)
                    _ah = _amr_hash(amr_readings)
                    _fph = str(sorted(faulty_pending_df["stand"].tolist())) if not faulty_pending_df.empty else ""
                    _kmh = str(sorted((k, len(v), sum(x["amr"] for x in v)) for k, v in kiosk_meters_map.items()))
                    with st.spinner("Building map…"):
                        _mobj = _cached_amr_map_obj(_gh, _dh, _ah, _fph, _kmh)

                    # Fragment: clicks rerun ONLY this section — the rest of the
                    # page stays put, and we feed the last zoom/centre back into
                    # the map so it never jumps back to the initial view.
                    @st.fragment
                    def _amr_map_and_history():
                        _view = st.session_state.get("amr_map_view", {})
                        map_ret = st_folium(
                            _mobj, use_container_width=True,
                            height=380 if IS_MOBILE else 500,
                            center=_view.get("center"), zoom=_view.get("zoom"),
                            returned_objects=["last_object_clicked_popup", "zoom", "center"],
                            key="amr_click_map")
                        if map_ret:
                            _z = map_ret.get("zoom")
                            _c = map_ret.get("center")
                            if isinstance(_c, dict) and "lat" in _c:
                                _c = (_c["lat"], _c.get("lng", _c.get("lon")))
                            if _z or _c:
                                st.session_state["amr_map_view"] = {"zoom": _z, "center": _c}
                        _clicked = (map_ret or {}).get("last_object_clicked_popup") or ""
                        if _clicked:
                            import re as _re_click
                            _mm = _re_click.search(r"Stand ID:\s*(\w+)", str(_clicked))
                            if _mm:
                                st.session_state["amr_selected_stand"] = _mm.group(1)
                                st.session_state["amr_stand_select"]   = _mm.group(1)
                        st.caption("💡 Click a stand for its graphs, or a kiosk for its meter list — "
                                   "the map keeps its position and only this section refreshes.")
                        st.markdown(_amr_legend_html, unsafe_allow_html=True)
                        st.divider()
                        _render_amr_history()

                    _amr_map_and_history()
                else:
                    _gh = _poly_hash(_amr_polygons,_amr_kiosks)
                    _dh = _df_hash(_amr_map_df)
                    _ah = _amr_hash(amr_readings)
                    _fph = str(sorted(faulty_pending_df["stand"].tolist())) if not faulty_pending_df.empty else ""
                    with st.spinner("Building map…"):
                        amr_map_html = cached_amr_map_html(_gh,_dh,_ah,_fph,_center,
                            _amr_polygons,_amr_kiosks,_amr_map_df,amr_readings,faulty_pending_df)
                        render_cached_map(amr_map_html, height=380 if IS_MOBILE else 500)
                    st.markdown(_amr_legend_html, unsafe_allow_html=True)
                    st.divider()
                    _render_amr_history()

# =====================================================================
# APARTMENT RETICULATION TAB — 3 levels: Minisub → Bulk → DBs → meters
# =====================================================================
if _PAGE == "Reticulation" and is_apartments:
    st.subheader("🏢 Apartment Reticulation — Minisub → Bulk Meter → Distribution Boards")
    st.caption("Block A (Helderberg) is fed from Minisub 2; Block B (Tafelberg) from Minisub 1. "
               "Each block has a bulk check meter, feeding DB check meters, feeding apartment meters. "
               "Click a DB to expand its meters.")

    aprt_hier = load_aprt_reticulation(data_path, mtime)

    if not aprt_hier:
        st.warning("No apartment electrical data found. Check the 'Aprt Elec' sheet.")
    else:
        total_aprt_meters = sum(sum(d["total"] for d in e["dbs"].values()) for e in aprt_hier.values())
        total_aprt_amr    = sum(sum(d["amr_count"] for d in e["dbs"].values()) for e in aprt_hier.values())
        ak1, ak2, ak3 = st.columns(3)
        ak1.metric("Total apartment meters", total_aprt_meters)
        ak2.metric("AMR commissioned",       total_aprt_amr)
        ak3.metric("Blocks",                 len(aprt_hier))
        st.divider()

        import json as _json_ar
        BLOCK_LABEL_AR = {"Block A": "Helderberg Suites", "Block B": "Tafelberg Suites"}
        aprt_diagram = []
        for block_name in sorted(aprt_hier.keys()):
            e = aprt_hier[block_name]
            # Every check meter under the bulk becomes a node: DBs (with children),
            # plus Lifts / UPS / Plant (usually no child meters — shown standalone).
            db_list = []
            for chk_name in sorted(e["checks"].keys()):
                c = e["checks"][chk_name]
                d = e["dbs"].get(chk_name)
                db_list.append({
                    "db": chk_name,
                    "check_serial": c["serial"],
                    "check_amr": c["amr"],
                    "total": d["total"] if d else 0,
                    "amr_count": d["amr_count"] if d else 0,
                    "meters": d["meters"] if d else [],
                    "serials": d["serials"] if d else {},
                    "amr_meters": d["amr_meters"] if d else [],
                })
            # Any DB group without a matching check meter row (data gap) still shown
            for db_name in sorted(e["dbs"].keys()):
                if db_name in e["checks"]:
                    continue
                d = e["dbs"][db_name]
                db_list.append({
                    "db": db_name, "check_serial": d["check_serial"],
                    "check_amr": d["check_amr"], "total": d["total"],
                    "amr_count": d["amr_count"], "meters": d["meters"],
                    "serials": d["serials"], "amr_meters": d["amr_meters"],
                })
            aprt_diagram.append({
                "block": block_name.strip(),
                "label": BLOCK_LABEL_AR.get(block_name.strip(), ""),
                "minisub": e["minisub"] or {"name": "?", "serial": ""},
                "bulk": e["bulk"] or {"stand": "Bulk", "serial": "", "amr": False},
                "dbs": db_list,
            })
        diagram_json = _json_ar.dumps(aprt_diagram)

        html = f"""<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'IBM Plex Mono','Courier New',monospace;background:#0e1117;color:#e0e0e0;padding:12px;}}
.blocks-row{{display:flex;gap:28px;flex-wrap:wrap;align-items:flex-start;justify-content:center;}}
.block-col{{display:flex;flex-direction:column;align-items:center;min-width:230px;max-width:300px;flex:1;}}
.ms-box{{background:#0f2440;border:2px solid #2E5F8A;border-radius:10px;padding:9px 16px;
        text-align:center;width:100%;}}
.ms-title{{font-size:13px;font-weight:700;color:#7FB3E0;}}
.ms-sub{{font-size:9px;color:#5B86B3;font-family:monospace;margin-top:2px;}}
.bulk-box{{background:#1a1530;border:2px solid #8B5CF6;border-radius:10px;padding:9px 16px;
          text-align:center;width:100%;}}
.bulk-title{{font-size:13px;font-weight:700;color:#A78BFA;}}
.bulk-sub{{font-size:9px;color:#8B5CF6;font-family:monospace;margin-top:2px;}}
.vline{{width:3px;height:18px;background:#5B86B3;}}
.vline.purple{{background:#8B5CF6;}}
.db-list{{width:100%;display:flex;flex-direction:column;}}
.db-drop{{width:3px;height:12px;background:#8B5CF6;margin:0 auto;}}
.db-node{{width:100%;background:#131c2b;border:1.5px solid #334d6e;border-radius:7px;
          cursor:pointer;padding:7px 10px;transition:border-color .15s;}}
.db-node:hover{{border-color:#8B5CF6;}}
.db-node.expanded{{border-color:#E69138;}}
.db-header{{display:flex;align-items:center;gap:6px;}}
.db-name{{font-size:11px;font-weight:700;color:#c8d8eb;min-width:70px;}}
.db-serial{{font-size:8px;color:#5B86B3;font-family:monospace;}}
.mini-bar{{flex:1;height:4px;border-radius:2px;background:#2a3f55;overflow:hidden;}}
.mini-fill{{height:100%;border-radius:2px;background:#8B5CF6;}}
.db-counts{{font-size:9px;color:#7A96B2;white-space:nowrap;}}
.db-detail{{display:none;background:#0d1520;border:1px solid #1e3050;border-top:none;
            border-radius:0 0 7px 7px;padding:7px 9px;}}
.db-detail.open{{display:block;}}
.meter-grid{{display:flex;flex-wrap:wrap;gap:3px;margin-top:4px;}}
.meter-chip{{padding:1px 5px;border-radius:4px;font-size:9px;font-weight:600;
             display:inline-flex;align-items:center;gap:2px;}}
.chip-amr{{background:#3F7D5C22;color:#6eb88a;border:1px solid #3F7D5C55;}}
.chip-no{{background:#1F3F6622;color:#7a9ec4;border:1px solid #334d6e;}}
.dot{{display:inline-block;width:5px;height:5px;border-radius:50%;}}
.extra-note{{font-size:8px;color:#5B86B3;margin-top:6px;}}
</style></head><body>
<div class="blocks-row" id="root"></div>
<script>
const data = {diagram_json};
function pct(a,t){{return t>0?Math.round(a/t*100):0;}}

data.forEach(b => {{
  const col = document.createElement('div'); col.className='block-col';

  // Level 1 — Minisub
  const ms = document.createElement('div'); ms.className='ms-box';
  ms.innerHTML = `<div class="ms-title">🔌 ${{b.minisub.name}}</div>
    <div class="ms-sub">Serial ${{b.minisub.serial}}</div>`;
  col.appendChild(ms);
  const v1 = document.createElement('div'); v1.className='vline'; col.appendChild(v1);

  // Level 2 — Block bulk meter
  const bulk = document.createElement('div'); bulk.className='bulk-box';
  bulk.innerHTML = `<div class="bulk-title">🏢 ${{b.block}} Bulk Meter</div>
    <div class="bulk-sub">${{b.label}} · Serial ${{b.bulk.serial}}${{b.bulk.amr?'':' · ⚠ AMR pending'}}</div>`;
  col.appendChild(bulk);
  const v2 = document.createElement('div'); v2.className='vline purple'; col.appendChild(v2);

  // Level 3 — DB check meters with apartment meters below
  const dbList = document.createElement('div'); dbList.className='db-list';
  b.dbs.forEach(d => {{
    const drop=document.createElement('div'); drop.className='db-drop'; dbList.appendChild(drop);
    const node=document.createElement('div'); node.className='db-node';
    const dp = pct(d.amr_count, d.total);
    const dbId = (b.block + '-' + d.db).replace(/[^a-zA-Z0-9]/g,'');
    const hasKids = d.meters && d.meters.length > 0;
    node.innerHTML = `<div class="db-header">
      <span class="db-name">${{d.db}}</span>
      <span class="db-serial">${{d.check_serial}}</span>
      <div class="mini-bar"><div class="mini-fill" style="width:${{dp}}%"></div></div>
      <span class="db-counts">${{hasKids ? d.amr_count + '/' + d.total : (d.check_amr ? 'AMR ✓' : 'no AMR')}}</span>
      ${{hasKids ? `<span style="font-size:9px;color:#5B86B3" id="chev-${{dbId}}">▾</span>` : ''}}
    </div>`;

    const det=document.createElement('div'); det.className='db-detail'; det.id='det-'+dbId;
    if (hasKids) {{
      const amrSet=new Set(d.amr_meters);
      const chips=d.meters.map(m => {{
        const ok=amrSet.has(m); const serial=(d.serials&&d.serials[m])||'';
        return `<span class="meter-chip ${{ok?'chip-amr':'chip-no'}}"
          title="${{m}} · Serial ${{serial||'—'}} · AMR ${{ok?'✓':'pending'}}">
          <span class="dot" style="background:${{ok?'#3F7D5C':'#334d6e'}}"></span>${{m.split('/').pop()}}</span>`;
      }}).join('');
      det.innerHTML = `<div style="font-size:8px;color:#8B5CF6;margin-bottom:3px">
        CHECK METER ${{d.check_serial}} → ${{d.total}} apartment meters (${{d.amr_count}} AMR)</div>
        <div class="meter-grid">${{chips}}</div>`;
      node.addEventListener('click', () => {{
        const open = det.classList.toggle('open');
        node.classList.toggle('expanded', open);
      }});
    }} else {{
      node.style.cursor = 'default';
      det.innerHTML = `<div style="font-size:8px;color:#5B86B3">
        No child meters allocated to this check meter yet.</div>`;
    }}
    dbList.appendChild(node); dbList.appendChild(det);
  }});
  col.appendChild(dbList);


  document.getElementById('root').appendChild(col);
}});
</script></body></html>"""
        components.html(html, height=650 if IS_MOBILE else 900, scrolling=True)

# =====================================================================
# APARTMENT INSTALLED TAB (when in apartment mode)
# =====================================================================
if _PAGE == "Installed" and is_apartments:
    st.subheader("🟩 Apartment meters — installed list")
    for mtype in sorted(df["meter_type"].unique()):
        sub = df[df["meter_type"]==mtype]
        inst = sub[sub["installed"]]
        with st.expander(f"**{mtype}** — {len(inst)}/{len(sub)} installed"):
            disp = inst[["stand","unit_type","wbho_section","serial","amr"]].rename(
                columns={"stand":"Stand","unit_type":"Block","wbho_section":"DB/Block","serial":"Serial","amr":"AMR"})
            st.dataframe(disp, width='stretch', hide_index=True)


# =====================================================================
# BALANCING TAB — parent vs children consumption with losses
# =====================================================================
if _PAGE == "Balancing":
    st.subheader("⚖️ Consumption Balancing — Electrical (kWh)")
    st.caption(
        "Energy-flow view of the whole site: each parent check meter compared against the sum of "
        "its children over the selected window, with losses and estimates made visible."
    )

    # ── Time selection: slider or date range, locked to available data ──
    _, _, _db_min, _db_max = db_stats()
    _now = datetime.now()
    try:
        _earliest = datetime.fromisoformat(_db_min) if _db_min else (_now - timedelta(days=30))
    except Exception:
        _earliest = _now - timedelta(days=30)
    _max_days = max(1, min(365, int((_now - _earliest).total_seconds() // 86400) + 1))

    selm, selt = st.columns([2.4, 1.4])
    with selm:
        bal_mode = st.radio("Time selection", ["🎚️ Quick slider", "📅 Date range"],
                            horizontal=True, key="bal_mode")
    with selt:
        bal_tol = st.radio("Anchor tolerance", ["±1h", "±3h", "±6h"], index=1,
                           horizontal=True, key="bal_tol",
                           help="Readings must fall within this window of both anchor times to count as measured.")
    _tol_h = {"±1h": 1, "±3h": 3, "±6h": 6}[bal_tol]

    if "slider" in bal_mode.lower():
        _days_back = st.slider(
            "Window — days back from now", 1, _max_days,
            value=min(7, _max_days), key="bal_days",
            help=f"Data available from {_earliest.strftime('%d %b %Y %H:%M')} — "
                 f"the slider is capped there so the window never starts before the first reading.")
        _start = max(_now - timedelta(days=_days_back), _earliest)
        _end   = _now
    else:
        from datetime import time as _dtime
        _dr = st.date_input(
            "Period (start is locked to the first available reading)",
            value=(max(_now - timedelta(days=7), _earliest).date(), _now.date()),
            min_value=_earliest.date(), max_value=_now.date(),
            key="bal_dates")
        if isinstance(_dr, (list, tuple)) and len(_dr) == 2:
            _d0, _d1 = _dr
        elif isinstance(_dr, (list, tuple)) and len(_dr) == 1:
            _d0 = _d1 = _dr[0]
        else:
            _d0 = _d1 = _dr
        _start = max(datetime.combine(_d0, _dtime.min), _earliest)
        _end   = min(datetime.combine(_d1, _dtime.max), _now)

    cons = db_get_consumption(_start.isoformat(), _end.isoformat(), _tol_h)

    if not cons:
        st.info("No consumption data in the selected window yet — the history DB fills up as hourly readings import.")
    else:
        spans = [c["span_hours"] for c in cons.values() if c.get("span_hours")]
        avg_span = sum(spans) / len(spans) if spans else 0
        st.caption(f"Window: **{_start.strftime('%d %b %H:%M')} → {_end.strftime('%d %b %H:%M')}** · "
                   f"**{len(cons)}** meters aligned to anchors (avg span {avg_span:,.1f}h, tolerance {bal_tol})")

        def _delta(serial):
            c = cons.get(str(serial).strip())
            return c["delta"] if c else None

        # ── Load hierarchies ────────────────────────────────────────────
        fs_hier   = load_kiosk_data(data_path, mtime)
        aprt_hier = load_aprt_reticulation(data_path, mtime)

        try:
            _xk = pd.ExcelFile(data_path)
            _en = find_sheet(_xk, ELEC_SHEET_CANDIDATES)
            _ek = _xk.parse(_en)
            _ek.columns = [str(c).strip() for c in _ek.columns]
            _ek["kiosk"] = _ek["Kiosk Number"].astype(str).str.strip()
            _ek["ser"]   = _ek["Meter Serial"].apply(
                lambda v: str(int(float(v))) if pd.notna(v) and str(v) not in ("nan", "") else "")
            kiosk_serials    = _ek[_ek["ser"] != ""].groupby("kiosk")["ser"].apply(list).to_dict()
            kiosk_all_counts = _ek.groupby("kiosk")["Stand Number"].count().to_dict()
            _ek["stand_s"]   = _ek["Stand Number"].astype(str).str.strip()
            kiosk_stand_serials = (
                _ek[_ek["ser"] != ""]
                .groupby("kiosk")[["stand_s", "ser"]]
                .apply(lambda g: list(zip(g["stand_s"], g["ser"])))
                .to_dict()
            )
        except Exception:
            kiosk_serials, kiosk_all_counts, kiosk_stand_serials = {}, {}, {}

        MS_TO_BLOCK = {}
        for _bn, _be in (aprt_hier or {}).items():
            if _be.get("minisub"):
                MS_TO_BLOCK[str(_be["minisub"]["name"]).upper()] = _bn

        # ── Pre-compute per-minisub structures (shared by KPIs + visuals) ──
        ms_view = []
        for ms_id in sorted(fs_hier.keys()):
            ms = fs_hier[ms_id]
            ms_serial = str(ms["serial"]).strip()
            ms_delta  = _delta(ms_serial)

            kiosk_rows = []
            kiosk_sum, kiosk_missing = 0.0, 0
            for k in ms["kiosks"]:
                serials = kiosk_serials.get(k["kiosk"], [])
                deltas  = [_delta(s) for s in serials]
                have    = [d for d in deltas if d is not None]
                ksum    = sum(have)
                kmiss   = max((kiosk_all_counts.get(k["kiosk"], len(serials))) - len(have), 0)
                kiosk_sum += ksum
                kiosk_missing += kmiss
                kiosk_rows.append({"label": f"⚡ {k['kiosk']}", "value": ksum, "missing": kmiss})

            blk = MS_TO_BLOCK.get(f"MS{int(float(ms_id))}" if str(ms_id).replace('.', '').isdigit()
                                  else str(ms_id).upper())
            bulk_delta, bulk_virtual, bulk_v_missing = None, False, 0
            if blk and aprt_hier[blk].get("bulk"):
                bs = aprt_hier[blk]["bulk"]["serial"]
                bulk_delta = _delta(bs)
                if bulk_delta is None:
                    _e = aprt_hier[blk]
                    _vchecks = dict(_e["checks"])
                    for _dbn, _d in _e["dbs"].items():
                        _vchecks.setdefault(_dbn, {"serial": _d["check_serial"], "amr": _d["check_amr"]})
                    _vd = [_delta(c["serial"]) for c in _vchecks.values()]
                    _vh = [x for x in _vd if x is not None]
                    if _vh:
                        bulk_delta   = sum(_vh)
                        bulk_virtual = True
                        bulk_v_missing = len(_vd) - len(_vh)

            children_total = kiosk_sum + (bulk_delta or 0)
            miss_total = kiosk_missing + (bulk_v_missing if bulk_virtual
                                          else (1 if (blk and bulk_delta is None) else 0))
            est = max(ms_delta - children_total, 0.0) if (ms_delta is not None and miss_total > 0) else 0.0
            loss = (ms_delta - children_total - est) if ms_delta is not None else None

            ms_view.append({
                "ms_id": ms_id, "serial": ms_serial, "delta": ms_delta,
                "kiosks": kiosk_rows, "kiosk_missing": kiosk_missing,
                "blk": blk, "bulk_delta": bulk_delta, "bulk_virtual": bulk_virtual,
                "children_total": children_total, "miss_total": miss_total,
                "est": est, "loss": loss,
            })

        # ── Site KPIs ──────────────────────────────────────────────────
        total_parent   = sum(m["delta"] for m in ms_view if m["delta"] is not None)
        total_children = sum(m["children_total"] for m in ms_view)
        total_est      = sum(m["est"] for m in ms_view)
        total_loss     = sum(m["loss"] for m in ms_view if m["loss"] is not None)
        loss_pct       = (total_loss / total_parent * 100) if total_parent else 0
        ms_measured_n  = sum(1 for m in ms_view if m["delta"] is not None)

        metric_row([
            ("🔌 Minisub intake", f"{total_parent:,.0f} kWh", f"{ms_measured_n}/{len(ms_view)} minisubs measured"),
            ("✅ Metered downstream", f"{total_children:,.0f} kWh"),
            ("🟣 Estimated (unmetered)", f"{total_est:,.0f} kWh",
             f"{sum(m['miss_total'] for m in ms_view)} meters", "off"),
            ("📉 Unexplained loss", f"{total_loss:,.0f} kWh",
             f"{loss_pct:+.1f}%", "inverse" if loss_pct > 10 else "normal"),
        ], desktop_cols=4)

        st.divider()

        # ── 📄 Consumption balance report generator ─────────────────────
        st.markdown("#### 📄 Consumption balance report")
        st.caption("Pick a Minisub or an apartment block and generate a full downstream "
                   "breakdown for the selected window — every level's difference plus each "
                   "individual meter, downloadable as CSV.")

        _rep_opts = [f"Minisub {m['ms_id']}" for m in ms_view] + sorted((aprt_hier or {}).keys())
        rc1, rc2 = st.columns([2.4, 1])
        with rc1:
            rep_scope = st.selectbox("Report scope", _rep_opts, key="bal_report_scope")
        with rc2:
            st.markdown("<div style='height:1.75em'></div>", unsafe_allow_html=True)
            if st.button("🧾 Generate report", key="bal_report_btn"):
                st.session_state["bal_report_go"] = rep_scope

        def _block_report_rows(blk):
            """Diff rows + meter rows for one apartment block."""
            e = aprt_hier[blk]
            bulk_serial = e["bulk"]["serial"]
            bulk_d = _delta(bulk_serial)
            all_checks = dict(e["checks"])
            for dbn, d in e["dbs"].items():
                all_checks.setdefault(dbn, {"serial": d["check_serial"], "amr": d["check_amr"]})

            diff_rows, meter_rows = [], []
            checks_sum, checks_missing = 0.0, 0
            for cname in sorted(all_checks.keys()):
                cd = _delta(all_checks[cname]["serial"])
                meter_rows.append({"Section": f"{blk} · check meters", "Item": cname,
                                   "Serial": all_checks[cname]["serial"],
                                   "Consumption (kWh)": round(cd, 2) if cd is not None else None,
                                   "Note": "" if cd is not None else "no aligned reading"})
                if cd is None:
                    checks_missing += 1
                else:
                    checks_sum += cd
            _bd = bulk_d if bulk_d is not None else checks_sum
            _best = max(bulk_d - checks_sum, 0.0) if (bulk_d is not None and checks_missing) else 0.0
            diff_rows.append({
                "Level": f"{blk} Bulk → check meters",
                "Parent (kWh)": round(_bd, 2), "Children (kWh)": round(checks_sum, 2),
                "Estimated (kWh)": round(_best, 2),
                "Difference (kWh)": round(_bd - checks_sum - _best, 2),
                "Diff %": round((_bd - checks_sum - _best) / _bd * 100, 2) if _bd else None,
                "Unmetered": checks_missing,
                "Note": "virtual bulk (Σ checks)" if bulk_d is None else "",
            })
            for cname in sorted(e["dbs"].keys()):
                d = e["dbs"][cname]
                p = _delta(d["check_serial"])
                m_sum, m_miss = 0.0, 0
                for stand, ser in sorted(d["serials"].items()):
                    md = _delta(ser)
                    meter_rows.append({"Section": f"{blk} · {cname}", "Item": stand, "Serial": ser,
                                       "Consumption (kWh)": round(md, 2) if md is not None else None,
                                       "Note": "" if md is not None else "no aligned reading"})
                    if md is None:
                        m_miss += 1
                    else:
                        m_sum += md
                _e = max(p - m_sum, 0.0) if (p is not None and m_miss) else 0.0
                diff_rows.append({
                    "Level": f"{cname} → apartment meters",
                    "Parent (kWh)": round(p, 2) if p is not None else None,
                    "Children (kWh)": round(m_sum, 2), "Estimated (kWh)": round(_e, 2),
                    "Difference (kWh)": round(p - m_sum - _e, 2) if p is not None else None,
                    "Diff %": round((p - m_sum - _e) / p * 100, 2) if p else None,
                    "Unmetered": m_miss,
                    "Note": "" if p is not None else "check meter not aligned",
                })
            return diff_rows, meter_rows

        if st.session_state.get("bal_report_go") == rep_scope:
            diff_rows, meter_rows = [], []
            if rep_scope.startswith("Minisub"):
                _mid = rep_scope.replace("Minisub", "").strip()
                m = next((x for x in ms_view if str(x["ms_id"]) == _mid), None)
                if m:
                    _p = m["delta"] if m["delta"] is not None else m["children_total"]
                    diff_rows.append({
                        "Level": f"MS{m['ms_id']} → kiosks + block bulk",
                        "Parent (kWh)": round(_p, 2),
                        "Children (kWh)": round(m["children_total"], 2),
                        "Estimated (kWh)": round(m["est"], 2),
                        "Difference (kWh)": round(m["loss"], 2) if m["loss"] is not None else None,
                        "Diff %": round(m["loss"] / _p * 100, 2) if (m["loss"] is not None and _p) else None,
                        "Unmetered": m["miss_total"],
                        "Note": "MS meter not aligned — virtual (Σ children)" if m["delta"] is None else "",
                    })
                    for k in m["kiosks"]:
                        kname = k["label"].replace("⚡ ", "")
                        for stand, ser in kiosk_stand_serials.get(kname, []):
                            md = _delta(ser)
                            meter_rows.append({"Section": f"MS{m['ms_id']} · {kname}",
                                               "Item": f"Stand {stand}", "Serial": ser,
                                               "Consumption (kWh)": round(md, 2) if md is not None else None,
                                               "Note": "" if md is not None else "no aligned reading"})
                    # Downstream block, if this minisub feeds one
                    if m["blk"]:
                        bd, bm = _block_report_rows(m["blk"])
                        diff_rows.extend(bd)
                        meter_rows.extend(bm)
            else:
                diff_rows, meter_rows = _block_report_rows(rep_scope)

            _diff_df  = pd.DataFrame(diff_rows)
            _meter_df = pd.DataFrame(meter_rows)
            _aligned_n = int(_meter_df["Consumption (kWh)"].notna().sum()) if not _meter_df.empty else 0

            st.success(f"**{rep_scope}** · {_start.strftime('%d %b %Y %H:%M')} → "
                       f"{_end.strftime('%d %b %Y %H:%M')} · tolerance {bal_tol} · "
                       f"{_aligned_n}/{len(_meter_df)} meters aligned")
            st.markdown("**Differences by level**")
            st.dataframe(_diff_df, width='stretch', hide_index=True)
            st.markdown("**Meter-level detail**")
            st.dataframe(_meter_df, width='stretch', hide_index=True, height=320)

            import hashlib as _hl_rep
            _hdr = (f"Consumption balance report,{rep_scope}\n"
                    f"Window,{_start.isoformat()} to {_end.isoformat()},Tolerance,{bal_tol}\n\n")
            _csv = (_hdr + "DIFFERENCES BY LEVEL\n" + _diff_df.to_csv(index=False)
                    + "\nMETER DETAIL\n" + _meter_df.to_csv(index=False)).encode("utf-8")
            st.download_button(
                "⬇️ Download report (CSV)", _csv,
                file_name=f"balance_report_{rep_scope.replace(' ', '_')}_{_start.strftime('%Y%m%d')}-{_end.strftime('%Y%m%d')}.csv",
                mime="text/csv",
                key=f"bal_rep_dl_{_hl_rep.md5(_csv).hexdigest()[:8]}")

        st.divider()

        # ── Level 1 — Minisub energy flow ───────────────────────────────
        st.markdown("#### 🔌 Level 1 — Minisub energy flow")
        for m in ms_view:
            _ms_id, _delta_v = m["ms_id"], m["delta"]
            if _delta_v is not None and _delta_v > 0:
                _lp = (m["loss"] / _delta_v * 100) if m["loss"] is not None else 0
                badge = "🟢" if _lp <= 3 else ("🟠" if _lp <= 10 else "🔴")
                head = f"{badge} Minisub {_ms_id} — {_delta_v:,.1f} kWh in · loss {_lp:.1f}%"
                if m["est"] > 0:
                    head += f" · est. {m['est']:,.1f} kWh"
            else:
                head = f"⚫ Minisub {_ms_id} — no reading yet (children measured: {m['children_total']:,.1f} kWh)"

            with st.expander(head, expanded=False):
                children = [(k["label"], k["value"], "child") for k in m["kiosks"] if k["value"] > 0]
                if m["blk"] and m["bulk_delta"] is not None:
                    kind = "est" if m["bulk_virtual"] else "child"
                    lbl = f"🏢 {m['blk']}" + (" (virtual Σ)" if m["bulk_virtual"] else "")
                    children.append((lbl, m["bulk_delta"], kind))
                fig = build_balance_sankey(
                    f"MS{_ms_id}", _delta_v, children, est_val=m["est"],
                    height=280 if IS_MOBILE else 340)
                st.plotly_chart(fig, width='stretch', key=f"bal_sankey_ms_{_ms_id}")

                notes = []
                if m["bulk_virtual"]:
                    notes.append(f"ℹ️ {m['blk']} bulk has no AMR yet — shown as the sum of its check meters (purple).")
                if m["est"] > 0:
                    notes.append(f"🟣 {m['est']:,.1f} kWh attributed to {m['miss_total']} unmetered meter(s) — "
                                 "replaced by real readings as they come online.")
                unmetered_kiosks = [k for k in m["kiosks"] if k["missing"]]
                if unmetered_kiosks:
                    notes.append("Unmetered per kiosk: " + " · ".join(
                        f"{k['label'].replace('⚡ ', '')} ×{k['missing']}" for k in unmetered_kiosks))
                if m["loss"] is not None and m["loss"] < -0.01:
                    notes.append(f"🟠 Children exceed the minisub by {abs(m['loss']):,.1f} kWh — "
                                 "check anchor alignment or CT ratios.")
                for n in notes:
                    st.caption(n)

        st.divider()

        # ── Levels 2–3 — Apartment blocks ───────────────────────────────
        st.markdown("#### 🏢 Levels 2–3 — Apartment blocks")
        for blk in sorted((aprt_hier or {}).keys()):
            e = aprt_hier[blk]
            if not e.get("bulk"):
                continue
            bulk_serial = e["bulk"]["serial"]
            bulk_delta  = _delta(bulk_serial)

            all_checks = dict(e["checks"])
            for dbn, d in e["dbs"].items():
                all_checks.setdefault(dbn, {"serial": d["check_serial"], "amr": d["check_amr"]})

            check_vals, check_missing = [], 0
            for cname in sorted(all_checks.keys()):
                cd = _delta(all_checks[cname]["serial"])
                if cd is None:
                    check_missing += 1
                else:
                    check_vals.append((cname, cd))
            checks_sum = sum(v for _, v in check_vals)
            b_est  = max(bulk_delta - checks_sum, 0.0) if (bulk_delta is not None and check_missing > 0) else 0.0
            b_loss = (bulk_delta - checks_sum - b_est) if bulk_delta is not None else None

            if bulk_delta is not None:
                _blp = (b_loss / bulk_delta * 100) if (b_loss is not None and bulk_delta) else 0
                badge = "🟢" if _blp <= 3 else ("🟠" if _blp <= 10 else "🔴")
                head = f"{badge} {blk} — bulk {bulk_delta:,.1f} kWh · loss {_blp:.1f}%"
            else:
                head = f"⚫ {blk} — bulk not on AMR yet (checks measured: {checks_sum:,.1f} kWh)"

            with st.expander(head, expanded=False):
                # Level 2 Sankey: bulk → every check meter
                fig2 = build_balance_sankey(
                    f"{blk} Bulk", bulk_delta,
                    [(c, v, "child") for c, v in check_vals],
                    est_val=b_est, height=300 if IS_MOBILE else 380)
                st.plotly_chart(fig2, width='stretch', key=f"bal_sankey_{blk}")
                if check_missing:
                    st.caption(f"{check_missing} check meter(s) without aligned readings in this window.")

                # Level 3 bullet chart: each DB check meter vs its children
                st.markdown("**Per-DB: check meter (◆) vs apartment meters**")
                rows = []
                for cname in sorted(all_checks.keys()):
                    c = all_checks[cname]
                    p = _delta(c["serial"])
                    d = e["dbs"].get(cname)
                    if d:
                        apt_deltas = [_delta(s) for s in d["serials"].values()]
                        apt_have   = [x for x in apt_deltas if x is not None]
                        measured   = sum(apt_have)
                        n_missing  = len(apt_deltas) - len(apt_have)
                        est = max(p - measured, 0.0) if (p is not None and n_missing > 0) else 0.0
                        rows.append({"label": cname, "parent": p, "measured": measured,
                                     "est": est, "missing": n_missing})
                    else:
                        rows.append({"label": f"{cname} (end-point)", "parent": p,
                                     "measured": p or 0.0, "est": 0.0, "missing": 0})
                if rows:
                    fig3 = build_db_bullet_chart(rows)
                    st.plotly_chart(fig3, width='stretch', key=f"bal_bullet_{blk}")
                    st.caption("Bars: 🟩 measured apartment usage + 🟪 estimate for unmetered units · "
                               "◆ diamond = the DB check meter's own measurement. "
                               "Bar short of the diamond = loss within that DB. "
                               "End-point loads (lifts, plant, UPS) show their own usage as the bar.")

        st.caption(
            "**Methodology** — Timestamp-aligned consumption: every meter's usage is measured between "
            "its readings closest to the same two anchor times (window start and end, within the "
            "selected tolerance), so parents and children are compared over the same physical period. "
            "Meters that can't be aligned count as unmetered. Where a parent has a reading but some "
            "children don't, the positive residual is shown in purple as an estimate attributed to the "
            "unmetered meters and is replaced by real data as those meters come online. A parent "
            "without AMR uses the sum of its children as a virtual parent until its readings arrive."
        )

# =====================================================================
# APARTMENT FLOOR PLAN TAB — clickable Plotly schematic + AMR history
# =====================================================================
if _PAGE == "AMR Live" and is_apartments:
    st.subheader("🏬 Block Floor Plans — AMR Import Status")
    st.caption(
        "Schematic per the architectural drawings: odd units north side, even south, three wings per "
        "floor. Same AMR colour scheme as AMR Live. **Click any unit to open its usage graphs below.**"
    )

    # Current AMR readings: cache → DB fallback
    _fp_readings = {}
    _cached = load_amr_cache()
    if _cached and _cached.get("readings"):
        _fp_readings = _cached["readings"]
    if not _fp_readings:
        _fp_readings = load_latest_from_db()

    # ── Overall AMR stats per meter type ───────────────────────────────
    st.markdown("#### 📡 AMR status — Electricity · Cold Water · Hot Water")
    _type_icons_fp = {"Electrical": "⚡", "Water (Cold)": "💧", "Water (Hot)": "♨️"}
    for _t in ["Electrical", "Water (Cold)", "Water (Hot)"]:
        _tsub = df[(df["meter_type"] == _t) & df["installed"] & df["serial"].str.len().gt(0)]
        if _tsub.empty:
            continue
        _amr_in  = _tsub[_tsub["amr"]]
        _pending = _tsub[~_tsub["amr"]]
        _badges  = {"amr-green": 0, "amr-yellow": 0, "amr-orange": 0, "amr-red": 0, "amr-never": 0}
        _lowbat  = 0
        for _, _r in _amr_in.iterrows():
            _rd = _fp_readings.get(_r["serial"])
            _, _, _b = amr_status_info(_rd.get("reading_date") if _rd else None)
            _badges[_b] = _badges.get(_b, 0) + 1
            if _rd and int(_rd.get("low_battery", 0)):
                _lowbat += 1
        _tt = len(_amr_in)
        metric_row([
            (f"{_type_icons_fp.get(_t,'')} {_t}", _tt, "AMR commissioned"),
            ("🟢 <24h",  _badges["amr-green"], f"{round(_badges['amr-green']/_tt*100) if _tt else 0}%"),
            ("🟡 1–3d",  _badges["amr-yellow"]),
            ("🟠 4–7d",  _badges["amr-orange"]),
            ("🔴 7d+",   _badges["amr-red"]),
            ("⚫ Never",  _badges["amr-never"]),
            ("🔧 Pending", len(_pending)),
            ("🔋 Low bat", _lowbat),
        ])

    st.divider()

    fp_type = st.radio("Colour by meter type", ["Electrical", "Water (Cold)", "Water (Hot)"],
                       horizontal=True, key="fp_meter_type")

    _fp_df = df[df["meter_type"] == fp_type]
    serial_by_stand   = dict(zip(_fp_df["stand"], _fp_df["serial"]))
    amr_flag_by_stand = dict(zip(_fp_df["stand"], _fp_df["amr"]))
    eb_port_by_stand  = dict(zip(_fp_df["stand"], _fp_df.get("eb_port", pd.Series([""] * len(_fp_df))).fillna("").astype(str)))

    import re as _re_fp
    import plotly.graph_objects as _go_fp

    BLOCK_NAMES = {"A": "Block A · Helderberg Suites", "B": "Block B · Tafelberg Suites"}
    FLOOR_LABEL = {"00": "Ground", "01": "First", "02": "Second"}

    blocks_fp = {}
    for stand in df["stand"].unique():
        m = _re_fp.match(r"Aprt BL ([AB])/(\d+)/(\d+)", str(stand))
        if not m:
            continue
        blk, floor, unit = m.group(1), m.group(2), int(m.group(3))
        serial = str(serial_by_stand.get(stand, "") or "")
        amr_fitted = bool(amr_flag_by_stand.get(stand, False))
        reading = _fp_readings.get(serial) if serial else None
        rd_iso  = reading.get("reading_date") if reading else None
        label, color, badge = amr_status_info(rd_iso)
        if not amr_fitted:
            _ebp = str(eb_port_by_stand.get(stand, "") or "").strip()
            if _ebp and _ebp not in ("nan", "None"):
                color, label = "#2F4B7C", f"Meter off — AMR wired (EB port {_ebp})"
            else:
                color, label = "#33415c", "AMR not fitted"
        blocks_fp.setdefault(blk, {}).setdefault(floor, []).append({
            "unit": unit, "unit_str": m.group(3), "stand": stand,
            "serial": serial, "color": color, "label": label,
            "value": reading.get("reading_value") if reading else None,
            "low_bat": int(reading.get("low_battery", 0)) if reading else 0,
        })

    def _fp_wing(n):
        n = n % 100
        return 0 if n <= 14 else 1 if n <= 28 else 2

    clicked_stand = None
    for blk in sorted(blocks_fp.keys()):
        st.markdown(f"**🏢 {BLOCK_NAMES.get(blk, 'Block ' + blk)}** — {fp_type}")

        xs, ys, colors, texts, customs, hovers, line_colors, line_widths = [], [], [], [], [], [], [], []
        floor_annotations = []
        floors_sorted = sorted(blocks_fp[blk].keys(), reverse=True)   # Second at top

        for fi, fl in enumerate(floors_sorted):
            y_base = (len(floors_sorted) - 1 - fi) * 3   # 6, 3, 0 bottom-up
            floor_annotations.append((y_base + 0.5, FLOOR_LABEL.get(fl, fl)))
            for u in blocks_fp[blk][fl]:
                n = u["unit"] % 100
                is_odd = n % 2 == 1
                col_idx = (n + 1) // 2 - 1 if is_odd else n // 2 - 1
                x = col_idx + _fp_wing(u["unit"]) * 1.6
                y = y_base + (1 if is_odd else 0)
                xs.append(x); ys.append(y)
                colors.append(u["color"])
                texts.append(u["unit_str"])
                customs.append(u["stand"])
                _uv = u["value"]; _ust = u["stand"]; _use = u["serial"] or "—"; _ul = u["label"]
                val_s = f"<br>Reading: {_uv}" if _uv is not None else ""
                bat_s = "<br>🔋 Low battery" if u["low_bat"] else ""
                hovers.append(f"<b>{_ust}</b><br>Serial: {_use}"
                              f"<br>{_ul}{val_s}{bat_s}<br><i>Click for graphs ↓</i>")
                line_colors.append("#FFF176" if u["low_bat"] else "#0e1117")
                line_widths.append(2.5 if u["low_bat"] else 1)

        fig = _go_fp.Figure(_go_fp.Scatter(
            x=xs, y=ys, mode="markers+text",
            marker=dict(symbol="square", size=26, color=colors,
                        line=dict(color=line_colors, width=line_widths)),
            text=texts, textfont=dict(size=8, color="white", family="monospace"),
            customdata=customs,
            hovertemplate="%{hovertext}<extra></extra>", hovertext=hovers,
        ))
        for ya, lbl in floor_annotations:
            fig.add_annotation(x=-1.8, y=ya, text=lbl, showarrow=False,
                               font=dict(size=10, color="#7A96B2"), textangle=-90)
        fig.update_layout(
            height=330, margin=dict(l=10, r=10, t=8, b=8),
            plot_bgcolor="#0E1117", paper_bgcolor="#0E1117",
            xaxis=dict(visible=False, range=[-2.5, 25]),
            yaxis=dict(visible=False, range=[-0.8, len(floors_sorted) * 3 - 0.2]),
            showlegend=False, dragmode=False,
        )
        ev = st.plotly_chart(fig, width='stretch',
                             on_select="rerun", selection_mode="points",
                             key=f"fp_plot_{blk}")
        try:
            pts = ev.selection.points if ev and ev.selection else []
        except Exception:
            pts = []
        if pts:
            clicked_stand = pts[0].get("customdata")

    # Legend
    st.markdown(
        "<div style='display:flex;flex-wrap:wrap;gap:14px;font-size:11px'>"
        "<span>🟩 &lt;24h</span><span>🟨 1–3d</span><span>🟧 4–7d</span><span>🟥 7d+</span>"
        "<span><span style='display:inline-block;width:11px;height:11px;background:#607080;border-radius:2px'></span> AMR fitted, never seen</span>"
        "<span><span style='display:inline-block;width:11px;height:11px;background:#2F4B7C;border-radius:2px'></span> Meter off (AMR wired)</span>"
        "<span><span style='display:inline-block;width:11px;height:11px;background:#33415c;border-radius:2px'></span> AMR not fitted</span>"
        "<span><span style='display:inline-block;width:11px;height:11px;outline:2px solid #FFF176;outline-offset:-2px;border-radius:2px'></span> Low battery</span>"
        "</div>", unsafe_allow_html=True)

    if clicked_stand:
        st.session_state["fp_stand_select"] = clicked_stand
        st.session_state["fp_selected_stand"] = clicked_stand

    st.divider()

    # ── History panel: clicked unit or manual search ───────────────────
    st.markdown("##### 📈 Unit history — electricity, cold & hot water")
    all_stands_fp = sorted(df["stand"].unique().tolist())
    _default_fp = st.session_state.get("fp_selected_stand")
    if _default_fp not in all_stands_fp:
        _default_fp = all_stands_fp[0] if all_stands_fp else None

    sel_fp = st.selectbox(
        "Unit (auto-selected when you click the floor plan)",
        options=all_stands_fp,
        index=all_stands_fp.index(_default_fp) if _default_fp in all_stands_fp else 0,
        key="fp_stand_select",
    )
    if sel_fp:
        st.session_state["fp_selected_stand"] = sel_fp
        render_stand_history(sel_fp, df, key_prefix="fp")

# =====================================================================
# STOCK TAB — inventory tracking with auto-deduction (freestanding)
# =====================================================================
if _PAGE == "Stock" and not is_apartments:
    st.subheader("📦 Stock on Hand — Freestanding Project")
    st.caption(
        "Stock counts auto-deduct from your commissioning spreadsheet: installations and faulty "
        "replacements **dated after the last stocktake** reduce the relevant item automatically. "
        "Use the forms below when you receive stock, program meters, or do a recount."
    )

    stock = db_stock_get()
    if not stock:
        st.warning("Stock tracking needs the Supabase database connection ([db] url in secrets). "
                   "Once connected, initial counts are seeded automatically.")
    else:
        # ── Auto-consumption from the spreadsheet since each baseline ──
        def _consumed_since(item_key, baseline_iso):
            """Units consumed after the baseline date, per item rules."""
            try:
                bdate = pd.Timestamp(datetime.fromisoformat(baseline_iso).date())
            except Exception:
                return 0, ""
            water = df[df["meter_type"] == "Water"]
            elec  = df[df["meter_type"] == "Electrical"]
            if item_key == "water_meter":
                inst = int((water["commission_date"] > bdate).sum())
                repl = int((water["faulty_replaced"] & (water["replacement_date"] > bdate)).sum())
                return inst + repl, f"{inst} installs + {repl} replacements"
            if item_key == "water_box":
                inst = int((water["commission_date"] > bdate).sum())
                return inst, f"{inst} new installs (replacements reuse box)"
            if item_key == "elec_programmed":
                inst = int((elec["commission_date"] > bdate).sum())
                repl = int((elec["faulty_replaced"] & (elec["replacement_date"] > bdate)).sum()) \
                       if "faulty_replaced" in elec.columns else 0
                return inst + repl, f"{inst} installs + {repl} replacements"
            return 0, "manual only (not auto-deducted)"

        # ── Current levels ──────────────────────────────────────────────
        ITEM_ICONS = {"water_box": "🧰", "water_meter": "💧",
                      "elec_programmed": "⚡", "elec_unprogrammed": "🔌"}
        levels = {}
        low_items = []
        for key, s in stock.items():
            used, used_note = _consumed_since(key, s["baseline_date"])
            current = s["baseline_qty"] + s["adjustments"] - used
            levels[key] = {"current": current, "used": used, "used_note": used_note, **s}
            if current <= s["reorder_level"]:
                low_items.append((key, current, s["reorder_level"]))

        if low_items:
            st.error("🚨 **Reorder needed:** " + " · ".join(
                f"{ITEM_ICONS.get(k,'')} {stock[k]['label']} — {c} left (reorder at {r})"
                for k, c, r in low_items))

        order = ["water_meter", "water_box", "elec_programmed", "elec_unprogrammed"]
        metric_row([
            (f"{ITEM_ICONS.get(k,'')} {levels[k]['label']}", levels[k]["current"],
             f"−{levels[k]['used']} used since count" if levels[k]["used"] else "no usage since count",
             "inverse" if levels[k]["current"] <= levels[k]["reorder_level"] else "off")
            for k in order if k in levels
        ], desktop_cols=4)

        with st.expander("ℹ️ How each count is calculated"):
            for k in order:
                if k not in levels:
                    continue
                l = levels[k]
                bdate_s = l["baseline_date"][:10]
                st.caption(
                    f"{ITEM_ICONS.get(k,'')} **{l['label']}** = {l['baseline_qty']} counted on {bdate_s} "
                    f"{'+' if l['adjustments'] >= 0 else '−'} {abs(l['adjustments'])} manual adjustments "
                    f"− {l['used']} auto-deducted ({l['used_note']}) = **{l['current']}**")

        st.divider()

        # ── Demand & allocation ─────────────────────────────────────────
        st.markdown("##### 📋 Demand vs stock — allocation guide")
        water = df[df["meter_type"] == "Water"]
        elec  = df[df["meter_type"] == "Electrical"]
        w_out  = int((~water["installed"]).sum())
        e_out  = int((~elec["installed"]).sum())
        w_flt  = int((water["faulty"] & ~water["faulty_replaced"]).sum())
        e_flt  = int((elec["faulty"] & ~elec["faulty_replaced"]).sum()) if "faulty" in elec.columns else 0

        def _alloc_line(icon, label, stock_now, n_faulty, n_installs, extra=""):
            to_faulty   = min(stock_now, n_faulty)
            to_installs = min(max(stock_now - to_faulty, 0), n_installs)
            shortfall   = (n_faulty + n_installs) - stock_now
            cols = st.columns([2.2, 1.2, 1.6, 1.6, 2])
            cols[0].markdown(f"{icon} **{label}**")
            cols[1].markdown(f"Stock: **{stock_now}**")
            cols[2].markdown(f"→ faulty replacements: **{to_faulty}** / {n_faulty}")
            cols[3].markdown(f"→ new installs: **{to_installs}** / {n_installs}")
            if shortfall > 0:
                cols[4].markdown(f"<span style='color:#BD4B2C;font-weight:700'>Order ≥ {shortfall}</span>"
                                 + (f" <span style='color:#7A96B2;font-size:12px'>{extra}</span>" if extra else ""),
                                 unsafe_allow_html=True)
            else:
                cols[4].markdown(f"<span style='color:#2E7D52;font-weight:700'>✓ covered</span>"
                                 + (f" <span style='color:#7A96B2;font-size:12px'>{extra}</span>" if extra else ""),
                                 unsafe_allow_html=True)

        _alloc_line("💧", "Water meters (Axioma 20mm)",
                    levels.get("water_meter", {}).get("current", 0), w_flt, w_out)
        _alloc_line("🧰", "Water meter boxes",
                    levels.get("water_box", {}).get("current", 0), 0, w_out,
                    extra="boxes needed for new installs only")
        _unprog = levels.get("elec_unprogrammed", {}).get("current", 0)
        _alloc_line("⚡", "Elec meters (programmed)",
                    levels.get("elec_programmed", {}).get("current", 0), e_flt, e_out,
                    extra=f"+{_unprog} unprogrammed available to program")
        st.caption("Allocation prioritises faulty replacements first, then new installations. "
                   "'Order' figures ignore unprogrammed units — program those first to extend cover.")

        st.divider()

        # ── Actions ────────────────────────────────────────────────────
        ac1, ac2, ac3 = st.columns(3)

        with ac1:
            st.markdown("**➕ Receive / adjust stock**")
            with st.form("stock_adjust_form", clear_on_submit=True):
                _keys = [k for k in order if k in stock]
                a_item = st.selectbox("Item", _keys,
                                      format_func=lambda k: f"{ITEM_ICONS.get(k,'')} {stock[k]['label']}")
                a_qty  = st.number_input("Quantity (+ received / − issued)", min_value=-500,
                                         max_value=500, value=0, step=1)
                a_note = st.text_input("Note (e.g. 'PO 1234 delivery')")
                if st.form_submit_button("Save movement") and a_qty != 0:
                    if db_stock_adjust(a_item, a_qty, a_note):
                        st.success(f"Recorded {a_qty:+d} × {stock[a_item]['label']}")
                        st.rerun()
                    else:
                        st.error("Could not save — check DB connection.")

        with ac2:
            st.markdown("**🔧 Program Hexing meters**")
            with st.form("stock_program_form", clear_on_submit=True):
                p_qty = st.number_input("Move unprogrammed → programmed", min_value=1,
                                        max_value=500, value=1, step=1)
                if st.form_submit_button("Program"):
                    ok1 = db_stock_adjust("elec_unprogrammed", -int(p_qty), "programmed batch")
                    ok2 = db_stock_adjust("elec_programmed",  int(p_qty), "programmed batch")
                    if ok1 and ok2:
                        st.success(f"Moved {int(p_qty)} meters to programmed stock")
                        st.rerun()
                    else:
                        st.error("Could not save — check DB connection.")

        with ac3:
            st.markdown("**🔄 Stocktake (reset count)**")
            with st.form("stock_baseline_form", clear_on_submit=True):
                _keys = [k for k in order if k in stock]
                b_item = st.selectbox("Item", _keys, key="bl_item",
                                      format_func=lambda k: f"{ITEM_ICONS.get(k,'')} {stock[k]['label']}")
                b_qty  = st.number_input("Counted quantity on hand", min_value=0,
                                         max_value=10000, value=0, step=1)
                b_reo  = st.number_input("Reorder alert level", min_value=0, max_value=1000,
                                         value=5, step=1)
                if st.form_submit_button("Set new count"):
                    if db_stock_set_baseline(b_item, b_qty, b_reo):
                        st.success(f"{stock[b_item]['label']} reset to {int(b_qty)} as of today")
                        st.rerun()
                    else:
                        st.error("Could not save — check DB connection.")

        st.caption("Movements and counts are stored in Supabase and survive app redeploys. "
                   "Auto-deduction uses the spreadsheet's Meter Commission Date and Replacement Date "
                   "columns — keep those up to date and stock stays accurate by itself.")