"""
Standalone AMR reading fetcher — runs hourly via GitHub Actions,
independent of whether the Streamlit app is open.

Connects to the SFTP server, finds every CSV newer than the most recent
file already in the Supabase history DB (or the last 24h if the DB is
empty), parses them, and upserts the readings.

Required environment variables (set as GitHub Actions secrets):
  SFTP_HOST, SFTP_PORT, SFTP_USERNAME, SFTP_PASSWORD, SFTP_DIRECTORY
  DB_URL   (Supabase pooler connection string)
"""
import io
import os
import re
import sys
from datetime import datetime, timedelta

import pandas as pd
import paramiko
import psycopg2
import psycopg2.extras

CSV_FILENAME_RE = re.compile(r"^[A-Z]+_(\d{8})_(\d{6})\.csv$", re.IGNORECASE)


def parse_csv_filename_ts(name):
    m = CSV_FILENAME_RE.match(os.path.basename(name))
    if m:
        try:
            return datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S")
        except ValueError:
            pass
    return None


def parse_amr_csv(csv_bytes, file_ts):
    text = csv_bytes.decode("utf-8", errors="replace")
    df = pd.read_csv(io.StringIO(text))
    df.columns = [c.strip() for c in df.columns]
    df["_addr"] = df["METER_ADDRESS"].astype(str).str.strip()
    df = df[~df["_addr"].str.endswith("NR")].copy()
    df["_reading_dt"] = pd.to_datetime(
        df["READING_DATE"].astype(str).str.replace(r"\s*GMT[+-]\d+", "", regex=True),
        dayfirst=True, errors="coerce",
    )
    rows = []
    for _, r in df.iterrows():
        serial = r["_addr"]
        if not serial or serial == "nan" or pd.isna(r["_reading_dt"]):
            continue
        rows.append((
            serial,
            r["_reading_dt"].isoformat(),
            float(r["READING_VALUE"]) if pd.notna(r["READING_VALUE"]) else None,
            int(r.get("LOW_BATTERY", 0) or 0),
            file_ts.isoformat() if file_ts else None,
        ))
    return rows


def main():
    host = os.environ["SFTP_HOST"]
    port = int(os.environ.get("SFTP_PORT", 22))
    user = os.environ["SFTP_USERNAME"]
    pwd = os.environ["SFTP_PASSWORD"]
    directory = os.environ["SFTP_DIRECTORY"]
    db_url = os.environ["DB_URL"]

    conn = psycopg2.connect(db_url, sslmode="require", connect_timeout=15)
    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS amr_readings (
                serial        TEXT NOT NULL,
                reading_date  TEXT NOT NULL,
                reading_value DOUBLE PRECISION,
                low_battery   INTEGER DEFAULT 0,
                file_ts       TEXT,
                PRIMARY KEY (serial, reading_date)
            )""")
        cur.execute("SELECT MAX(file_ts) FROM amr_readings WHERE file_ts IS NOT NULL")
        row = cur.fetchone()
    conn.commit()

    last_ts = datetime.fromisoformat(row[0]) if row and row[0] else None
    cutoff = last_ts if last_ts else datetime.now() - timedelta(hours=24)
    print(f"DB last file_ts: {last_ts} — fetching files newer than {cutoff}")

    transport = paramiko.Transport((host, port))
    transport.connect(username=user, password=pwd)
    sftp = paramiko.SFTPClient.from_transport(transport)

    files = [f for f in sftp.listdir(directory) if f.lower().endswith(".csv")]
    new_files = sorted(
        [f for f in files if (parse_csv_filename_ts(f) or datetime.min) > cutoff],
        key=lambda f: parse_csv_filename_ts(f) or datetime.min,
    )
    print(f"{len(files)} CSVs on server, {len(new_files)} newer than cutoff")

    total_rows = 0
    for fname in new_files:
        file_ts = parse_csv_filename_ts(fname)
        buf = io.BytesIO()
        sftp.getfo(directory.rstrip("/") + "/" + fname, buf)
        rows = parse_amr_csv(buf.getvalue(), file_ts)
        if rows:
            with conn.cursor() as cur:
                psycopg2.extras.execute_values(
                    cur,
                    """INSERT INTO amr_readings
                       (serial, reading_date, reading_value, low_battery, file_ts)
                       VALUES %s ON CONFLICT (serial, reading_date) DO NOTHING""",
                    rows, page_size=500,
                )
            conn.commit()
        total_rows += len(rows)
        print(f"  {fname}: {len(rows)} readings")

    transport.close()

    with conn.cursor() as cur:
        cur.execute("SELECT COUNT(*), COUNT(DISTINCT serial), MAX(file_ts) FROM amr_readings")
        n, s, mx = cur.fetchone()
    conn.close()
    print(f"Done. Processed {len(new_files)} files / {total_rows} rows. "
          f"DB now: {n:,} readings, {s} serials, latest file {mx}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"FETCH FAILED: {e}", file=sys.stderr)
        sys.exit(1)