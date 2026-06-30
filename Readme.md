# Sitari Evergreen — Meter Commissioning (Streamlit app)

A dashboard your technicians can open to see what's outstanding, upcoming, overdue,
and already installed — based on the snag-list deadlines in your commissioning
spreadsheet.

## How it gets its data
This app does **not** have an upload button. It reads whichever file matching
`EVG_SIT_FS_Meter_Commissioning_*.xlsx` is sitting in this same folder, picking the
most recently modified one if there's more than one. That means your workflow is:

1. Update the spreadsheet as techs install meters.
2. Save it into this project folder (replacing or alongside the old copy).
3. Commit and push to GitHub:
   ```
   git add EVG_SIT_FS_Meter_Commissioning_*.xlsx
   git commit -m "Update meter commissioning data"
   git push
   ```
4. If deployed on Streamlit Cloud, it redeploys automatically within a minute or two.
   Technicians just refresh the page to see the latest data.

## What it shows
- **Outstanding** — every meter not yet installed, filterable by type/section/stand.
- **Upcoming** — due within 14 days, plus a further-ahead view grouped by week.
- **Overdue** — past Snag Date 4 and still not installed, with days overdue.
- **Installed** — full installation log, flags any installed after their deadline.
- **Calendar** — a month-by-month calendar showing installed (✅), due soon (🟧),
  and overdue (🟥) days at a glance.
- **Sections** — progress by WBHO subsection and by unit type (A1/A2/A3/B2/B3).

Each table view has a CSV download button for anyone who wants a list to take
into the field.

## Run it yourself (quickest way to test)
1. Install Python 3.10+ if you don't have it.
2. In this folder, run:
   ```
   pip install -r requirements.txt
   streamlit run app.py
   ```
3. It opens in your browser at `http://localhost:8501`.

## Share it with your technicians
**Streamlit Community Cloud (free, easiest to share a link)**
1. Push this folder (`app.py`, `requirements.txt`, and the spreadsheet) to a GitHub repo.
2. Go to https://share.streamlit.io, sign in, and deploy from that repo.
3. You'll get a shareable link (e.g. `https://yourapp.streamlit.app`) — send that to your
   technicians. Restrict viewers via Streamlit Cloud's sharing settings if you don't
   want it public, since the repo will contain your project data.

**Run it on a shared computer/server on site**
1. Run `streamlit run app.py --server.address 0.0.0.0` on a PC on your site network.
2. Technicians on the same network browse to `http://<that PC's IP>:8501`.

## Notes
- The app expects the column names from your current sheet (Stand Number, Section,
  WBHO Subsection, Meter Commissioning/Commission Date, AMR Commissioned/Installed,
  Snag Date 4). If you rename a column, update the matching list near the top of
  `app.py` (`coalesce_col` calls).
- "Deadline" is always Snag Date 4 — the latest date you can deliver installations by.
- Because the spreadsheet now lives in the repo, keep the repo private if it contains
  anything you wouldn't want public.