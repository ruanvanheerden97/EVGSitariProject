# Sitari Evergreen — Outstanding Meters (Streamlit app)

A simple web app your technicians can open to see which meters are still outstanding,
based on the snag-list deadlines in your commissioning spreadsheet.

## What it does
- Upload the latest `EVG_SIT_FS_Meter_Commissioning_*.xlsx` (any filename/date works).
- Reads the **Water meters** and **Elec Meters** tabs.
- Shows outstanding meters, due-soon (within 14 days), and overdue meters.
- Filterable by meter type, section, unit type, and stand number — so a technician
  can pull up just "what's due in Section 9" or "what's overdue, electrical only".
- Each technician can download the filtered list as a CSV.
- Nothing is saved on the server — every session starts fresh and only uses the file
  that's uploaded in that session, so you always control which version is in use.

## Run it yourself (quickest way to test)
1. Install Python 3.10+ if you don't have it.
2. In this folder, run:
   ```
   pip install -r requirements.txt
   streamlit run app.py
   ```
3. It opens in your browser at `http://localhost:8501`. Upload your spreadsheet there.

## Share it with your technicians
You have two practical options:

**Option A — Streamlit Community Cloud (free, easiest to share a public/private link)**
1. Push this folder (`app.py` + `requirements.txt`) to a GitHub repo.
2. Go to https://share.streamlit.io, sign in, and deploy from that repo.
3. You'll get a shareable link (e.g. `https://yourapp.streamlit.app`) — send that to your
   technicians. They open it on a phone or laptop, upload the current spreadsheet, and
   see the outstanding list. You can restrict who can view it via Streamlit Cloud's
   sharing/viewer settings if you don't want it public.

**Option B — Run it on a shared computer/server on site**
1. Run `streamlit run app.py --server.address 0.0.0.0` on a PC connected to your site
   network/Wi-Fi.
2. Technicians on the same network browse to `http://<that-PC's-IP>:8501`.
3. This keeps everything local with no internet dependency, but only works while that
   PC is on and everyone's on the same network.

Either way, the workflow stays the same: whenever you update the spreadsheet, re-upload
it in the app (or have whoever's managing the deployed link do that) — technicians then
just refresh the page.

## Notes
- The app expects the column names from your current sheet (Stand Number, Section,
  WBHO Subsection, Meter Commissioning/Commission Date, AMR Commissioned/Installed,
  Snag Date 1–4). If you rename a column, update the matching list near the top of
  `app.py` (`coalesce_col` calls).
- "Deadline" in the app is always Snag Date 4, matching what you described as the
  latest date you can deliver installations by.