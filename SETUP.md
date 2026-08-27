# CP Pendency — auto-updating dashboard (Path B)

Gmail → GitHub Actions (Python) → Drive `agg.json` → Apps Script dashboard.
One-time setup, then it refreshes itself every 30 min.

## What you'll create
- 1 Google Cloud project (for the API credential)
- 1 OAuth refresh token (minted once on your laptop)
- 3 Drive items: `agg.json`, `cp_state.json`, and the Apps Script web app
- 1 GitHub repo (holds the code + runs the scheduler for free)

---

## Step 1 — Google Cloud credential (10 min, once)
1. console.cloud.google.com → create a project (e.g. "cp-pendency").
2. APIs & Services → **Enable**: Gmail API, Google Drive API, Google Sheets API.
3. OAuth consent screen → **External** → add your own Gmail as a **Test user**.
4. Credentials → Create credentials → **OAuth client ID** → **Desktop app** →
   download JSON → save as `client_secret.json` next to `get_token.py`.

## Step 2 — Mint the refresh token (once, on your laptop)
```
pip install google-auth-oauthlib
python get_token.py
```
Log in as the mailbox that RECEIVES the CP emails. It prints three values —
keep them for Step 4.

## Step 3 — Create the two Drive files
1. In Drive, create two empty files named `agg.json` and `cp_state.json`
   (right-click → new → or upload a file containing just `{}`).
2. Open each, copy its **file ID** from the URL
   (`drive.google.com/file/d/<THIS_PART>/view`).

## Step 4 — GitHub repo + secrets
1. Create a private repo, push these files (keep the folder structure).
2. Repo → Settings → Secrets and variables → **Actions** → add:
   - `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `GOOGLE_REFRESH_TOKEN` (from Step 2)
   - `DRIVE_AGG_FILE_ID`, `DRIVE_STATE_FILE_ID` (from Step 3)
   - `GMAIL_QUERY` — the search that finds the mail, e.g.
     `from:reports@xpressbees.com subject:(CP) has:attachment filename:xlsb newer_than:2d`
   - `SHEET_ID` — optional; a Sheet with a tab named `Trend` for history. Leave unset to skip.
3. Actions tab → run **CP Pendency refresh** once manually (workflow_dispatch)
   to confirm it works. Check the log says "Updated dashboard from …".

## Step 5 — Deploy the dashboard
1. script.google.com → New project → paste `Code.gs` and add an HTML file
   named exactly `Index` with the contents of `apps-script/Index.html`.
2. In `Code.gs` set `AGG_FILE_ID` to your `agg.json` file ID.
3. Deploy → New deployment → **Web app** → Execute as **Me**,
   Access **Anyone within <your org>** → copy the URL. That's your live dashboard.

---

## Tuning
- **Schedule**: edit the cron in `.github/workflows/cp.yml`. Current = every 30 min,
  08:00–22:00 IST. GitHub cron is UTC; IST = UTC+5:30.
- **If the report's sheet/tab name changes** each week (e.g. "CP XB Aug 27-29"),
  tell me — the parser currently targets that exact tab; I can make it auto-detect
  the second sheet instead so it never breaks on the date.
- **Trend line**: set `SHEET_ID` and the job logs one row per new snapshot — easy to
  chart pendency over the day.
