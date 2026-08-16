# Life & Work Planner — setup guide

Same split as Toka: Apps Script is a pure backend API (Google Sheets is the
database), and Index.html is a static page hosted on GitHub Pages that talks
to that API over `fetch()`. Two separate deployments, ~10 minutes total.

## 1. Create the spreadsheet

1. Go to [sheets.google.com](https://sheets.google.com) → Blank spreadsheet.
2. Name it whatever you like, e.g. "Life & Work Planner Data".

## 2. Add the backend

1. In the sheet, go to **Extensions → Apps Script**.
2. Delete the placeholder `Code.gs` content and paste in **Code.gs**.
3. That's the only file this project needs now — no HTML file goes in Apps
   Script anymore.

## 3. Run setup once

1. In the function dropdown, select **setup**, then click **Run** (▶).
2. First run: Google asks you to authorize — click through (Advanced → Go
   to [project name] (unsafe) is expected for your own unpublished script).
3. Check the "Tasks" and "Projects" tabs appeared in your spreadsheet with
   headers and the default project list (Toka, House, Yoga, Arabic, Life,
   Matcha, Other — edit colors/add more directly in the Projects sheet).

This also installs a daily 6am trigger that generates the next instance of
any recurring task automatically.

## 4. Deploy the backend as a web app

1. **Deploy → New deployment** → gear icon → **Web app**.
2. Execute as: **Me**. Who has access: **Anyone** — this is required, not
   optional, since GitHub Pages calling in is a different origin than
   "Only myself" allows.
3. Click **Deploy**, authorize again if asked.
4. Copy the **Web app URL** (ends in `/exec`) — this is your API URL.

## 5. Host the frontend on GitHub Pages

1. On github.com (account **claire-vr**), create a new repository — e.g.
   `life-work-planner`. Public repos get free Pages hosting.
2. Open **index.html** locally, find this line near the top of the
   `<script>` block:
   ```js
   const API_URL = 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
   ```
   Replace the placeholder with the `/exec` URL you copied in step 4.
3. Upload the edited `index.html` to the new repo (Add file → Upload files
   on github.com works fine, no git command line needed). Make sure the
   filename is lowercase `index.html` — GitHub Pages looks for that exact
   name as the site's default page.
4. In the repo: **Settings → Pages** → Source: **Deploy from a branch** →
   Branch: `main` / `(root)` → Save.
5. GitHub gives you a URL like `https://claire-vr.github.io/life-work-planner/`
   — that's your planner, live in a minute or two.

## 6. Apple Calendar

This is one-way, import only: click **🍏 Import from Apple** in the planner
and paste a public Apple Calendar share link (Calendar app → right-click a
calendar → Share Calendar → check "Public Calendar" → copy the link). Events
from that calendar get pulled into the planner; nothing is pushed back out
to Apple Calendar.

## Notes on how it works

- **Two separate deployments now**: editing Code.gs requires redeploying the
  Apps Script web app (or using its `/dev` URL) for changes to take effect.
  Editing index.html requires re-uploading to GitHub (or `git push` if you
  set up the repo with git locally instead of the web upload flow).
- **Recurring tasks** live in the same "Tasks" sheet, flagged `IsTemplate =
  TRUE`. The daily trigger checks each template's `NextRunDate`; if it's
  arrived, a real task row is created and the template's date advances.
- **Projects and colors** are just rows in the "Projects" sheet — add,
  rename, or recolor by editing that sheet directly.
- If the page loads but nothing appears and the browser console shows a
  CORS or network error, double-check the Apps Script deployment's access
  is set to "Anyone" (not "Anyone with Google account" or "Only myself").

## If something breaks

Apps Script keeps an execution log: in the Apps Script editor, click the
clock icon on the left ("Executions") to see errors from the daily trigger
or from API calls made by the page. For frontend issues, open the browser's
developer console (View → Developer → JavaScript Console in Safari) — fetch
errors show up there. Paste any error text back to me and I can debug it.
