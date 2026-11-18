# Test Equity – Cash Application report updater

MIS selects the two agents' EOD files, clicks **Update report**, and the app:

1. Reads each EOD sheet and cleans it (agent names, `$` amounts, dates, tick boxes).
2. Appends the rows to the right sheets of the master report, with Month and Week filled in.
3. Sets the Daily date (cell B6) on the current month sheet, e.g. `Nov'26`.
4. Saves the updated report back to a private GitHub repo (every day's version is kept) and offers a download.

Daily, Weekly and the month sheets are formulas, so Excel updates them when the file is opened.

## Where each EOD sheet goes

| EOD file / sheet | Report sheet | Task written |
|---|---|---|
| Varsha / Billtrust | Billtrust | Billtrust Review |
| Varsha / Key offs | Keyoff | Keyoff |
| Varsha / Unapplied cash | Unapplied Cash | Unapplied Cash |
| Gokulnath / CCP_HISCO | Credit Card | Credit Card Processing (Hisco) |
| Gokulnath / CCP_TE | Credit Card | Credit Card Processing (Test Equity) |
| Gokulnath / TE_ MATCH REVIEW | Billtrust | Billtrust Review (Task Type = Match Review, Status = Completed) |

Not covered because the sample EOD files had no sheet for them: **COA, Unposted Report, Correlation HISCO**.
Send sample files and the engine can be extended.

## Setup (about 45 minutes, once)

Use the same pattern as your other apps: a **public** repo with the code only, and a **private** repo with the report only.

### Part 1 – Public code repo
1. github.com → **+** → **New repository** → name `cash-app-app` → **Public** → Create.
2. **Add file → Upload files** → upload `app.py`, `ca_report.py`, `requirements.txt`, `README.md` → **Commit changes**.
   Do NOT upload the report here.

### Part 2 – Private data repo
3. **+** → **New repository** → name `cash-app-data` → **Private** → Create.
4. Rename your latest report to exactly `report.xlsx` (keep a backup copy of the original).
5. **Add file → Upload files** → upload `report.xlsx` → **Commit changes**. (Your file is about 5 MB, well under GitHub's 25 MB web limit.)
6. Check the page shows a **Private** label and lists `report.xlsx`.

### Part 3 – Token for the private repo
7. Profile picture → **Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token**.
8. Name: a new unused name, e.g. `cash-app-token`. Expiration: 90 days (set a calendar reminder).
9. Repository access: **Only select repositories** → `cash-app-data`.
10. Permissions → Repository permissions → **Contents: Read and write**.
11. **Generate token**, copy it (starts `github_pat_`). You only see it once.

### Part 4 – Deploy on Streamlit
12. share.streamlit.io → **Create app** → **Deploy a public app from GitHub**.
13. Repository `yourname/cash-app-app`, Branch `main`, Main file path `app.py`.
14. **Advanced settings → Secrets**, paste (straight quotes, your own values):

```
APP_PASSWORD = "choose-a-long-password"

[GITHUB]
token = "github_pat_PASTE_YOUR_TOKEN"
repo = "yourname/cash-app-data"
path = "report.xlsx"
branch = "main"
```
15. **Deploy** and wait a few minutes. The link you get is the one MIS uses.

## Test before real use
- The real 16-Nov EOD files will be refused with *"Already loaded"*, because your report already has 16-Nov. That is the duplicate guard working.
- For a real test, use the next day's files. Check the Nov'26, Daily and Weekly sheets in Excel afterwards.

## Daily routine
1. Open the app link, enter the password.
2. Select both EOD files together (Ctrl+click) and check the table (sheet, agent, date, rows).
3. Click **Update report**. Read any yellow warnings. Wait for **Saved back to GitHub**.
4. Click **Download updated report**, open it in Excel, share it.

Either agent's file can be loaded on its own. The guard checks date + agent + sheet.

## New month (for example 1 December)
The app needs a sheet named `Dec'26` to exist (it sets the Daily date there). Do what you do today:
copy the `Nov'26` sheet, rename it `Dec'26`, change the month text in cell A2, and update the
month text/start date on the Weekly and Daily sheets. Then upload that file to the private repo
as `report.xlsx` **before** loading the day's EOD files.
Until the sheet exists the app warns you and still adds the rows.

## If something goes wrong
- **Restore:** upload a good copy named exactly `report.xlsx` to the private repo (**Add file → Upload files**), or use the file's **History** to download an earlier version.
- **GitHub 404:** file must be exactly `report.xlsx` at the top level; `repo` in Secrets must match the address bar; the token must cover `cash-app-data`.
- **GitHub 401:** token wrong or expired. Make a new one and paste it in Secrets (Manage app → ⋮ → Settings → Secrets).
- **Save fails with 409:** someone saved a moment earlier. Refresh the page and run again.
- **Yellow "agent is not one of ..."**: a new name. Rows are still added as typed; check them. A new agent also needs a new column on the month sheets.
- **Yellow "New status ..."**: an agent typed a spelling the report has not used before. Rows are added as typed.
- Never share the token. Renew it before it expires, or saving stops.
