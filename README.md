# PMV Toolkit Tracker v8

A clean rebuild of the PMV Toolkit Tracker while keeping the original essence:

- GitHub Pages frontend
- Google Apps Script backend
- Google Sheets database
- SPM login with Employee ID + registered phone
- Server-side authentication
- Session management
- Daily PMV toolkit/article reporting
- Duplicate submission protection
- Kit/article total validation
- Admin dashboard
- Office-wise submitted/pending status
- Consolidated totals
- Session audit
- CSV export
- Responsive mobile UI

## Backend setup

1. Create/open the Google Spreadsheet.
2. Open Apps Script.
3. Add `Code.gs` from `apps-script/Code.gs`.
4. Add/update `appsscript.json`.
5. Open **Project Settings → Script Properties**.
6. Add:

   `SPREADSHEET_ID = YOUR_SPREADSHEET_ID`

7. Run `setupWorkbook()` once and authorize the script.
8. Add users to `USER_SET`.

### USER_SET columns

`EMPLOYEE_ID | PHONE | ROLE | SOL_ID | ACTIVE | NAME | SESSION_DAYS`

Example:

`10326670 | 9876543210 | SPM | 18231301 | TRUE | Example Name | 7`

Use your real authorized data in the spreadsheet. Do not put real credentials in GitHub.

## Web App deployment

Deploy:

- Type: Web app
- Execute as: Me
- Who has access: Anyone

Copy the `/exec` URL and put it in:

`js/config.js`

The supplied config already contains the `/exec` URL provided for this rebuild.

## GitHub Pages

Upload the contents of this project to the repository root.

The GitHub Pages URL should load:

`index.html`

After deployment, use a fresh/incognito browser tab if the previous site cached old JavaScript.

## Important architecture change

The old project used a hidden iframe and an Apps Script HTML response as a bridge. This rebuild keeps that transport because direct browser `fetch()` to an Apps Script web app can run into cross-origin/redirect limitations.

The response bridge now posts explicitly to `window.parent` and has a fallback to `window.top`.

The frontend also has a server health action so deployment/configuration problems are easier to diagnose.

## After every Code.gs change

Apps Script code is deployed separately from GitHub.

Always:

**Deploy → Manage deployments → Edit → New version → Deploy**

Then reload GitHub Pages.

## Sheets

The backend creates/uses:

- `OFFICE_MASTER`
- `USER_SET`
- `DAILY_RECORD`
- `SESSION_LOG`

Do not rename these sheets without changing `Code.gs`.
