# PMV Toolkit Tracker — Refined Secure Build

## Architecture

- Frontend: GitHub Pages
- Backend: Google Apps Script Web App
- Database: one Google Sheets workbook
- Frontend configuration: `js/config.js`

## Important configuration

### 1. Apps Script spreadsheet

Open Apps Script → Project Settings → Script Properties and create:

`SPREADSHEET_ID = YOUR_GOOGLE_SHEET_ID`

The backend deliberately reads the spreadsheet ID from Script Properties instead of trusting a browser request.

### 2. Run workbook setup

Run `setupWorkbook()` once from Apps Script.

It creates:

- `OFFICE_MASTER`
- `USER_SET`
- `DAILY_RECORD`
- `SESSION_LOG`

Then populate `USER_SET`:

`EMPLOYEE_ID | PHONE | ROLE | SOL_ID | ACTIVE | NAME | SESSION_DAYS`

Use `ROLE=SPM` or `ROLE=ADMIN`.

### 3. Deploy Apps Script

Deploy as Web App:

- Execute as: Me / User deploying the web app
- Who has access: Anyone

Copy the generated `/exec` URL.

### 4. Configure GitHub Pages

Open:

`js/config.js`

Set:

`APP_SCRIPT_API_URL: 'https://script.google.com/macros/s/XXXXX/exec'`

`SPREADSHEET_ID` is provided as a reference field, but the backend does not use the browser value. Keep the real spreadsheet ID in Apps Script Script Properties.

## Security improvements in this build

- Server-side session validation
- Role-based SPM/Admin authorization
- Session expiry capped at 7 days
- Server-side duplicate submission prevention
- Script Lock around submission
- Server-side numeric range validation
- Server-side kit reconciliation
- Server-side article reconciliation
- Login rate limiting
- Inactive-user blocking
- Session/IP/user-agent audit
- HTML escaping on dashboard output
- Request IDs for iframe responses
- No direct Google Sheet access from browser
- No spreadsheet ID used as an authentication credential

## Known inconsistency fixed

The previous frontend contained:

`async async function loadDashboard()`

This is invalid JavaScript and can stop the entire page script from parsing.

The refined build contains:

`async function loadDashboard()`

## Deployment check

1. Configure Script Property `SPREADSHEET_ID`.
2. Run `setupWorkbook()`.
3. Fill `USER_SET`.
4. Deploy Apps Script and copy `/exec`.
5. Put the `/exec` URL in `js/config.js`.
6. Push the files to GitHub.
7. Enable GitHub Pages.
8. Test SPM login.
9. Test one valid submission.
10. Confirm duplicate submission is blocked.
11. Test Admin login.
12. Test consolidated report and pending-office list.
