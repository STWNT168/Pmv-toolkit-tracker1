# SPM Tool Kits Daily Reporting Web App — GitHub Pages v3

## Architecture
- Frontend: GitHub Pages (`index.html`)
- Backend: Google Apps Script Web App (`apps-script/Code.gs`)
- Database: Google Sheets workbook

## v3 features
- SPM dashboard and Admin dashboard
- Employee ID + phone login
- Employee ID automatically maps to SOL ID and office name
- Daily kit + article entry
- `ALL_KITS` = Invalid Mobile Kits + Deliverable Kits + Incomplete Kits + Without Proper Details Kits
- `SIMILAR_ARTICLE` = Invalid Mobile Articles + Deliverable Articles + Incomplete Articles + Without Proper Details Articles
- Server-side duplicate check for same date + employee + SOL ID
- 7-day session validity
- Session log with IP, user-agent, login/logout/last-seen and expiry
- SPM previous-submission history
- Admin filters, totals, completion/pending offices, session audit and CSV export
- Office master seeded from the uploaded office/SOL workbook list

## Deployment
1. Create a Google Sheet and copy its ID.
2. Open Extensions → Apps Script.
3. Paste `apps-script/Code.gs` and `apps-script/appsscript.json`.
4. In Apps Script Project Settings → Script Properties add `SPREADSHEET_ID` = your Google Sheet ID.
5. Run `setupWorkbook()` once and authorize it.
6. Add SPM/Admin credentials to `USER_SET`.
7. Deploy → New deployment → Web app. Execute as you; access as anyone who needs the portal.
8. Copy the `/exec` URL.
9. Create a GitHub repository and upload this project.
10. Enable GitHub Pages using GitHub Actions. The included workflow deploys `/docs`.
11. Open the GitHub Pages URL and paste the Apps Script `/exec` URL in Backend setup.

## USER_SET
Columns:
`EMPLOYEE_ID | PHONE | ROLE | SOL_ID | ACTIVE | NAME | SESSION_DAYS`

Use `ROLE=SPM` for SPM users and `ROLE=ADMIN` for administrators. Keep `SESSION_DAYS` at 7 or lower.

## Important security note
GitHub Pages is only the public frontend. Credentials, authorization, duplicate checking and data validation are enforced by Apps Script. The browser sends a client-observed public IP from `api.ipify.org`; this is audit metadata and should not be treated as a cryptographic identity signal.

## GitHub Pages layout

`index.html` is intentionally in the **repository root**. You do not need a `docs` folder.

The SPM/Admin portal lets you enter the deployed Apps Script `/exec` URL on the first screen. The URL is saved in the browser's local storage. This means you can deploy the same GitHub Pages files and configure the backend URL from the website without editing the HTML.
