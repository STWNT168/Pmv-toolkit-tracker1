# SPM Tool Kits Daily Reporting — GitHub Pages v5

## Architecture
- Frontend: GitHub Pages, `index.html` at repository root
- Backend: Google Apps Script Web App
- Database: one Google Sheets workbook
- Sheets: `OFFICE_MASTER`, `USER_SET`, `DAILY_RECORD`, `SESSION_LOG`

## Main v5 update: Admin consolidated report
The Admin dashboard now provides a consolidated report across all selected records.

### Kits
- All Kits
- Invalid Mobile Kits
- Deliverable Kits
- Incomplete Kits
- Without Proper Details/Address Kits

### Articles
- Similar / Tool Kit Articles
- Invalid Mobile Articles
- Deliverable Articles
- Incomplete Articles
- Without Proper Details/Address Articles

The dashboard also shows:
- Total offices
- Submitted offices
- Pending offices
- Office-wise report containing every kit and article field
- Date range / specific date / office filters
- Session/IP audit
- CSV export

All consolidated totals are calculated server-side from `DAILY_RECORD`.

## SPM
- Employee ID + registered phone login
- Automatic SOL ID and office mapping
- Separate kit and article fields
- Server-side reconciliation
- Server-side duplicate prevention
- 7-day maximum session
- Previous submission history

## Server validation
`ALL_KITS = Invalid Mobile Kits + Deliverable Kits + Incomplete Kits + Without Proper Details/Address Kits`

`SIMILAR_ARTICLE = Invalid Mobile Articles + Deliverable Articles + Incomplete Articles + Without Proper Details/Address Articles`

## Deploy
1. Create a Google Sheet.
2. Open Extensions → Apps Script.
3. Copy `apps-script/Code.gs` and `apps-script/appsscript.json`.
4. Add Script Property `SPREADSHEET_ID` with your Google Sheet ID.
5. Run `setupWorkbook()` once.
6. Fill `USER_SET`.
7. Deploy Apps Script as a Web App.
8. Copy the `/exec` URL.
9. Upload this repository to GitHub with `index.html` at the root.
10. Enable GitHub Pages using the included Actions workflow.
11. Open the Pages URL and paste the Apps Script `/exec` URL into Backend setup.

## USER_SET
`EMPLOYEE_ID | PHONE | ROLE | SOL_ID | ACTIVE | NAME | SESSION_DAYS`

Use `ROLE=SPM` or `ROLE=ADMIN`. Keep session validity at 7 days or less.
