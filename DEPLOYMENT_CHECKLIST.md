# Deployment Checklist — v3

- [ ] Create Google Sheet database workbook.
- [ ] Add Apps Script files from `apps-script/`.
- [ ] Set Script Property `SPREADSHEET_ID`.
- [ ] Run `setupWorkbook()` once.
- [ ] Confirm `OFFICE_MASTER` contains the required SOL IDs and office names.
- [ ] Add SPM users to `USER_SET`.
- [ ] Add one or more Admin users to `USER_SET` with `ROLE=ADMIN`.
- [ ] Deploy Apps Script as Web App and copy `/exec` URL.
- [ ] Upload this project to GitHub.
- [ ] Enable GitHub Pages through Actions.
- [ ] Open the GitHub Pages site and save the Apps Script backend URL.
- [ ] Test SPM login, office auto-mapping and today's duplicate protection.
- [ ] Test kit reconciliation and article reconciliation.
- [ ] Test Admin filters, pending offices, sessions and CSV export.
