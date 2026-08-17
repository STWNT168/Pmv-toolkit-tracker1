/*
 * PMV Toolkit Tracker — frontend configuration
 *
 * Paste your deployed Google Apps Script Web App /exec URL below.
 *
 * IMPORTANT SECURITY NOTE:
 * The Apps Script URL is not a password.
 * A Google Spreadsheet ID is also not a password, but because this
 * repository is public, the recommended secure location for the REAL
 * spreadsheet ID is Apps Script > Project Settings > Script Properties.
 *
 * The SPREADSHEET_ID field is provided here as a clearly marked place
 * for your deployment reference only. The frontend does NOT send it
 * to the backend and it is NOT used to open the sheet.
 */

window.PMV_CONFIG = Object.freeze({
  APP_NAME: 'PMV Toolkit Tracker',
  APP_SCRIPT_API_URL: '',
  SPREADSHEET_ID: '',

  REQUEST_TIMEOUT_MS: 30000,
  SESSION_STORAGE_KEY: 'pmvSessionV6',
  BACKEND_STORAGE_KEY: 'pmvBackendUrlV6',

  // Set true only if you want to allow the browser setup screen
  // to override APP_SCRIPT_API_URL. Default is false for consistency.
  ALLOW_BACKEND_OVERRIDE: false
});
