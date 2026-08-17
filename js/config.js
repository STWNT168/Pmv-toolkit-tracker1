/*
 * PMV Toolkit Tracker — frontend configuration
 * SECURITY: keep the real spreadsheet ID in Apps Script Script Properties.
 */
window.PMV_CONFIG = Object.freeze({
  APP_NAME: 'PMV Toolkit Tracker',
  APP_SCRIPT_API_URL: 'https://script.google.com/macros/s/AKfycbzdTwEfyuolUkyr9ME0_4u4iQmED25GMvyqqBcHFFPwehX355nCwn4jFhbmF7ILVHaD/exec',

  // Reference only. The frontend never opens the spreadsheet.
  // Keep the real value in Apps Script > Project Settings > Script Properties.
  SPREADSHEET_ID: '',

  REQUEST_TIMEOUT_MS: 45000,
  SESSION_STORAGE_KEY: 'pmvSessionV7',
  ALLOW_BACKEND_OVERRIDE: false
});
