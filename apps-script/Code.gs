/**
 * PMV Toolkit Tracker v8
 * Secure Google Apps Script backend.
 *
 * REQUIRED:
 *   Apps Script > Project Settings > Script Properties
 *   SPREADSHEET_ID = your Google Spreadsheet ID
 *
 * Run setupWorkbook() once after adding SPREADSHEET_ID.
 *
 * Deploy as:
 *   Web app
 *   Execute as: Me
 *   Who has access: Anyone
 */

const CFG = Object.freeze({
  SESSION_DAYS_DEFAULT: 7,
  SESSION_DAYS_MAX: 7,
  MAX_VALUE: 1000000,
  MAX_RECORDS: 100,
  MAX_SESSIONS: 1000,
  LOGIN_WINDOW_MS: 10 * 60 * 1000,
  LOGIN_MAX_ATTEMPTS: 8
});

function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error(
      'Backend is not configured. Add SPREADSHEET_ID in Apps Script Script Properties.'
    );
  }
  return SpreadsheetApp.openById(id);
}

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || 'health');

  if (action === 'health') {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        service: 'PMV Toolkit Tracker API',
        version: '8.0.0',
        time: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService
    .createHtmlOutput(
      '<!doctype html><html><head><base target="_top"></head>' +
      '<body><h3>PMV Toolkit Tracker API</h3>' +
      '<p>Backend is running.</p></body></html>'
    )
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const requestId =
    String((e && e.parameter && e.parameter.requestId) || Utilities.getUuid());

  try {
    const action = String((e && e.parameter && e.parameter.action) || '');
    const rawPayload =
      e && e.parameter && e.parameter.payload
        ? e.parameter.payload
        : '{}';

    const payload = JSON.parse(rawPayload);
    let result;

    switch (action) {
      case 'health':
        result = {
          ok: true,
          service: 'PMV Toolkit Tracker API',
          version: '8.0.0',
          time: new Date().toISOString()
        };
        break;

      case 'login':
        result = login(
          payload.employeeId,
          payload.phone,
          payload.ip,
          payload.userAgent
        );
        break;

      case 'getTodayStatus':
        result = getTodayStatus(payload.sessionId);
        break;

      case 'getMyRecords':
        result = getMyRecords(payload.sessionId, payload.limit);
        break;

      case 'submitDaily':
        result = submitDaily(payload.sessionId, payload.report);
        break;

      case 'adminSummary':
        result = adminSummary(payload.sessionId, payload.filters || {});
        break;

      case 'adminSessions':
        result = adminSessions(payload.sessionId, payload.limit);
        break;

      case 'getOffices':
        result = getOffices(payload.sessionId);
        break;

      case 'exportDailyCsv':
        result = exportDailyCsv(
          payload.sessionId,
          payload.filters || {}
        );
        break;

      case 'logout':
        result = logout(payload.sessionId);
        break;

      default:
        throw new Error('Unknown API action: ' + action);
    }

    return bridgeResponse_(requestId, true, result);
  } catch (err) {
    return bridgeResponse_(requestId, false, {
      message: String(err && err.message || err)
    });
  }
}

/*
 * IMPORTANT:
 * GitHub Pages calls the Apps Script Web App through a hidden iframe.
 * Apps Script HtmlService pages are iframe-sandboxed, so the response
 * explicitly posts to the immediate parent window.
 */
function bridgeResponse_(requestId, ok, data) {
  const message = JSON.stringify({
    requestId: requestId,
    ok: ok,
    data: data
  });

  const safe = message
    .replace(/\\/g, '\\\\')
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');

  const html =
    '<!doctype html><html><head><base target="_top"></head><body>' +
    '<script>' +
    '(function(){' +
    'var msg=' + safe + ';' +
    'try{' +
    'window.parent.postMessage(msg,"*");' +
    '}catch(e){' +
    'try{window.top.postMessage(msg,"*");}catch(ignore){}' +
    '}' +
    '})();' +
    '</script>' +
    '</body></html>';

  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupWorkbook() {
  const ss = getSS_();

  const specs = {
    OFFICE_MASTER: ['SOL_ID', 'OFFICE_NAME', 'ACTIVE'],

    USER_SET: [
      'EMPLOYEE_ID', 'PHONE', 'ROLE', 'SOL_ID',
      'ACTIVE', 'NAME', 'SESSION_DAYS'
    ],

    DAILY_RECORD: [
      'TIMESTAMP', 'REPORT_DATE', 'EMPLOYEE_ID', 'PHONE_LAST4',
      'SOL_ID', 'OFFICE_NAME',
      'ALL_KITS', 'SIMILAR_ARTICLE',
      'INVALID_MOBILE_KITS', 'INVALID_MOBILE_ARTICLES',
      'DELIVERABLE_KITS', 'DELIVERABLE_ARTICLES',
      'INCOMPLETE_KITS', 'INCOMPLETE_ARTICLES',
      'WITHOUT_PROPER_DETAILS_KITS',
      'WITHOUT_PROPER_DETAILS_ARTICLES',
      'KIT_TOTAL_CHECK', 'ARTICLE_TOTAL_CHECK',
      'STATUS', 'SUBMISSION_IP', 'SESSION_ID'
    ],

    SESSION_LOG: [
      'SESSION_ID', 'LOGIN_TIME', 'LAST_SEEN', 'LOGOUT_TIME',
      'EMPLOYEE_ID', 'ROLE', 'SOL_ID', 'OFFICE_NAME',
      'IP_ADDRESS', 'USER_AGENT', 'STATUS', 'EXPIRES_AT'
    ]
  };

  Object.keys(specs).forEach(function(name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);

    sh.getRange(1, 1, 1, specs[name].length)
      .setValues([specs[name]]);

    sh.setFrozenRows(1);
  });

  seedOfficeMaster_();

  return 'Workbook setup complete. Add SPM/ADMIN users to USER_SET.';
}

function seedOfficeMaster_() {
  const offices = [
    ['18231301', 'Arnas SO'],
    ['18214601', 'Banihal SO'],
    ['18214301', 'Batote SO'],
    ['18222201', 'Bhaderwah SO'],
    ['18214801', 'Chanderkot SO'],
    ['18214101', 'Chenani SO'],
    ['18220201', 'Doda MDG'],
    ['18220601', 'DHP SO'],
    ['18212101', 'Garhi SO'],
    ['18212401', 'Jaganoo SO'],
    ['18231201', 'Jyotipuram SO'],
    ['18214701', 'Kastigarh SO'],
    ['18230101', 'Katra SO'],
    ['18220101', 'Khellani SO'],
    ['18220401', 'Kishtwar SO'],
    ['18214201', 'Kud SO'],
    ['18231501', 'Mahore SO'],
    ['18212701', 'Majalta SO'],
    ['18212801', 'Majouri SO'],
    ['18216101', 'Mantalai SO'],
    ['18212501', 'Mir SO'],
    ['18210101', 'Mukherjee Road SO'],
    ['18210102', 'Omaramorh SO'],
    ['18220501', 'Palmar SO'],
    ['18520301', 'Poni SO'],
    ['18210401', 'PTC SO'],
    ['18214401', 'Ramban SO'],
    ['18212201', 'Ramnagar SO'],
    ['18214501', 'Ramsu SO'],
    ['18231101', 'Reasi SO'],
    ['18212601', 'Roan SO'],
    ['18210103', 'Shaktinagar SO'],
    ['18232001', 'SMVDU SO'],
    ['18220301', 'Thathri SO'],
    ['18210104', 'Distt. Court SO'],
    ['18222101', 'Udrana SO'],
    ['18231102', 'DOC Reasi'],
    ['18220701', 'Gandoh SO'],
    ['18210100', 'Udhampur HO']
  ];

  const sh = getSheet_('OFFICE_MASTER');

  if (sh.getLastRow() <= 1) {
    sh.getRange(2, 1, offices.length, 3)
      .setValues(offices.map(function(row) {
        return [row[0], row[1], true];
      }));
  }
}

function getSheet_(name) {
  const sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}

function rowsAsObjects_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(String);

  return values.slice(1).map(function(row) {
    const obj = {};
    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });
    return obj;
  });
}

function normalizePhone_(value) {
  return String(value || '').replace(/\D/g, '').slice(-10);
}

function dateKey_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  const text = String(value).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text;
  }

  const parsed = new Date(text);

  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(
      parsed,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  return text;
}

function todayKey_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

function findOffice_(solId) {
  const target = String(solId || '').trim();

  const office = rowsAsObjects_(getSheet_('OFFICE_MASTER')).find(function(row) {
    return String(row.SOL_ID).trim() === target &&
      String(row.ACTIVE).toUpperCase() !== 'FALSE';
  });

  return office
    ? {
        solId: String(office.SOL_ID),
        officeName: String(office.OFFICE_NAME)
      }
    : null;
}

/* ---------------- SECURITY ---------------- */

function login(employeeId, phone, ip, userAgent) {
  employeeId = String(employeeId || '').trim();
  phone = normalizePhone_(phone);
  ip = String(ip || 'Unavailable').slice(0, 100);
  userAgent = String(userAgent || '').slice(0, 500);

  if (!employeeId || phone.length !== 10) {
    throw new Error('Enter a valid Employee ID and 10-digit phone number.');
  }

  enforceLoginRateLimit_(employeeId, ip);

  const users = rowsAsObjects_(getSheet_('USER_SET'));

  const user = users.find(function(row) {
    return String(row.EMPLOYEE_ID).trim() === employeeId &&
      normalizePhone_(row.PHONE) === phone &&
      String(row.ACTIVE).toUpperCase() !== 'FALSE';
  });

  if (!user) {
    recordLoginFailure_(employeeId, ip);
    throw new Error(
      'Invalid Employee ID / phone number, or user is inactive.'
    );
  }

  clearLoginFailures_(employeeId, ip);

  const role = String(user.ROLE || 'SPM').toUpperCase();

  if (['SPM', 'ADMIN'].indexOf(role) === -1) {
    throw new Error('Invalid user role in USER_SET.');
  }

  const office = findOffice_(String(user.SOL_ID).trim());

  if (!office) {
    throw new Error('No active office found for this user SOL ID.');
  }

  const sessionId = Utilities.getUuid();
  const now = new Date();

  const days = Math.max(
    1,
    Math.min(
      CFG.SESSION_DAYS_MAX,
      Number(user.SESSION_DAYS) || CFG.SESSION_DAYS_DEFAULT
    )
  );

  const expires = new Date(
    now.getTime() + days * 24 * 60 * 60 * 1000
  );

  getSheet_('SESSION_LOG').appendRow([
    sessionId,
    now,
    now,
    '',
    employeeId,
    role,
    office.solId,
    office.officeName,
    ip,
    userAgent,
    'ACTIVE',
    expires
  ]);

  return {
    sessionId: sessionId,
    employeeId: employeeId,
    name: String(user.NAME || ''),
    role: role,
    solId: office.solId,
    officeName: office.officeName,
    expiresAt: expires.toISOString(),
    today: todayKey_()
  };
}

function enforceLoginRateLimit_(employeeId, ip) {
  const cache = CacheService.getScriptCache();

  const key = 'login_' + Utilities.base64EncodeWebSafe(
    employeeId + '|' + ip
  ).slice(0, 80);

  const raw = cache.get(key);
  if (!raw) return;

  const data = JSON.parse(raw);

  if (
    data.count >= CFG.LOGIN_MAX_ATTEMPTS &&
    Date.now() - data.first < CFG.LOGIN_WINDOW_MS
  ) {
    throw new Error('Too many login attempts. Please try again later.');
  }

  if (Date.now() - data.first >= CFG.LOGIN_WINDOW_MS) {
    cache.remove(key);
  }
}

function recordLoginFailure_(employeeId, ip) {
  const cache = CacheService.getScriptCache();

  const key = 'login_' + Utilities.base64EncodeWebSafe(
    employeeId + '|' + ip
  ).slice(0, 80);

  const raw = cache.get(key);

  const data = raw
    ? JSON.parse(raw)
    : { count: 0, first: Date.now() };

  data.count += 1;
  cache.put(key, JSON.stringify(data), 600);
}

function clearLoginFailures_(employeeId, ip) {
  const cache = CacheService.getScriptCache();

  const key = 'login_' + Utilities.base64EncodeWebSafe(
    employeeId + '|' + ip
  ).slice(0, 80);

  cache.remove(key);
}

function validateSession_(sessionId, requiredRole) {
  if (!sessionId) {
    throw new Error('Session expired. Please login again.');
  }

  const sh = getSheet_('SESSION_LOG');
  const rows = rowsAsObjects_(sh);

  const session = rows.find(function(row) {
    return String(row.SESSION_ID) === String(sessionId);
  });

  if (!session) throw new Error('Invalid session.');

  const expires = new Date(session.EXPIRES_AT);

  if (
    String(session.STATUS).toUpperCase() !== 'ACTIVE' ||
    isNaN(expires.getTime()) ||
    expires.getTime() < Date.now()
  ) {
    throw new Error('Session expired. Please login again.');
  }

  if (
    requiredRole &&
    String(session.ROLE).toUpperCase() !==
    String(requiredRole).toUpperCase()
  ) {
    throw new Error('Unauthorized.');
  }

  const data = sh.getDataRange().getValues();
  if (data.length > 1) {
    const headers = data[0].map(String);
    const sidIndex = headers.indexOf('SESSION_ID');
    const lastSeenIndex = headers.indexOf('LAST_SEEN');

    if (sidIndex >= 0 && lastSeenIndex >= 0) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][sidIndex]) === String(sessionId)) {
          sh.getRange(i + 1, lastSeenIndex + 1).setValue(new Date());
          break;
        }
      }
    }
  }

  return {
    sessionId: String(session.SESSION_ID),
    employeeId: String(session.EMPLOYEE_ID),
    role: String(session.ROLE).toUpperCase(),
    solId: String(session.SOL_ID),
    officeName: String(session.OFFICE_NAME),
    ip: String(session.IP_ADDRESS || ''),
    expiresAt: expires.toISOString()
  };
}

/* ---------------- COMMON ---------------- */

function logout(sessionId) {
  try {
    const sh = getSheet_('SESSION_LOG');
    const data = sh.getDataRange().getValues();

    if (data.length < 2) return true;

    const headers = data[0].map(String);
    const sid = headers.indexOf('SESSION_ID');
    const status = headers.indexOf('STATUS');
    const logoutTime = headers.indexOf('LOGOUT_TIME');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][sid]) === String(sessionId)) {
        sh.getRange(i + 1, status + 1).setValue('LOGGED_OUT');
        sh.getRange(i + 1, logoutTime + 1).setValue(new Date());
        break;
      }
    }
  } catch (_) {}

  return true;
}

function number_(value, label) {
  if (value === '' || value === null || typeof value === 'undefined') {
    throw new Error(label + ' is required.');
  }

  const n = Number(value);

  if (!Number.isInteger(n) || n < 0 || n > CFG.MAX_VALUE) {
    throw new Error(
      label + ' must be a whole number from 0 to ' +
      CFG.MAX_VALUE + '.'
    );
  }

  return n;
}

function getTodayStatus(sessionId) {
  const session = validateSession_(sessionId, 'SPM');
  const today = todayKey_();

  const record = rowsAsObjects_(getSheet_('DAILY_RECORD')).find(function(row) {
    return dateKey_(row.REPORT_DATE) === today &&
      String(row.EMPLOYEE_ID) === session.employeeId &&
      String(row.SOL_ID) === session.solId;
  });

  return record
    ? { submitted: true, record: sanitizeRecord_(record) }
    : { submitted: false };
}

function sanitizeRecord_(row) {
  return {
    reportDate: dateKey_(row.REPORT_DATE),
    employeeId: String(row.EMPLOYEE_ID || ''),
    solId: String(row.SOL_ID || ''),
    officeName: String(row.OFFICE_NAME || ''),
    allKits: Number(row.ALL_KITS || 0),
    similarArticle: Number(row.SIMILAR_ARTICLE || 0),
    invalidMobileKits: Number(row.INVALID_MOBILE_KITS || 0),
    invalidMobileArticles: Number(row.INVALID_MOBILE_ARTICLES || 0),
    deliverableKits: Number(row.DELIVERABLE_KITS || 0),
    deliverableArticles: Number(row.DELIVERABLE_ARTICLES || 0),
    incompleteKits: Number(row.INCOMPLETE_KITS || 0),
    incompleteArticles: Number(row.INCOMPLETE_ARTICLES || 0),
    withoutProperDetailsKits:
      Number(row.WITHOUT_PROPER_DETAILS_KITS || 0),
    withoutProperDetailsArticles:
      Number(row.WITHOUT_PROPER_DETAILS_ARTICLES || 0),
    kitTotalCheck: String(row.KIT_TOTAL_CHECK || ''),
    articleTotalCheck: String(row.ARTICLE_TOTAL_CHECK || ''),
    status: String(row.STATUS || '')
  };
}

/* ---------------- SPM ---------------- */

function getMyRecords(sessionId, limit) {
  const session = validateSession_(sessionId, 'SPM');

  const n = Math.max(
    1,
    Math.min(Number(limit) || 30, CFG.MAX_RECORDS)
  );

  return rowsAsObjects_(getSheet_('DAILY_RECORD'))
    .filter(function(row) {
      return String(row.EMPLOYEE_ID) === session.employeeId &&
        String(row.SOL_ID) === session.solId;
    })
    .slice(-n)
    .reverse()
    .map(sanitizeRecord_);
}

function submitDaily(sessionId, payload) {
  const session = validateSession_(sessionId, 'SPM');

  if (!payload || typeof payload !== 'object') {
    throw new Error('No report data received.');
  }

  const n = {
    allKits: number_(payload.allKits, 'All Kits'),
    similarArticle: number_(
      payload.similarArticle,
      'Similar Article / Tool Kit Article'
    ),
    invalidMobileKits: number_(
      payload.invalidMobileKits,
      'Invalid Mobile Kits'
    ),
    invalidMobileArticles: number_(
      payload.invalidMobileArticles,
      'Invalid Mobile Articles'
    ),
    deliverableKits: number_(
      payload.deliverableKits,
      'Deliverable Kits'
    ),
    deliverableArticles: number_(
      payload.deliverableArticles,
      'Deliverable Articles'
    ),
    incompleteKits: number_(
      payload.incompleteKits,
      'Incomplete Kits'
    ),
    incompleteArticles: number_(
      payload.incompleteArticles,
      'Incomplete Articles'
    ),
    withoutProperDetailsKits: number_(
      payload.withoutProperDetailsKits,
      'Without Proper Details Kits'
    ),
    withoutProperDetailsArticles: number_(
      payload.withoutProperDetailsArticles,
      'Without Proper Details Articles'
    )
  };

  const expectedKits =
    n.invalidMobileKits +
    n.deliverableKits +
    n.incompleteKits +
    n.withoutProperDetailsKits;

  const expectedArticles =
    n.invalidMobileArticles +
    n.deliverableArticles +
    n.incompleteArticles +
    n.withoutProperDetailsArticles;

  if (n.allKits !== expectedKits) {
    throw new Error(
      'Kit logic failed: All Kits must equal Invalid Mobile + ' +
      'Deliverable + Incomplete + Without Proper Details/Address.'
    );
  }

  if (n.similarArticle !== expectedArticles) {
    throw new Error(
      'Article logic failed: Similar Article / Tool Kit Article must ' +
      'equal Invalid Mobile + Deliverable + Incomplete + ' +
      'Without Proper Details/Address.'
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sh = getSheet_('DAILY_RECORD');
    const today = todayKey_();

    const duplicate = rowsAsObjects_(sh).some(function(row) {
      return dateKey_(row.REPORT_DATE) === today &&
        String(row.EMPLOYEE_ID) === session.employeeId &&
        String(row.SOL_ID) === session.solId;
    });

    if (duplicate) {
      throw new Error(
        "Today's report is already submitted for this office. " +
        'Duplicate submission blocked by server.'
      );
    }

    const now = new Date();

    sh.appendRow([
      now,
      today,
      session.employeeId,
      session.employeeId.slice(-4),
      session.solId,
      session.officeName,
      n.allKits,
      n.similarArticle,
      n.invalidMobileKits,
      n.invalidMobileArticles,
      n.deliverableKits,
      n.deliverableArticles,
      n.incompleteKits,
      n.incompleteArticles,
      n.withoutProperDetailsKits,
      n.withoutProperDetailsArticles,
      true,
      true,
      'SUBMITTED',
      String(payload.ip || session.ip || 'Unavailable').slice(0, 100),
      session.sessionId
    ]);

    return {
      ok: true,
      message: 'Daily report submitted successfully.',
      record: sanitizeRecord_({
        REPORT_DATE: today,
        EMPLOYEE_ID: session.employeeId,
        SOL_ID: session.solId,
        OFFICE_NAME: session.officeName,
        ALL_KITS: n.allKits,
        SIMILAR_ARTICLE: n.similarArticle,
        INVALID_MOBILE_KITS: n.invalidMobileKits,
        INVALID_MOBILE_ARTICLES: n.invalidMobileArticles,
        DELIVERABLE_KITS: n.deliverableKits,
        DELIVERABLE_ARTICLES: n.deliverableArticles,
        INCOMPLETE_KITS: n.incompleteKits,
        INCOMPLETE_ARTICLES: n.incompleteArticles,
        WITHOUT_PROPER_DETAILS_KITS:
          n.withoutProperDetailsKits,
        WITHOUT_PROPER_DETAILS_ARTICLES:
          n.withoutProperDetailsArticles,
        KIT_TOTAL_CHECK: true,
        ARTICLE_TOTAL_CHECK: true,
        STATUS: 'SUBMITTED'
      })
    };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- ADMIN ---------------- */

function adminSummary(sessionId, filters) {
  validateSession_(sessionId, 'ADMIN');

  filters = filters || {};

  const records = rowsAsObjects_(getSheet_('DAILY_RECORD'));
  const offices = rowsAsObjects_(getSheet_('OFFICE_MASTER'));

  const from = String(filters.from || '');
  const to = String(filters.to || '');
  const solId = String(filters.solId || '');
  const date = String(filters.date || '');

  if (from && to && from > to) {
    throw new Error('From date cannot be after To date.');
  }

  const rows = records
    .filter(function(row) {
      const d = dateKey_(row.REPORT_DATE);

      if (from && d < from) return false;
      if (to && d > to) return false;
      if (date && d !== date) return false;
      if (solId && String(row.SOL_ID) !== solId) return false;

      return true;
    })
    .map(sanitizeRecord_);

  const totals = rows.reduce(function(total, row) {
    total.allKits += row.allKits;
    total.similarArticle += row.similarArticle;
    total.invalidMobileKits += row.invalidMobileKits;
    total.invalidMobileArticles += row.invalidMobileArticles;
    total.deliverableKits += row.deliverableKits;
    total.deliverableArticles += row.deliverableArticles;
    total.incompleteKits += row.incompleteKits;
    total.incompleteArticles += row.incompleteArticles;
    total.withoutProperDetailsKits += row.withoutProperDetailsKits;
    total.withoutProperDetailsArticles += row.withoutProperDetailsArticles;
    return total;
  }, {
    allKits: 0,
    similarArticle: 0,
    invalidMobileKits: 0,
    invalidMobileArticles: 0,
    deliverableKits: 0,
    deliverableArticles: 0,
    incompleteKits: 0,
    incompleteArticles: 0,
    withoutProperDetailsKits: 0,
    withoutProperDetailsArticles: 0
  });

  /*
   * For a specific date, pending means offices without a report on that date.
   * For a date range, pending means offices with no report anywhere in the range.
   */
  const submitted = {};

  rows.forEach(function(row) {
    submitted[row.solId] = true;
  });

  const activeOffices = offices.filter(function(office) {
    return String(office.ACTIVE).toUpperCase() !== 'FALSE';
  });

  const pendingOffices = activeOffices
    .filter(function(office) {
      return !submitted[String(office.SOL_ID)];
    })
    .map(function(office) {
      return {
        solId: String(office.SOL_ID),
        officeName: String(office.OFFICE_NAME)
      };
    });

  return {
    rows: rows,
    totals: totals,
    pendingOffices: pendingOffices,
    pendingOfficeCount: pendingOffices.length,
    totalOffices: activeOffices.length,
    submittedOfficeCount: Object.keys(submitted).length
  };
}

function adminSessions(sessionId, limit) {
  validateSession_(sessionId, 'ADMIN');

  const n = Math.max(
    1,
    Math.min(Number(limit) || 200, CFG.MAX_SESSIONS)
  );

  return rowsAsObjects_(getSheet_('SESSION_LOG'))
    .slice(-n)
    .reverse()
    .map(function(row) {
      return {
        sessionId: String(row.SESSION_ID || ''),
        loginTime: String(row.LOGIN_TIME || ''),
        lastSeen: String(row.LAST_SEEN || ''),
        logoutTime: String(row.LOGOUT_TIME || ''),
        employeeId: String(row.EMPLOYEE_ID || ''),
        role: String(row.ROLE || ''),
        solId: String(row.SOL_ID || ''),
        officeName: String(row.OFFICE_NAME || ''),
        ip: String(row.IP_ADDRESS || ''),
        userAgent: String(row.USER_AGENT || ''),
        status: String(row.STATUS || ''),
        expiresAt: String(row.EXPIRES_AT || '')
      };
    });
}

function getOffices(sessionId) {
  validateSession_(sessionId, 'ADMIN');

  return rowsAsObjects_(getSheet_('OFFICE_MASTER'))
    .filter(function(row) {
      return String(row.ACTIVE).toUpperCase() !== 'FALSE';
    })
    .map(function(row) {
      return {
        solId: String(row.SOL_ID),
        officeName: String(row.OFFICE_NAME)
      };
    });
}

function exportDailyCsv(sessionId, filters) {
  const result = adminSummary(sessionId, filters || {});

  const headers = [
    'reportDate',
    'employeeId',
    'solId',
    'officeName',
    'allKits',
    'similarArticle',
    'invalidMobileKits',
    'invalidMobileArticles',
    'deliverableKits',
    'deliverableArticles',
    'incompleteKits',
    'incompleteArticles',
    'withoutProperDetailsKits',
    'withoutProperDetailsArticles',
    'status'
  ];

  return [
    headers.join(','),
    ...result.rows.map(function(row) {
      return headers.map(function(key) {
        return csv_(row[key]);
      }).join(',');
    })
  ].join('\n');
}

function csv_(value) {
  return '"' +
    String(value == null ? '' : value).replace(/"/g, '""') +
    '"';
}
