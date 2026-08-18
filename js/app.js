'use strict';

const CFG = window.PMV_CONFIG || {};
const $ = id => document.getElementById(id);
let state = { session: null };
const pending = new Map();

const TODAY = () => new Intl.DateTimeFormat('en-CA', {
  timeZone: 'Asia/Kolkata',
  year: 'numeric',
  month: '2-digit',
  day: '2-digit'
}).format(new Date());

function show(id, on = true) {
  const el = $(id);
  if (el) el.classList.toggle('hidden', !on);
}

function alertBox(id, text, kind) {
  const el = $(id);
  if (!el) return;
  el.className = 'alert ' + kind;
  el.textContent = text;
  show(id, true);
}

function configuredUrl() {
  return String(CFG.APP_SCRIPT_API_URL || '').trim();
}

function validBackend(url) {
  return /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:\?.*)?$/.test(url);
}

function makeRequestId() {
  try {
    if (crypto && crypto.randomUUID) return crypto.randomUUID();
  } catch (_) {}
  return Date.now() + '-' + Math.random().toString(36).slice(2);
}

/*
 * Apps Script HtmlService responses are sandboxed in an iframe.
 * We intentionally use a form -> hidden iframe transport instead of
 * fetch(), because cross-origin Apps Script redirects can otherwise
 * cause CORS problems from GitHub Pages.
 */
function request(action, payload = {}) {
  return new Promise((resolve, reject) => {
    const url = configuredUrl();

    if (!validBackend(url)) {
      reject(new Error('Apps Script /exec URL is missing or invalid.'));
      return;
    }

    const requestId = makeRequestId();
    const timeout = Number(CFG.REQUEST_TIMEOUT_MS) || 60000;

    const timer = setTimeout(() => {
      if (!pending.has(requestId)) return;
      pending.delete(requestId);
      reject(new Error(
        'Server response timed out. Please verify the Apps Script deployment, ' +
        'SPREADSHEET_ID and Web App access settings.'
      ));
    }, timeout);

    pending.set(requestId, { resolve, reject, timer });

    const form = document.createElement('form');
    form.method = 'POST';
    form.action = url;
    form.target = 'bridge';
    form.style.display = 'none';

    const fields = {
      action,
      requestId,
      payload: JSON.stringify(payload || {})
    };

    Object.keys(fields).forEach(name => {
      const input = document.createElement('input');
      input.type = 'hidden';
      input.name = name;
      input.value = fields[name];
      form.appendChild(input);
    });

    document.body.appendChild(form);

    try {
      form.submit();
    } catch (err) {
      clearTimeout(timer);
      pending.delete(requestId);
      reject(err);
    } finally {
      form.remove();
    }
  });
}

window.addEventListener('message', event => {
  const data = event.data;
  if (!data || !data.requestId) return;

  const item = pending.get(data.requestId);
  if (!item) return;

  pending.delete(data.requestId);
  clearTimeout(item.timer);

  if (data.ok) {
    item.resolve(data.data);
  } else {
    item.reject(new Error(
      data.data && data.data.message
        ? data.data.message
        : 'Server request failed.'
    ));
  }
});

function esc(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, m => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[m]));
}

function sessionKey() {
  return CFG.SESSION_STORAGE_KEY || 'pmvSessionV8';
}

function sessionSave(session) {
  state.session = session;
  sessionStorage.setItem(sessionKey(), JSON.stringify(session));
}

function sessionLoad() {
  try {
    const value = sessionStorage.getItem(sessionKey());
    return value ? JSON.parse(value) : null;
  } catch (_) {
    return null;
  }
}

function sessionClear() {
  state.session = null;
  sessionStorage.removeItem(sessionKey());
}

function fieldIds() {
  return [
    'allKits',
    'similarArticle',
    'invalidMobileKits',
    'invalidMobileArticles',
    'deliverableKits',
    'deliverableArticles',
    'incompleteKits',
    'incompleteArticles',
    'withoutProperDetailsKits',
    'withoutProperDetailsArticles'
  ];
}

async function doLogin() {
  show('loginErr', false);

  const employeeId = $('employeeId').value.trim();
  const phone = $('phone').value.replace(/\D/g, '');

  if (!employeeId || phone.length !== 10) {
    alertBox(
      'loginErr',
      'Enter Employee ID and a valid 10-digit registered phone number.',
      'error'
    );
    return;
  }

  $('loginBtn').disabled = true;

  try {
    const session = await request('login', {
      employeeId,
      phone,
      ip: 'Unavailable',
      userAgent: navigator.userAgent.slice(0, 300)
    });

    sessionSave(session);
    renderApp();
  } catch (err) {
    alertBox('loginErr', err.message || 'Login failed.', 'error');
  } finally {
    $('loginBtn').disabled = false;
  }
}

function renderApp() {
  const session = state.session;
  if (!session) return;

  if (!session.expiresAt || new Date(session.expiresAt) <= new Date()) {
    sessionClear();
    location.reload();
    return;
  }

  show('loginScreen', false);
  show('app', true);

  $('who').textContent =
    `${session.name || session.employeeId} · ${session.officeName} · ` +
    `SOL ${session.solId} · Session valid until ` +
    `${new Date(session.expiresAt).toLocaleString('en-IN')}`;

  const isAdmin = String(session.role).toUpperCase() === 'ADMIN';

  show('adminView', isAdmin);
  show('spmView', !isAdmin);

  if (isAdmin) {
    setAdminDates();
    loadOffices();
    loadDashboard();
  } else {
    $('solId').value = session.solId;
    $('officeName').value = session.officeName;
    checkToday();
    loadMyRecords();
  }
}

async function checkToday() {
  try {
    const result = await request('getTodayStatus', {
      sessionId: state.session.sessionId
    });

    if (result.submitted) {
      fillRecord(result.record);
      alertBox(
        'spmMsg',
        'Today’s report is already submitted. Duplicate submission is blocked by the server.',
        'success'
      );
    } else {
      clearEntry();
      alertBox('spmMsg', 'No report has been submitted for today.', 'info');
    }
  } catch (err) {
    alertBox('spmMsg', err.message, 'error');
  }
}

function fillRecord(record) {
  fieldIds().forEach(id => {
    $(id).value = record[id] ?? '';
    $(id).disabled = true;
  });
  $('submitBtn').disabled = true;
}

function clearEntry() {
  fieldIds().forEach(id => {
    $(id).value = '';
    $(id).disabled = false;
  });
  $('submitBtn').disabled = false;
}

async function submitReport() {
  if (!state.session) return;

  const report = {};
  let valid = true;

  fieldIds().forEach(id => {
    const value = Number($(id).value);
    if (!Number.isInteger(value) || value < 0 || value > 1000000) {
      valid = false;
    }
    report[id] = value;
  });

  if (!valid) {
    alertBox(
      'spmMsg',
      'All fields must contain whole numbers from 0 to 1,000,000.',
      'error'
    );
    return;
  }

  if (
    report.allKits !==
    report.invalidMobileKits +
    report.deliverableKits +
    report.incompleteKits +
    report.withoutProperDetailsKits
  ) {
    alertBox('spmMsg', 'Kit total does not match the four kit categories.', 'error');
    return;
  }

  if (
    report.similarArticle !==
    report.invalidMobileArticles +
    report.deliverableArticles +
    report.incompleteArticles +
    report.withoutProperDetailsArticles
  ) {
    alertBox('spmMsg', 'Article total does not match the four article categories.', 'error');
    return;
  }

  $('submitBtn').disabled = true;

  try {
    const result = await request('submitDaily', {
      sessionId: state.session.sessionId,
      report
    });

    alertBox(
      'spmMsg',
      result.message || 'Report submitted successfully.',
      'success'
    );

    fillRecord(result.record || report);
    loadMyRecords();
  } catch (err) {
    alertBox('spmMsg', err.message, 'error');
    $('submitBtn').disabled = false;
  }
}

async function loadMyRecords() {
  try {
    const rows = await request('getMyRecords', {
      sessionId: state.session.sessionId,
      limit: 30
    });

    $('myTable').innerHTML =
      '<thead><tr>' +
      '<th>Date</th><th>All Kits</th><th>Deliverable</th>' +
      '<th>Invalid Mobile</th><th>Incomplete</th>' +
      '<th>Without Details</th><th>Status</th>' +
      '</tr></thead><tbody>' +
      rows.map(row =>
        `<tr>
          <td>${esc(row.reportDate)}</td>
          <td>${Number(row.allKits || 0).toLocaleString('en-IN')}</td>
          <td>${Number(row.deliverableKits || 0).toLocaleString('en-IN')}</td>
          <td>${Number(row.invalidMobileKits || 0).toLocaleString('en-IN')}</td>
          <td>${Number(row.incompleteKits || 0).toLocaleString('en-IN')}</td>
          <td>${Number(row.withoutProperDetailsKits || 0).toLocaleString('en-IN')}</td>
          <td><span class="badge good">${esc(row.status)}</span></td>
        </tr>`
      ).join('') +
      '</tbody>';
  } catch (err) {
    alertBox('spmMsg', err.message, 'error');
  }
}

function setAdminDates() {
  const today = TODAY();
  $('specificDate').value = today;
  $('fromDate').value = today;
  $('toDate').value = today;
}

function filters() {
  return {
    from: $('fromDate').value,
    to: $('toDate').value,
    date: $('specificDate').value,
    solId: $('filterSol').value
  };
}

async function loadOffices() {
  try {
    const offices = await request('getOffices', {
      sessionId: state.session.sessionId
    });

    $('filterSol').innerHTML =
      '<option value="">All offices</option>' +
      offices.map(o =>
        `<option value="${esc(o.solId)}">${esc(o.officeName)} (${esc(o.solId)})</option>`
      ).join('');
  } catch (err) {
    alertBox('adminMsg', err.message, 'error');
  }
}

async function loadDashboard() {
  const filter = filters();

  if (filter.from && filter.to && filter.from > filter.to) {
    alertBox('adminMsg', 'From date cannot be after To date.', 'error');
    return;
  }

  try {
    alertBox('adminMsg', 'Loading consolidated data…', 'info');

    const [summary, offices] = await Promise.all([
      request('adminSummary', {
        sessionId: state.session.sessionId,
        filters: filter
      }),
      request('getOffices', {
        sessionId: state.session.sessionId
      })
    ]);

    renderAdmin(summary, offices);
    alertBox('adminMsg', 'Dashboard updated successfully.', 'success');
  } catch (err) {
    alertBox('adminMsg', err.message, 'error');
  }
}

function renderAdmin(summary, offices) {
  $('kTotal').textContent = summary.totalOffices;
  $('kSubmitted').textContent = summary.submittedOfficeCount;
  $('kPending').textContent = summary.pendingOfficeCount;

  const pct = summary.totalOffices
    ? Math.round(summary.submittedOfficeCount / summary.totalOffices * 100)
    : 0;

  $('kCompletion').textContent = pct + '%';
  $('kProgress').style.width = pct + '%';

  const t = summary.totals;

  const cards = [
    ['All Kits', t.allKits],
    ['Similar Articles', t.similarArticle],
    ['Deliverable Kits', t.deliverable],
    ['Incomplete Kits', t.incomplete],
    ['Invalid Mobile Kits', t.invalidMobileKits],
    ['Without Details Kits', t.withoutProperDetailsKits],
    ['Deliverable Articles', t.deliverableArticles],
    ['Incomplete Articles', t.incompleteArticles]
  ];

  $('stats').innerHTML = cards.map(card =>
    `<div class="kpi">
      <small>${esc(card[0]).toUpperCase()}</small>
      <strong>${Number(card[1] || 0).toLocaleString('en-IN')}</strong>
      <div class="accent"></div>
    </div>`
  ).join('');

  const bySol = {};
  summary.rows.forEach(row => {
    bySol[row.solId] = row;
  });

  $('reportTable').innerHTML =
    '<thead><tr>' +
    '<th>Office</th><th>SOL ID</th><th>Report</th>' +
    '<th>All Kits</th><th>Deliverable</th><th>Invalid</th>' +
    '<th>Incomplete</th><th>Without Details</th><th>Articles</th>' +
    '</tr></thead><tbody>' +
    offices.map(office => {
      const row = bySol[office.solId];
      return `<tr>
        <td>${esc(office.officeName)}</td>
        <td>${esc(office.solId)}</td>
        <td>${row
          ? '<span class="badge good">SUBMITTED</span>'
          : '<span class="badge bad">PENDING</span>'}</td>
        <td>${row ? row.allKits : 0}</td>
        <td>${row ? row.deliverableKits : 0}</td>
        <td>${row ? row.invalidMobileKits : 0}</td>
        <td>${row ? row.incompleteKits : 0}</td>
        <td>${row ? row.withoutProperDetailsKits : 0}</td>
        <td>${row ? row.similarArticle : 0}</td>
      </tr>`;
    }).join('') +
    '</tbody>';

  const bars = [
    ['Deliverable', t.deliverableKits],
    ['Invalid Mobile', t.invalidMobileKits],
    ['Incomplete', t.incompleteKits],
    ['Without Details', t.withoutProperDetailsKits]
  ];

  const max = Math.max(1, ...bars.map(x => Number(x[1] || 0)));

  $('bars').innerHTML = bars.map(item =>
    `<div class="bar-row">
      <b>${esc(item[0])}</b>
      <div class="bar"><span style="width:${Math.round(Number(item[1] || 0) / max * 100)}%"></span></div>
      <strong>${Number(item[1] || 0).toLocaleString('en-IN')}</strong>
    </div>`
  ).join('');

  $('pendingTable').innerHTML =
    '<thead><tr><th>#</th><th>Office</th><th>SOL ID</th><th>Action</th></tr></thead>' +
    '<tbody>' +
    summary.pendingOffices.map((office, index) =>
      `<tr>
        <td>${index + 1}</td>
        <td>${esc(office.officeName)}</td>
        <td>${esc(office.solId)}</td>
        <td><span class="badge bad">UPDATE REQUIRED</span></td>
      </tr>`
    ).join('') +
    '</tbody>';
}

async function loadSessions() {
  show('sessionPanel', true);

  try {
    const rows = await request('adminSessions', {
      sessionId: state.session.sessionId,
      limit: 100
    });

    $('sessionTable').innerHTML =
      '<thead><tr>' +
      '<th>Login</th><th>Employee ID</th><th>Role</th>' +
      '<th>Office</th><th>Status</th><th>Expires</th>' +
      '</tr></thead><tbody>' +
      rows.map(row =>
        `<tr>
          <td>${esc(row.loginTime)}</td>
          <td>${esc(row.employeeId)}</td>
          <td>${esc(row.role)}</td>
          <td>${esc(row.officeName)}</td>
          <td>${esc(row.status)}</td>
          <td>${esc(row.expiresAt)}</td>
        </tr>`
      ).join('') +
      '</tbody>';
  } catch (err) {
    alertBox('adminMsg', err.message, 'error');
  }
}

async function downloadCsv() {
  try {
    const csv = await request('exportDailyCsv', {
      sessionId: state.session.sessionId,
      filters: filters()
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');

    a.href = url;
    a.download = 'PMV_Daily_Report_' + TODAY() + '.csv';
    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (err) {
    alertBox('adminMsg', err.message, 'error');
  }
}

async function doLogout() {
  try {
    if (state.session) {
      await request('logout', { sessionId: state.session.sessionId });
    }
  } catch (_) {}

  sessionClear();
  location.reload();
}

function refreshCurrent() {
  if (!state.session) return;

  if (String(state.session.role).toUpperCase() === 'ADMIN') {
    loadDashboard();
  } else {
    checkToday();
    loadMyRecords();
  }
}

async function checkServerHealth() {
  try {
    const result = await request('health', {});

    if (result && result.database && result.database.ok) {
      $('health').textContent =
        '● Server + Spreadsheet connected';
      $('health').className = 'health online';
      return;
    }

    const message =
      result &&
      result.database &&
      result.database.message
        ? result.database.message
        : 'Spreadsheet connection failed.';

    $('health').textContent =
      '● Database unavailable';
    $('health').className = 'health offline';

    alertBox('loginErr', message, 'error');

  } catch (err) {
    $('health').textContent =
      '● Backend unavailable';
    $('health').className =
      'health offline';

    alertBox(
      'loginErr',
      err.message ||
        'Apps Script Web App is not responding.',
      'error'
    );
  }
}

(function boot() {
  const session = sessionLoad();

  if (session) {
    state.session = session;
    renderApp();
  } else {
    show('loginScreen', true);
    checkServerHealth();
  }
})();
