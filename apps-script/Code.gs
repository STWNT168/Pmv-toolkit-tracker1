/**
 * SPM Tool Kits Reporting v3
 * GitHub Pages frontend -> Apps Script Web App -> Google Sheets
 *
 * v3 additions:
 * - Separate SPM dashboard with today's status + previous submissions
 * - Full kit and article fields
 * - Server-side kit and article reconciliation
 * - Admin dashboard with both kit/article totals and office completion
 * - 7-day sessions, IP/user-agent audit and server duplicate protection
 */

const CFG = {
  SESSION_DAYS: 7,
  MAX_VALUE: 1000000,
  KIT_FIELDS: ['allKits','invalidMobileKits','deliverableKits','incompleteKits','withoutProperDetailsKits'],
  ARTICLE_FIELDS: ['similarArticle','invalidMobileArticles','deliverableArticles','incompleteArticles','withoutProperDetailsArticles']
};

function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Backend is not configured. Set Script Property SPREADSHEET_ID first.');
  return SpreadsheetApp.openById(id);
}

function doGet() {
  return HtmlService.createHtmlOutput('<h3>SPM Tool Kits API v3</h3><p>Backend is running. Use the GitHub Pages frontend.</p>')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const requestId = String((e && e.parameter && e.parameter.requestId) || Utilities.getUuid());
  try {
    const action = String((e && e.parameter && e.parameter.action) || '');
    const payload = e && e.parameter && e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    let result;
    switch (action) {
      case 'login': result = login(payload.employeeId, payload.phone, payload.ip, payload.userAgent); break;
      case 'getTodayStatus': result = getTodayStatus(payload.sessionId); break;
      case 'getMyRecords': result = getMyRecords(payload.sessionId, payload.limit || 30); break;
      case 'submitDaily': result = submitDaily(payload.sessionId, payload.report); break;
      case 'adminSummary': result = adminSummary(payload.sessionId, payload.filters || {}); break;
      case 'adminSessions': result = adminSessions(payload.sessionId, payload.limit || 200); break;
      case 'getOffices': result = getOffices(payload.sessionId); break;
      case 'exportDailyCsv': result = exportDailyCsv(payload.sessionId, payload.filters || {}); break;
      case 'logout': result = logout(payload.sessionId); break;
      default: throw new Error('Unknown API action.');
    }
    return bridgeResponse_(requestId, true, result);
  } catch (err) {
    return bridgeResponse_(requestId, false, {message: String(err && err.message || err)});
  }
}

function bridgeResponse_(requestId, ok, data) {
  const json = JSON.stringify({requestId, ok, data})
    .replace(/\\/g,'\\\\').replace(/</g,'\\u003c').replace(/>/g,'\\u003e').replace(/&/g,'\\u0026');
  return HtmlService.createHtmlOutput('<!doctype html><html><body><script>window.parent.postMessage(' + json + ', "*");</script></body></html>')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupWorkbook() {
  const ss = getSS_();
  const specs = {
    OFFICE_MASTER: ['SOL_ID','OFFICE_NAME','ACTIVE'],
    USER_SET: ['EMPLOYEE_ID','PHONE','ROLE','SOL_ID','ACTIVE','NAME','SESSION_DAYS'],
    DAILY_RECORD: [
      'TIMESTAMP','REPORT_DATE','EMPLOYEE_ID','PHONE_LAST4','SOL_ID','OFFICE_NAME',
      'ALL_KITS','SIMILAR_ARTICLE',
      'INVALID_MOBILE_KITS','INVALID_MOBILE_ARTICLES',
      'DELIVERABLE_KITS','DELIVERABLE_ARTICLES',
      'INCOMPLETE_KITS','INCOMPLETE_ARTICLES',
      'WITHOUT_PROPER_DETAILS_KITS','WITHOUT_PROPER_DETAILS_ARTICLES',
      'KIT_TOTAL_CHECK','ARTICLE_TOTAL_CHECK','STATUS','SUBMISSION_IP','SESSION_ID'
    ],
    SESSION_LOG: ['SESSION_ID','LOGIN_TIME','LAST_SEEN','LOGOUT_TIME','EMPLOYEE_ID','ROLE','SOL_ID','OFFICE_NAME','IP_ADDRESS','USER_AGENT','STATUS','EXPIRES_AT']
  };
  Object.keys(specs).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1,1,1,specs[name].length).setValues([specs[name]]);
    sh.setFrozenRows(1);
  });
  seedOfficeMaster_();
  return 'Workbook setup complete. Add SPM and ADMIN credentials to USER_SET.';
}

function seedOfficeMaster_() {
  const offices = [
    ['18231301','Arnas SO'],['18214601','Banihal SO'],['18214301','Batote SO'],['18222201','Bhaderwah SO'],
    ['18214801','Chanderkot SO'],['18214101','Chenani SO'],['18220201','Doda MDG'],['18220601','DHP SO'],
    ['18212101','Garhi SO'],['18212401','Jaganoo SO'],['18231201','Jyotipuram SO'],['18214701','Kastigarh SO'],
    ['18230101','Katra SO'],['18220101','Khellani SO'],['18220401','Kishtwar SO'],['18214201','Kud SO'],
    ['18231501','Mahore SO'],['18212701','Majalta SO'],['18212801','Majouri SO'],['18216101','Mantalai SO'],
    ['18212501','Mir SO'],['18210101','Mukherjee Road SO'],['18210102','Omaramorh SO'],['18220501','Palmar SO'],
    ['18520301','Poni SO'],['18210401','PTC SO'],['18214401','Ramban SO'],['18212201','Ramnagar SO'],
    ['18214501','Ramsu SO'],['18231101','Reasi SO'],['18212601','Roan SO'],['18210103','Shaktinagar SO'],
    ['18232001','SMVDU SO'],['18220301','Thathri SO'],['18210104','Distt. Court SO'],['18222101','Udrana SO'],
    ['18231102','DOC Reasi'],['18220701','Gandoh SO'],['18210100','Udhampur HO']
  ];
  const sh = getSheet_('OFFICE_MASTER');
  if (sh.getLastRow() <= 1) sh.getRange(2,1,offices.length,3).setValues(offices.map(r=>[r[0],r[1],true]));
}

function getSheet_(name) {
  const sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}
function rowsAsObjects_(sh) {
  const v=sh.getDataRange().getValues(); if(v.length<2)return[];
  const h=v[0].map(String);
  return v.slice(1).map(r=>{const o={};h.forEach((x,i)=>o[x]=r[i]);return o;});
}
function normalizePhone_(p){return String(p||'').replace(/\D/g,'').slice(-10);}
function todayKey_(){return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');}

function findOffice_(solId) {
  const o=rowsAsObjects_(getSheet_('OFFICE_MASTER')).find(x=>String(x.SOL_ID).trim()===String(solId).trim() && String(x.ACTIVE).toUpperCase()!=='FALSE');
  return o ? {solId:String(o.SOL_ID),officeName:String(o.OFFICE_NAME)} : null;
}

function login(employeeId, phone, ip, userAgent) {
  employeeId=String(employeeId||'').trim(); phone=normalizePhone_(phone);
  if(!employeeId || phone.length!==10) throw new Error('Enter valid Employee ID and 10-digit phone number.');
  const user=rowsAsObjects_(getSheet_('USER_SET')).find(u=>String(u.EMPLOYEE_ID).trim()===employeeId && normalizePhone_(u.PHONE)===phone && String(u.ACTIVE).toUpperCase()!=='FALSE');
  if(!user) throw new Error('Invalid Employee ID / phone number, or user is inactive.');
  const role=String(user.ROLE||'SPM').toUpperCase();
  if (!['SPM','ADMIN'].includes(role)) throw new Error('Invalid user role in USER_SET.');
  const office=findOffice_(String(user.SOL_ID).trim()); if(!office) throw new Error('No active office found for this user SOL ID.');
  const sessionId=Utilities.getUuid(), now=new Date();
  const days=Math.max(1,Math.min(7,Number(user.SESSION_DAYS)||CFG.SESSION_DAYS));
  const expires=new Date(now.getTime()+days*24*60*60*1000);
  getSheet_('SESSION_LOG').appendRow([sessionId,now,now,'',employeeId,role,office.solId,office.officeName,String(ip||'Unavailable'),String(userAgent||''),'ACTIVE',expires]);
  return {sessionId,employeeId,name:String(user.NAME||''),role,solId:office.solId,officeName:office.officeName,expiresAt:expires.toISOString(),today:todayKey_()};
}

function validateSession_(sessionId, requiredRole) {
  if(!sessionId) throw new Error('Session expired. Please login again.');
  const sh=getSheet_('SESSION_LOG'), rows=rowsAsObjects_(sh), s=rows.find(x=>String(x.SESSION_ID)===String(sessionId));
  if(!s) throw new Error('Invalid session.');
  const exp=new Date(s.EXPIRES_AT);
  if(String(s.STATUS).toUpperCase()!=='ACTIVE' || exp.getTime()<Date.now()) throw new Error('Session expired. Please login again.');
  if(requiredRole && String(s.ROLE).toUpperCase()!==String(requiredRole).toUpperCase()) throw new Error('Unauthorized.');
  const data=sh.getDataRange().getValues(), h=data[0].map(String), sid=h.indexOf('SESSION_ID'), last=h.indexOf('LAST_SEEN');
  for(let i=1;i<data.length;i++){if(String(data[i][sid])===String(sessionId)){sh.getRange(i+1,last+1).setValue(new Date());break;}}
  return {sessionId:s.SESSION_ID,employeeId:String(s.EMPLOYEE_ID),role:String(s.ROLE).toUpperCase(),solId:String(s.SOL_ID),officeName:String(s.OFFICE_NAME),ip:String(s.IP_ADDRESS||''),expiresAt:exp.toISOString()};
}

function logout(sessionId){
  try{const sh=getSheet_('SESSION_LOG'),d=sh.getDataRange().getValues(),h=d[0].map(String),sid=h.indexOf('SESSION_ID'),st=h.indexOf('STATUS'),lo=h.indexOf('LOGOUT_TIME');for(let i=1;i<d.length;i++){if(String(d[i][sid])===String(sessionId)){sh.getRange(i+1,st+1).setValue('LOGGED_OUT');sh.getRange(i+1,lo+1).setValue(new Date());break;}}}catch(e){}
  return true;
}

function number_(v,label){
  const n=Number(v);
  if(!Number.isInteger(n)||n<0||n>CFG.MAX_VALUE) throw new Error(label+' must be a whole number from 0 to '+CFG.MAX_VALUE+'.');
  return n;
}

function getTodayStatus(sessionId){
  const s=validateSession_(sessionId,'SPM');
  const rec=rowsAsObjects_(getSheet_('DAILY_RECORD')).find(r=>String(r.REPORT_DATE)===todayKey_()&&String(r.EMPLOYEE_ID)===s.employeeId&&String(r.SOL_ID)===s.solId);
  return rec?{submitted:true,record:sanitizeRecord_(rec)}:{submitted:false};
}

function sanitizeRecord_(r){
  return {
    reportDate:String(r.REPORT_DATE), employeeId:String(r.EMPLOYEE_ID), solId:String(r.SOL_ID), officeName:String(r.OFFICE_NAME),
    allKits:Number(r.ALL_KITS||0), similarArticle:Number(r.SIMILAR_ARTICLE||0),
    invalidMobileKits:Number(r.INVALID_MOBILE_KITS||0), invalidMobileArticles:Number(r.INVALID_MOBILE_ARTICLES||0),
    deliverableKits:Number(r.DELIVERABLE_KITS||0), deliverableArticles:Number(r.DELIVERABLE_ARTICLES||0),
    incompleteKits:Number(r.INCOMPLETE_KITS||0), incompleteArticles:Number(r.INCOMPLETE_ARTICLES||0),
    withoutProperDetailsKits:Number(r.WITHOUT_PROPER_DETAILS_KITS||0), withoutProperDetailsArticles:Number(r.WITHOUT_PROPER_DETAILS_ARTICLES||0),
    kitTotalCheck:String(r.KIT_TOTAL_CHECK), articleTotalCheck:String(r.ARTICLE_TOTAL_CHECK), status:String(r.STATUS)
  };
}

function getMyRecords(sessionId, limit){
  const s=validateSession_(sessionId,'SPM');
  const n=Math.max(1,Math.min(Number(limit)||30,100));
  return rowsAsObjects_(getSheet_('DAILY_RECORD')).filter(r=>String(r.EMPLOYEE_ID)===s.employeeId&&String(r.SOL_ID)===s.solId).slice(-n).reverse().map(sanitizeRecord_);
}

function submitDaily(sessionId,payload){
  const s=validateSession_(sessionId,'SPM'); if(!payload)throw new Error('No report data received.');
  const n={
    allKits:number_(payload.allKits,'All Kits'),
    similarArticle:number_(payload.similarArticle,'Similar Article / Tool Kit Article'),
    invalidMobileKits:number_(payload.invalidMobileKits,'Invalid Mobile Kits'),
    invalidMobileArticles:number_(payload.invalidMobileArticles,'Invalid Mobile Articles'),
    deliverableKits:number_(payload.deliverableKits,'Deliverable Kits'),
    deliverableArticles:number_(payload.deliverableArticles,'Deliverable Articles'),
    incompleteKits:number_(payload.incompleteKits,'Incomplete Kits'),
    incompleteArticles:number_(payload.incompleteArticles,'Incomplete Articles'),
    withoutProperDetailsKits:number_(payload.withoutProperDetailsKits,'Without Proper Details Kits'),
    withoutProperDetailsArticles:number_(payload.withoutProperDetailsArticles,'Without Proper Details Articles')
  };
  const expectedKits=n.invalidMobileKits+n.deliverableKits+n.incompleteKits+n.withoutProperDetailsKits;
  const expectedArticles=n.invalidMobileArticles+n.deliverableArticles+n.incompleteArticles+n.withoutProperDetailsArticles;
  if(n.allKits!==expectedKits) throw new Error('Kit logic failed: All Kits must equal Invalid Mobile + Deliverable + Incomplete + Without Proper Details/Address.');
  if(n.similarArticle!==expectedArticles) throw new Error('Article logic failed: Similar Article / Tool Kit Article must equal the four article categories.');
  const lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    const sh=getSheet_('DAILY_RECORD');
    const dup=rowsAsObjects_(sh).some(r=>String(r.REPORT_DATE)===todayKey_()&&String(r.EMPLOYEE_ID)===s.employeeId&&String(r.SOL_ID)===s.solId);
    if(dup)throw new Error('Today\'s report is already submitted for this office. Duplicate submission blocked by server.');
    const now=new Date();
    sh.appendRow([
      now,todayKey_(),s.employeeId,s.employeeId.slice(-4),s.solId,s.officeName,
      n.allKits,n.similarArticle,n.invalidMobileKits,n.invalidMobileArticles,n.deliverableKits,n.deliverableArticles,
      n.incompleteKits,n.incompleteArticles,n.withoutProperDetailsKits,n.withoutProperDetailsArticles,
      n.allKits===expectedKits,n.similarArticle===expectedArticles,'SUBMITTED',String(payload.ip||s.ip||'Unavailable'),s.sessionId
    ]);
    return {ok:true,message:'Daily report submitted successfully.',record:sanitizeRecord_({REPORT_DATE:todayKey_(),EMPLOYEE_ID:s.employeeId,SOL_ID:s.solId,OFFICE_NAME:s.officeName,...{
      ALL_KITS:n.allKits,SIMILAR_ARTICLE:n.similarArticle,INVALID_MOBILE_KITS:n.invalidMobileKits,INVALID_MOBILE_ARTICLES:n.invalidMobileArticles,
      DELIVERABLE_KITS:n.deliverableKits,DELIVERABLE_ARTICLES:n.deliverableArticles,INCOMPLETE_KITS:n.incompleteKits,INCOMPLETE_ARTICLES:n.incompleteArticles,
      WITHOUT_PROPER_DETAILS_KITS:n.withoutProperDetailsKits,WITHOUT_PROPER_DETAILS_ARTICLES:n.withoutProperDetailsArticles,KIT_TOTAL_CHECK:true,ARTICLE_TOTAL_CHECK:true,STATUS:'SUBMITTED'
    }})};
  }finally{lock.releaseLock();}
}

function adminSummary(sessionId,filters){
  validateSession_(sessionId,'ADMIN');
  const records=rowsAsObjects_(getSheet_('DAILY_RECORD')), offices=rowsAsObjects_(getSheet_('OFFICE_MASTER'));
  const from=String(filters.from||''),to=String(filters.to||''),solId=String(filters.solId||''),date=String(filters.date||'');
  const out=records.filter(r=>{const d=String(r.REPORT_DATE);if(from&&d<from)return false;if(to&&d>to)return false;if(date&&d!==date)return false;if(solId&&String(r.SOL_ID)!==solId)return false;return true;}).map(sanitizeRecord_);
  const totals=out.reduce((a,r)=>{
    a.allKits+=r.allKits;a.similarArticle+=r.similarArticle;a.invalid+=r.invalidMobileKits;a.invalidArticles+=r.invalidMobileArticles;
    a.deliverable+=r.deliverableKits;a.deliverableArticles+=r.deliverableArticles;a.incomplete+=r.incompleteKits;a.incompleteArticles+=r.incompleteArticles;
    a.details+=r.withoutProperDetailsKits;a.detailsArticles+=r.withoutProperDetailsArticles;return a;
  },{allKits:0,similarArticle:0,invalid:0,invalidArticles:0,deliverable:0,deliverableArticles:0,incomplete:0,incompleteArticles:0,details:0,detailsArticles:0});
  const submitted={};out.forEach(r=>submitted[r.solId]=true);
  const pendingOffices=offices.filter(o=>String(o.ACTIVE).toUpperCase()!=='FALSE').filter(o=>!submitted[String(o.SOL_ID)]).map(o=>({solId:String(o.SOL_ID),officeName:String(o.OFFICE_NAME)}));
  return {rows:out,totals,pendingOffices,totalOffices:offices.filter(o=>String(o.ACTIVE).toUpperCase()!=='FALSE').length,submittedOfficeCount:Object.keys(submitted).length};
}

function adminSessions(sessionId,limit){
  validateSession_(sessionId,'ADMIN');
  return rowsAsObjects_(getSheet_('SESSION_LOG')).slice(-Math.min(Number(limit)||200,1000)).reverse().map(r=>({sessionId:String(r.SESSION_ID),loginTime:String(r.LOGIN_TIME),lastSeen:String(r.LAST_SEEN),logoutTime:String(r.LOGOUT_TIME||''),employeeId:String(r.EMPLOYEE_ID),role:String(r.ROLE),solId:String(r.SOL_ID),officeName:String(r.OFFICE_NAME),ip:String(r.IP_ADDRESS||''),userAgent:String(r.USER_AGENT||''),status:String(r.STATUS),expiresAt:String(r.EXPIRES_AT)}));
}
function getOffices(sessionId){validateSession_(sessionId,'ADMIN');return rowsAsObjects_(getSheet_('OFFICE_MASTER')).filter(o=>String(o.ACTIVE).toUpperCase()!=='FALSE').map(o=>({solId:String(o.SOL_ID),officeName:String(o.OFFICE_NAME)}));}
function exportDailyCsv(sessionId,filters){
  const r=adminSummary(sessionId,filters),h=['reportDate','employeeId','solId','officeName','allKits','similarArticle','invalidMobileKits','invalidMobileArticles','deliverableKits','deliverableArticles','incompleteKits','incompleteArticles','withoutProperDetailsKits','withoutProperDetailsArticles','status'];
  return [h.join(','),...r.rows.map(x=>h.map(k=>csv_(x[k])).join(','))].join('\n');
}
function csv_(v){return '"'+String(v==null?'':v).replace(/"/g,'""')+'"';}
