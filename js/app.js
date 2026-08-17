
'use strict';
const CFG=window.PMV_CONFIG||{};
let state={session:null}; const pending=new Map(); const $=id=>document.getElementById(id);
const TODAY=()=>new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Kolkata',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
function show(id,on=true){$(id).classList.toggle('hidden',!on)}
function alertBox(id,text,kind){const e=$(id);e.className='alert '+kind;e.textContent=text;show(id,true)}
function configuredUrl(){return String(CFG.APP_SCRIPT_API_URL||'').trim()}
function validBackend(u){return /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:\?.*)?$/.test(u)}
function request(action,payload){
 return new Promise((resolve,reject)=>{
  const url=configuredUrl(); if(!validBackend(url)){reject(new Error('Apps Script /exec URL is missing or invalid.'));return}
  const requestId=(crypto&&crypto.randomUUID)?crypto.randomUUID():Date.now()+'-'+Math.random().toString(36).slice(2);
  pending.set(requestId,{resolve,reject});
  const f=document.createElement('form'); f.method='POST'; f.action=url; f.target='bridge'; f.style.display='none';
  [['action',action],['requestId',requestId],['payload',JSON.stringify(payload||{})]].forEach(([n,v])=>{const i=document.createElement('input');i.name=n;i.value=v;f.appendChild(i)});
  document.body.appendChild(f);f.submit();f.remove();
  setTimeout(()=>{if(pending.has(requestId)){pending.delete(requestId);reject(new Error('Server response timed out. The Apps Script deployment may be old or not updated.'))}},Number(CFG.REQUEST_TIMEOUT_MS)||45000);
 })
}
window.addEventListener('message',e=>{
 const d=e.data;if(!d||!d.requestId||!pending.has(d.requestId))return;
 const p=pending.get(d.requestId);pending.delete(d.requestId);
 d.ok?p.resolve(d.data):p.reject(new Error(d.data&&d.data.message||'Server request failed.'));
});
function esc(v){return String(v==null?'':v).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function sessionSave(s){state.session=s;sessionStorage.setItem(CFG.SESSION_STORAGE_KEY||'pmvSessionV7',JSON.stringify(s))}
function sessionLoad(){try{const x=sessionStorage.getItem(CFG.SESSION_STORAGE_KEY||'pmvSessionV7');return x?JSON.parse(x):null}catch{return null}}
function sessionClear(){state.session=null;sessionStorage.removeItem(CFG.SESSION_STORAGE_KEY||'pmvSessionV7')}
function fieldIds(){return ['allKits','similarArticle','invalidMobileKits','invalidMobileArticles','deliverableKits','deliverableArticles','incompleteKits','incompleteArticles','withoutProperDetailsKits','withoutProperDetailsArticles']}
async function doLogin(){
 show('loginErr',false);const employeeId=$('employeeId').value.trim();const phone=$('phone').value.replace(/\D/g,'');
 if(!employeeId||phone.length!==10){alertBox('loginErr','Enter Employee ID and a valid 10-digit registered phone number.','error');return}
 $('loginBtn').disabled=true;
 try{const r=await request('login',{employeeId,phone,ip:'Unavailable',userAgent:navigator.userAgent.slice(0,300)});sessionSave(r);renderApp()}catch(e){alertBox('loginErr',e.message||'Login failed.','error')}finally{$('loginBtn').disabled=false}
}
function renderApp(){
 const s=state.session;if(!s)return;
 if(!s.expiresAt||new Date(s.expiresAt)<=new Date()){sessionClear();location.reload();return}
 show('loginScreen',false);show('app',true);$('who').textContent=`${s.name||s.employeeId} · ${s.officeName} · SOL ${s.solId} · Session valid until ${new Date(s.expiresAt).toLocaleString()}`;
 const admin=String(s.role).toUpperCase()==='ADMIN';show('adminView',admin);show('spmView',!admin);
 if(admin){setAdminDates();loadOffices();loadDashboard()}else{$('solId').value=s.solId;$('officeName').value=s.officeName;checkToday();loadMyRecords()}
}
async function checkToday(){
 try{const r=await request('getTodayStatus',{sessionId:state.session.sessionId});if(r.submitted){fillRecord(r.record);alertBox('spmMsg','Today’s report is already submitted. Duplicate submission is blocked by the server.','success')}else{clearEntry();alertBox('spmMsg','No report has been submitted for today.','info')}}catch(e){alertBox('spmMsg',e.message,'error')}
}
function fillRecord(r){fieldIds().forEach(k=>$(k).value=r[k]??'');fieldIds().forEach(k=>$(k).disabled=true);$('submitBtn').disabled=true}
function clearEntry(){fieldIds().forEach(k=>{$(k).value='';$(k).disabled=false});$('submitBtn').disabled=false}
async function submitReport(){
 if(!state.session)return;const report={};let valid=true;
 fieldIds().forEach(k=>{const n=Number($(k).value);if(!Number.isInteger(n)||n<0||n>1000000)valid=false;report[k]=n});
 if(!valid){alertBox('spmMsg','All fields must contain whole numbers from 0 to 1,000,000.','error');return}
 if(report.allKits!==report.invalidMobileKits+report.deliverableKits+report.incompleteKits+report.withoutProperDetailsKits){alertBox('spmMsg','Kit total does not match the four kit categories.','error');return}
 if(report.similarArticle!==report.invalidMobileArticles+report.deliverableArticles+report.incompleteArticles+report.withoutProperDetailsArticles){alertBox('spmMsg','Article total does not match the four article categories.','error');return}
 $('submitBtn').disabled=true;
 try{const r=await request('submitDaily',{sessionId:state.session.sessionId,report});alertBox('spmMsg',r.message||'Report submitted successfully.','success');fillRecord(r.record||report);loadMyRecords()}catch(e){alertBox('spmMsg',e.message,'error');$('submitBtn').disabled=false}
}
async function loadMyRecords(){
 try{const rows=await request('getMyRecords',{sessionId:state.session.sessionId,limit:30});$('myTable').innerHTML='<thead><tr><th>Date</th><th>All Kits</th><th>Deliverable</th><th>Invalid Mobile</th><th>Incomplete</th><th>Without Details</th><th>Status</th></tr></thead><tbody>'+rows.map(r=>`<tr><td>${esc(r.reportDate)}</td><td>${r.allKits}</td><td>${r.deliverableKits}</td><td>${r.invalidMobileKits}</td><td>${r.incompleteKits}</td><td>${r.withoutProperDetailsKits}</td><td><span class="badge good">${esc(r.status)}</span></td></tr>`).join('')+'</tbody>'}catch(e){alertBox('spmMsg',e.message,'error')}
}
function setAdminDates(){const t=TODAY();$('specificDate').value=t;$('fromDate').value=t;$('toDate').value=t}
async function loadOffices(){
 try{const offices=await request('getOffices',{sessionId:state.session.sessionId});$('filterSol').innerHTML='<option value="">All offices</option>'+offices.map(o=>`<option value="${esc(o.solId)}">${esc(o.officeName)} (${esc(o.solId)})</option>`).join('')}catch(e){alertBox('adminMsg',e.message,'error')}
}
function filters(){return{from:$('fromDate').value,to:$('toDate').value,date:$('specificDate').value,solId:$('filterSol').value}}
async function loadDashboard(){
 const f=filters();if(f.from&&f.to&&f.from>f.to){alertBox('adminMsg','From date cannot be after To date.','error');return}
 try{
  alertBox('adminMsg','Loading consolidated data…','info');
  const [s,c,offices]=await Promise.all([request('adminSummary',{sessionId:state.session.sessionId,filters:f}),request('adminConsolidatedReport',{sessionId:state.session.sessionId,filters:f}),request('getOffices',{sessionId:state.session.sessionId})]);
  renderAdmin(s,c,offices);alertBox('adminMsg','Dashboard updated successfully.','success');
 }catch(e){alertBox('adminMsg',e.message,'error')}
}
function renderAdmin(s,c,offices){
 $('kTotal').textContent=s.totalOffices;$('kSubmitted').textContent=s.submittedOfficeCount;$('kPending').textContent=s.pendingOfficeCount;
 const pct=s.totalOffices?Math.round(s.submittedOfficeCount/s.totalOffices*100):0;$('kCompletion').textContent=pct+'%';$('kProgress').style.width=pct+'%';
 const t=c.consolidated;const cards=[['All Kits',t.allKits],['Similar Articles',t.similarArticle],['Deliverable Kits',t.deliverableKits],['Incomplete Kits',t.incompleteKits],['Invalid Mobile Kits',t.invalidMobileKits],['Without Details Kits',t.withoutProperDetailsKits]];
 $('stats').innerHTML=cards.map(x=>`<div class="kpi"><small>${esc(x[0]).toUpperCase()}</small><strong>${Number(x[1]||0).toLocaleString('en-IN')}</strong><div class="accent"></div></div>`).join('');
 const bySol={};s.rows.forEach(r=>bySol[r.solId]=r);
 $('reportTable').innerHTML='<thead><tr><th>Office</th><th>SOL ID</th><th>Report</th><th>All Kits</th><th>Deliverable</th><th>Invalid</th><th>Incomplete</th><th>Without Details</th><th>Articles</th></tr></thead><tbody>'+offices.map(o=>{const r=bySol[o.solId];return `<tr><td>${esc(o.officeName)}</td><td>${esc(o.solId)}</td><td>${r?'<span class="badge good">SUBMITTED</span>':'<span class="badge bad">PENDING</span>'}</td><td>${r?r.allKits:0}</td><td>${r?r.deliverableKits:0}</td><td>${r?r.invalidMobileKits:0}</td><td>${r?r.incompleteKits:0}</td><td>${r?r.withoutProperDetailsKits:0}</td><td>${r?r.similarArticle:0}</td></tr>`}).join('')+'</tbody>';
 const vals=[['Deliverable',t.deliverableKits],['Invalid Mobile',t.invalidMobileKits],['Incomplete',t.incompleteKits],['Without Details',t.withoutProperDetailsKits]];const max=Math.max(1,...vals.map(x=>Number(x[1]||0)));
 $('bars').innerHTML=vals.map(x=>`<div class="bar-row"><b>${esc(x[0])}</b><div class="bar"><span style="width:${Math.round(Number(x[1]||0)/max*100)}%"></span></div><strong>${Number(x[1]||0).toLocaleString('en-IN')}</strong></div>`).join('');
 $('pendingTable').innerHTML='<thead><tr><th>#</th><th>Office</th><th>SOL ID</th><th>Action</th></tr></thead><tbody>'+s.pendingOffices.map((o,i)=>`<tr><td>${i+1}</td><td>${esc(o.officeName)}</td><td>${esc(o.solId)}</td><td><span class="badge bad">UPDATE REQUIRED</span></td></tr>`).join('')+'</tbody>';
}
async function loadSessions(){
 show('sessionPanel',true);try{const rows=await request('adminSessions',{sessionId:state.session.sessionId,limit:100});$('sessionTable').innerHTML='<thead><tr><th>Login</th><th>Employee ID</th><th>Role</th><th>Office</th><th>Status</th><th>Expires</th></tr></thead><tbody>'+rows.map(r=>`<tr><td>${esc(r.loginTime)}</td><td>${esc(r.employeeId)}</td><td>${esc(r.role)}</td><td>${esc(r.officeName)}</td><td>${esc(r.status)}</td><td>${esc(r.expiresAt)}</td></tr>`).join('')+'</tbody>'}catch(e){alertBox('adminMsg',e.message,'error')}
}
async function downloadCsv(){
 try{const csv=await request('exportDailyCsv',{sessionId:state.session.sessionId,filters:filters()});const blob=new Blob([csv],{type:'text/csv;charset=utf-8'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='PMV_Daily_Report_'+TODAY()+'.csv';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}catch(e){alertBox('adminMsg',e.message,'error')}
}
async function doLogout(){try{if(state.session)await request('logout',{sessionId:state.session.sessionId})}catch{}sessionClear();location.reload()}
function refreshCurrent(){if(!state.session)return;if(String(state.session.role).toUpperCase()==='ADMIN')loadDashboard();else{checkToday();loadMyRecords()}}
(function boot(){const s=sessionLoad();if(s){state.session=s;renderApp()}else{show('loginScreen',true)}})();
