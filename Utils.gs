/*******************************************************
 * Utils.gs — Shared utilities & core app plumbing
 *******************************************************/

const CONFIG = {
  ADMIN_SPREADSHEET_ID: '1VlrlxDQJZ2qpVB_ssSTurIf01i7j1Q5VDI1LGU_OTp8',
  ONWARD_SPREADSHEET_ID: '1s8EIoOhvoTQdaDwe584nuP-NF9evQo-z6dPzGWewz0M',
  ITAM_SPREADSHEET_ID: '1o9QLOSy78K09WJVqYYUHEFURLrPpPWGq3M_9blr48Ak',
  IREAL_SPREADSHEET_ID: '1gTSs6Z9NEyxGiU76wt9Au4xFeIkwY701Zk4ZCRDbgKQ',
  LATTE_SPREADSHEET_ID: '1ycyYIxCflwMSHDZS_Yn_xjEMBgrjAcJhy5Zro_G_Dks',
  SHEETS: { USERS:'USERS', LIST:'List', TMS:'TMS', CONCERNS:'Employee Concerns' },
  TIMEZONE: 'Asia/Manila',
  TIME_FMT: 'MM-dd-yyyy HH:mm:ss',
  SESSION_KEY: 'HR_TMS_SESSION',
  SESSION_TTL_MIN: 720,
  SELECTED_CO_KEY: 'HR_SELECTED_CO',
};

/* ---------------- Web App Entrypoint & Views ---------------- */
function doGet(e){
  const qp = (e && e.parameter) || {};
  if (String(qp.logout || '') === '1') { try{ logout(); }catch(_){ } }
  const session = getSession();
  const view = session ? 'Dashboard' : 'Login';
  return evalView_(view);
}
function evalView_(viewName){
  try {
    const tpl = HtmlService.createTemplateFromFile(viewName);
    return tpl.evaluate()
      .setTitle('HR Service Monitoring')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err){
    const msg = 'Failed to render view "'+viewName+'". Ensure '+viewName+
                '.html exists. Details: ' + (err && err.message ? err.message : err);
    return HtmlService.createHtmlOutput(
      '<pre style="color:#fca5a5;background:#111;padding:12px;border-radius:8px">'+
      msg.replace(/</g,'&lt;').replace(/>/g,'&gt;')+'</pre>');
  }
}
function include(filename){ return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function renderView_(viewName){ return HtmlService.createTemplateFromFile(viewName).evaluate().getContent(); }
function loginAndRender(email, password){ const res = login(email,password); if(!res||!res.ok) return res; return {ok:true, html:renderView_('Dashboard')}; }
function logoutAndRender(){ try{ logout(); }catch(_){ } return {ok:true, html:renderView_('Login')}; }
function getWebAppUrl(){ return ScriptApp.getService().getUrl(); }

/* ---------------- Sheet Openers ---------------- */
function getAdminSs()  { return SpreadsheetApp.openById(CONFIG.ADMIN_SPREADSHEET_ID); }
function getOnwardSs() { return SpreadsheetApp.openById(CONFIG.ONWARD_SPREADSHEET_ID); }
function getItamSs()   { return SpreadsheetApp.openById(CONFIG.ITAM_SPREADSHEET_ID); }
function getIrealSs()  { return SpreadsheetApp.openById(CONFIG.IREAL_SPREADSHEET_ID); }
function getLatteSs()  { return SpreadsheetApp.openById(CONFIG.LATTE_SPREADSHEET_ID); }

/* ---------------- Generic Helpers ---------------- */
function getSheetByAnyName_(ss, names){
  for (const n of names){ const sh = ss.getSheetByName(n); if (sh) return sh; }
  const map = new Map(); ss.getSheets().forEach(s => map.set(s.getName().toLowerCase(), s));
  for (const n of names){ const hit = map.get(String(n).toLowerCase()); if (hit) return hit; }
  return null;
}
function readSheetAsObjects(sheet){
  if (!sheet) throw new Error('Target sheet not found.');
  const vals = sheet.getDataRange().getDisplayValues();
  if (vals.length < 2) return [];
  const headers = vals[0].map(h => String(h).trim());
  return vals.slice(1)
    .filter(r => r.some(c => String(c).trim() !== ''))
    .map(row => { const o={}; headers.forEach((h,i)=>o[h]=row[i]); return o; });
}
function formatNow(){ return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, CONFIG.TIME_FMT); }

/* ---------------- Auth ---------------- */
function getUsersSheet_(){
  const sh = getSheetByAnyName_(getAdminSs(), [CONFIG.SHEETS.USERS,'USERS','users']);
  if (!sh) throw new Error('Create "USERS" sheet with headers.');
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(h=>String(h).trim());
  ['Email','Password','Role','Department','Company ID','Display Name','CompanyCode','AccountCode']
    .forEach(req=>{
      if (!headers.some(h => h.toLowerCase() === req.toLowerCase())) {
        throw new Error('USERS sheet missing header "'+req+'". Found: '+headers.join(' | '));
      }
    });
  return sh;
}
function normalizeCreds_(email,password){ return { e:String(email||'').trim().toLowerCase(), p:String(password||'').trim() }; }
function login(email,password){
  const {e,p} = normalizeCreds_(email,password);
  const rows = readSheetAsObjects(getUsersSheet_());
  const found = rows.find(r => String(r['Email']||'').trim().toLowerCase() === e);
  if (!found) return {ok:false, message:'User not found.'};
  if (String(found['Password']||'').trim() !== p) return {ok:false, message:'Invalid password.'};
  const companyCode = String(found['CompanyCode']||'').trim().toUpperCase();
  const accountCode = String(found['AccountCode']||'').trim().toUpperCase();
  const session = {
    email:String(found['Email']||'').trim(),
    role:String(found['Role']||'').trim() || 'Employee',
    department:String(found['Department']||'').trim(),
    displayName:String(found['Display Name']||found['DisplayName']||'').trim(),
    companyId:String(found['Company ID']||found['CompanyID']||'').trim(),
    companyCode, accountCode,
    issuedAt: Date.now(),
    expiresAt: Date.now() + CONFIG.SESSION_TTL_MIN*60*1000
  };
  PropertiesService.getUserProperties().setProperty(CONFIG.SESSION_KEY, JSON.stringify(session));
  let selected = 'Onward';
  if (companyCode==='ITM'||companyCode==='ITAM') selected='ITAM';
  else if (companyCode==='ONW'||companyCode==='ONWARD') selected='Onward';
  else if (companyCode==='IREAL') selected='IREAL';
  else if (companyCode==='LATTE') selected='LATTE';
  PropertiesService.getUserProperties().setProperty(CONFIG.SELECTED_CO_KEY, selected);
  return {ok:true, message:'Login successful.', session};
}
function logout(){ PropertiesService.getUserProperties().deleteProperty(CONFIG.SESSION_KEY); return {ok:true}; }
function getSession(){ const raw=PropertiesService.getUserProperties().getProperty(CONFIG.SESSION_KEY); if(!raw) return null; try{ const s=JSON.parse(raw); if(!s.expiresAt||Date.now()>s.expiresAt){ logout(); return null; } return s; }catch{ logout(); return null; } }
function requireAuth(){ const s=getSession(); if(!s) throw new Error('Not authenticated.'); return s; }
function getSessionInfoForClient(){ const s=getSession(); if(!s) return null; return { email:s.email, role:s.role, department:s.department, displayName:s.displayName, companyId:s.companyId, companyCode:s.companyCode, accountCode:s.accountCode }; }

/* ---------------- Company selection + routing ---------------- */
function canonicalCompany_(company){ const v=String(company||'').trim().toUpperCase(); if(v==='ONWARD')return'Onward'; if(v==='ITAM')return'ITAM'; if(v==='IREAL')return'IREAL'; if(v==='LATTE')return'LATTE'; return 'Onward'; }
function setCompany(company){ requireAuth(); const cRaw=String(company||'').trim(); const u=cRaw.toUpperCase(); const ok=['ONWARD','ITAM','IREAL','LATTE','Onward','ITAM','IREAL','LATTE']; if(!ok.includes(cRaw)&&!ok.includes(u)) throw new Error('Invalid company selected.'); const c=canonicalCompany_(cRaw); PropertiesService.getUserProperties().setProperty(CONFIG.SELECTED_CO_KEY, c); return {ok:true, company:c}; }
function getCompany(){ try{ const v=PropertiesService.getUserProperties().getProperty(CONFIG.SELECTED_CO_KEY); return v || 'Onward'; }catch(e){ return 'Onward'; } }
function getSpreadsheetByCompany_(name){ const c=canonicalCompany_(name); if(c==='Onward')return getOnwardSs(); if(c==='ITAM')return getItamSs(); if(c==='IREAL')return getIrealSs(); if(c==='LATTE')return getLatteSs(); return getOnwardSs(); }
function getActiveSpreadsheet_(){ const company=getCompany(); return getSpreadsheetByCompany_(company); }

/* ---------------- Company code ↔ name & SS mapping ---------------- */
const COMPANY_PREFIX_TO_NAME = Object.freeze({ ONW:'Onward', ITM:'ITAM', IRL:'IREAL', LTE:'LATTE' });
function companyFromRequestId_(requestId){ const m=String(requestId||'').toUpperCase().match(/^([A-Z]{3})-/); if(!m) return null; return COMPANY_PREFIX_TO_NAME[m[1]]||null; }
function companyNameFromCode_(code){ const key=String(code||'').toUpperCase().slice(0,3); return COMPANY_PREFIX_TO_NAME[key]||null; }
function getCompanySpreadsheet_(companyName){ const key=String(companyName||'').trim().toUpperCase(); switch(key){ case'ONWARD':return getOnwardSs(); case'ITAM':return getItamSs(); case'IREAL':return getIrealSs(); case'LATTE':return getLatteSs(); default: throw new Error('Unknown company for routing: '+key); } }
function getSpreadsheetFromRequestId_(requestId){
  const company = companyFromRequestId_(requestId);
  return company ? getCompanySpreadsheet_(company) : getActiveSpreadsheet_();
}

/* ---------------- Request ID + Appends ---------------- */
function generateRequestId_(companyCodeParam, accountCodeParam){
  const s = requireAuth();
  const companyCode = String((companyCodeParam||s.companyCode)||'').toUpperCase();
  const accountCode = String((accountCodeParam||s.accountCode)||'').toUpperCase();
  if (!companyCode || !accountCode) throw new Error('Missing company/account for Request ID.');
  const companyName = companyNameFromCode_(companyCode) || canonicalCompany_(getCompany());
  const ss = getCompanySpreadsheet_(companyName);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try{
    const sh = ss.getSheetByName(CONFIG.SHEETS.TMS);
    if (!sh) throw new Error('TMS sheet not found.');
    const vals = sh.getDataRange().getDisplayValues();
    const hdr = vals.length ? vals[0] : [];
    const idIdx = hdr.findIndex(h => String(h).toLowerCase() === 'request id');
    if (idIdx === -1) throw new Error('TMS missing "Request ID" column.');
    let seq = 1;
    const prefix = companyCode + '-' + accountCode + '-';
    for (let r=1;r<vals.length;r++){
      const cell=String(vals[r][idIdx]||'').trim();
      if (cell.startsWith(prefix)){
        const n = parseInt(cell.split('-').pop(),10);
        if (Number.isFinite(n) && n >= seq) seq = n + 1;
      }
    }
    const num = ('000' + seq).slice(-4);
    return `${companyCode}-${accountCode}-${num}`;
  } finally { try{ lock.releaseLock(); }catch(_){ } }
}
function appendToCompanyTms_(companyName, rowObj){
  const ss = getCompanySpreadsheet_(companyName);
  const sh = ss.getSheetByName(CONFIG.SHEETS.TMS);
  if (!sh) throw new Error('TMS sheet not found in ' + companyName);
  const lastCol = sh.getLastColumn() || 1;
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||'').trim());
  const norm = s => String(s||'').replace(/\s+/g,'').toLowerCase();
  const row = headers.map(h=>{
    const direct = rowObj[h]; if (direct !== undefined) return direct;
    const target = norm(h);
    for (const k in rowObj){ if (norm(k) === target) return rowObj[k]; }
    return '';
  });
  sh.getRange(sh.getLastRow()+1, 1, 1, headers.length).setValues([row]);
}
function getOrCreateUploadFolder_(companyCode){
  const name='HR Uploads – '+String(companyCode||'GEN').toUpperCase();
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

/* ---------------- App bootstrap helpers ---------------- */
function getBootstrapDataLight(){ requireAuth(); return { user:getSessionInfoForClient(), services:getServices() }; }
function getBootstrapData(){
  requireAuth();
  return {
    user: getSessionInfoForClient(),
    services: getServices(),
    logs: getLogData(),
    concerns: getConcernsData(),
    kpis: getKPIData()
  };
}
function getKPIData(){
  const session = requireAuth();
  const role = (session.role||'').toLowerCase();
  const email = (session.email||'').toLowerCase();
  const dept  = (session.department||'').toLowerCase();

  let logs = getLogData();
  if (role === 'employee'){
    const emailKey = logs.length ? Object.keys(logs[0]).find(k => k.toLowerCase().includes('email')) : null;
    if (emailKey) logs = logs.filter(r => String(r[emailKey]||'').toLowerCase() === email);
  } else if (role === 'department head'){
    logs = logs.filter(r => (r['Department']||'').toLowerCase() === dept);
  }

  let concerns = getConcernsData();
  if (role === 'employee'){
    const nameKey = concerns.length ? (Object.keys(concerns[0]).includes('Employee Name') ? 'Employee Name' : null) : null;
    if (nameKey) concerns = concerns.filter(r => (r[nameKey]||'').toLowerCase().includes(email.split('@')[0]));
  } else if (role === 'department head'){
    concerns = concerns.filter(r => (r['Department']||'').toLowerCase() === dept);
  }

  const tatByReq = {};
  logs.forEach(r => { const id=r['Request ID']||''; const tat=parseInt(r['TAT (mins)']||'0',10)||0; tatByReq[id]=(tatByReq[id]||0)+tat; });
  const tatArray = Object.values(tatByReq);
  const averageTAT = tatArray.length ? Math.round(tatArray.reduce((a,b)=>a+b,0)/tatArray.length) : 0;

  let ontime=0, delayed=0;
  concerns.forEach(c => {
    const s=(c['SLA Status']||'').toLowerCase();
    if (s.includes('on-time') || s.includes('ontime') || s==='on time') ontime++;
    else if (s.includes('delay') || s==='late') delayed++;
  });
  const total = ontime + delayed;
  const slaCompliance = total ? Math.round((ontime/total)*100) : 0;

  const byDept = {};
  logs.forEach(r => { const d=r['Department']||'Unspecified'; byDept[d]=(byDept[d]||0)+1; });
  const requestsByDepartment = Object.keys(byDept).map(k => ({department:k, count:byDept[k]}));

  return {
    userEmail: session.email,
    role: session.role,
    averageTAT,
    slaCompliance,
    requestsByDepartment,
    delayedVsOnTime: [{label:'On-time', value:ontime}, {label:'Delayed', value:delayed}]
  };
}

/* ---------------- Compatibility wrappers for renamed backends ---------------- */
// These wrappers preserve existing frontend calls.
function getServices(){ return getServicesBackend(); }
function getProcessSteps(service){ return getProcessStepsBackend(service); }
function getCategoryOptions(){ return getCategoryOptionsBackend(); }
function getLogData(){ return getLogDataBackend(); }
function getFilteredLogData(filters){ return getFilteredLogDataBackend(filters); }
function logAction(action, requestId, service, processStep, details){ return logActionBackend(action, requestId, service, processStep, details); }
function getConcernsDataActiveCompany(opts){ return getConcernsDataActiveCompanyBackend(opts); }
function getConcernsData(){ return getConcernsDataBackend(); }
function getKPIReportDataV2(filters){ return getKPIReportDataV2Backend(filters); }
function getKPIReportDataV3(filters){ return getKPIReportDataV3Backend(filters); }
function submitForm(dto, files){ return submitFormBackend(dto, files); }
function submitFormAndRouteByCompany(data, files){ return submitFormAndRouteByCompanyBackend(data, files); }
