/**
 * ════════════════════════════════════════════════════════════════════
 *  CGSO · Property & Asset Management System
 *  Google Apps Script Web App — Code.gs  v4.0
 *
 *  HOW TO DEPLOY:
 *  1. Apps Script editor → select ALL → delete → paste this file
 *  2. Save (Ctrl+S)
 *  3. Deploy → Manage deployments → pencil → New version → Deploy
 *
 *  Verify: open <your-exec-url>?action=version  → should show "4.0"
 *  Diagnostics: <url>?action=diag
 *  Default admin:  admin / Admin@1234
 *
 *  v4.0 — Secure User Management + PAR/ICS Sheet Separation:
 *    • PAR records saved to "PAR_Records" sheet
 *    • ICS records saved to "ICS_Records" sheet
 *    • Legacy "Properties" sheet kept for backwards compatibility
 *    • New user accounts get auto-generated temp password + email
 *    • Forgot Password: sendOTP → verifyOTP → resetPassword (email OTP)
 *    • OTPs stored in PropertiesService with 10-min TTL + brute-force protection
 *    • migratePasswords endpoint fixes any legacy hashed passwords
 * ════════════════════════════════════════════════════════════════════
 */

// ─── ❶  CONFIGURATION ─────────────────────────────────────────────
const _RAW_SPREADSHEET_ID  = '17gYWzaaaOZsrmjA5kZh5rch5PIqH4LL9qv4V19PmXRo';
const _RAW_DRIVE_FOLDER_ID = '10bd3yZVuJF8w-WthAzs8qy48u_7PSNjW';

const SPREADSHEET_ID = (function(raw) {
  var s = String(raw||'').trim();
  var m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return s.split('/')[0].split('?')[0].split('#')[0];
})(_RAW_SPREADSHEET_ID);

const DRIVE_FOLDER_ID = (function(raw) {
  var s = String(raw||'').trim();
  var m = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return s.split('/')[0].split('?')[0].split('#')[0];
})(_RAW_DRIVE_FOLDER_ID);

const APP_TITLE        = 'AMIS · Asset Management and Inventory System';
const SESSION_TTL_MS   = 6 * 60 * 60 * 1000;
const OTP_TTL_MS       = 10 * 60 * 1000;
const MAX_UPLOAD_BYTES = 10 * 1024 * 1024;
// How long (ms) to cache the "sheets are initialized" flag — 24 hours
const INIT_CACHE_TTL_MS = 24 * 60 * 60 * 1000;
// ⚑  Bump this key whenever new sheets are added (e.g. PRS_Records).
//    The old key stays in PropertiesService but is never read — the next
//    request finds V2 missing, calls ensureSheets(), creates the new sheet,
//    then writes V2.  Cost: one extra ensureSheets() call per instance.
const _INIT_KEY = 'CGSO_SHEETS_INIT_V2';

// ─── ❷  SHEET NAMES ───────────────────────────────────────────────
const SH = {
  PROPERTIES   : 'Properties',
  PAR_RECORDS  : 'PAR_Records',
  ICS_RECORDS  : 'ICS_Records',
  PRS_RECORDS  : 'PRS_Records',
  REAL_PROP    : 'RealProperty',
  HISTORY      : 'AssetHistory',
  OFFICERS     : 'Officers',
  DEPARTMENTS  : 'Departments',
  PPE_MAIN     : 'PPEMainCategories',
  PPE_SUB      : 'PPESubCategories',
  FUND_CLUSTERS: 'FundClusters',
  CONDITIONS   : 'Conditions',
  ACQUISITIONS : 'AcquisitionModes',
  ACCT_STATUSES: 'AccountingStatuses',
  USERS        : 'Users',
  AUDIT_LOG    : 'AuditLog',
  CONFIG       : 'Config',
};

// ─── ❸  COLUMN HEADERS ────────────────────────────────────────────
const PROP_COLS = [
  'id','propno','name','category','acctcode','subcatcode','qty',
  'fundcluster','unit','serial','brand','specification','maincategory',
  'cost','doctype','dept','custodian','designation','location',
  'status','remarks','plateno','color','engineno','chassisno',
  'supplier','pono','podate','drno','drdate','accountno','voucheramt','acctstatus',
  'issuedby','issuedpos','notedby','notedpos','preparedby','preparedpos',
  'date','pdfLink','pdfName','createdAt','updatedAt','createdBy',
];

// PRS_Records extends PROP_COLS with PRS-specific fields
const PRS_COLS = [
  ...PROP_COLS,
  'prsPurpose',   // Disposal | Repair | Returned to Stock | Unserviceable
  'prsDate',      // Date of return
  'prsReceivedBy',// Officer who received the returned item
  'prsReceivedByPos',
  'prsRemarks',
  'prsNo',        // PRS document number
  'prsProcessedBy',
  'prsProcessedAt',
];

const HDR = {
  PROPERTIES   : PROP_COLS,
  PAR_RECORDS  : PROP_COLS,
  ICS_RECORDS  : PROP_COLS,
  PRS_RECORDS  : PRS_COLS,
  REAL_PROP: [
    'id','recno','rtype','category','name','location','lotno','titleno','taxdecno',
    'area','appraisedval','cost','assessedval','date','acquisition',
    'fundcluster','dept','yearconstruct','floorarea','floors','material',
    'status','remarks','pdfLink','pdfName','createdAt','updatedAt','createdBy',
  ],
  HISTORY: ['id','propId','type','officer','designation','dept','date','location','remarks','recordedBy','recordedAt'],
  OFFICERS     : ['id','name','designation','dept'],
  DEPARTMENTS  : ['id','name','head','code'],
  PPE_MAIN     : ['id','name','code'],
  PPE_SUB      : ['id','name','mainCatId','code'],
  FUND_CLUSTERS: ['id','code','name'],
  CONDITIONS   : ['id','label','style'],
  ACQUISITIONS : ['id','name','desc'],
  ACCT_STATUSES: ['id','name','desc'],
  USERS: ['id','name','username','password','email','dept','designation','role','status','permissions','createdAt','lastLogin'],
  AUDIT_LOG: ['id','ts','action','property','user','details'],
  CONFIG    : ['key','value'],
};

// ═══════════════════════════════════════════════════════════════════
//  ❹  SESSION STORE
// ═══════════════════════════════════════════════════════════════════
const _SESS_KEY = 'CGSO_SESSIONS_V2';
const _OTP_KEY  = 'CGSO_OTPS_V1';

function _loadSessions() {
  try { const r = PropertiesService.getScriptProperties().getProperty(_SESS_KEY); return r ? JSON.parse(r) : {}; }
  catch(e) { logErr('_loadSessions',e); return {}; }
}
function _saveSessions(map) {
  try {
    // Prune expired sessions before every save to keep the property size small
    const now = Date.now();
    Object.keys(map).forEach(k => { if (map[k].expires < now) delete map[k]; });
    const json = JSON.stringify(map);
    PropertiesService.getScriptProperties().setProperty(_SESS_KEY, json);
  } catch(e) { logErr('_saveSessions',e); throw e; }
}
function _sessionSet(token, data) {
  const map=_loadSessions(), now=Date.now();
  Object.keys(map).forEach(k=>{ if(map[k].expires<now) delete map[k]; });
  map[token]=data; _saveSessions(map);
}
function _sessionGet(token) {
  if(!token) return null;
  const map=_loadSessions(), s=map[token];
  if(!s) return null;
  if(Date.now()>s.expires) { delete map[token]; _saveSessions(map); return null; }
  return s;
}
function _sessionDel(token) {
  if(!token) return;
  const map=_loadSessions();
  if(map[token]) { delete map[token]; _saveSessions(map); }
}

// ─── OTP store ────────────────────────────────────────────────────
function _loadOTPs() {
  try { const r=PropertiesService.getScriptProperties().getProperty(_OTP_KEY); return r?JSON.parse(r):{}; }
  catch(_) { return {}; }
}
function _saveOTPs(map) {
  try { PropertiesService.getScriptProperties().setProperty(_OTP_KEY,JSON.stringify(map)); } catch(_) {}
}
function _setOTP(username, otp) {
  const map=_loadOTPs(), now=Date.now();
  Object.keys(map).forEach(k=>{ if(map[k].expires<now) delete map[k]; });
  map[username]={otp:String(otp), expires:now+OTP_TTL_MS, attempts:0};
  _saveOTPs(map);
}
function _getOTP(username) {
  const map=_loadOTPs(), entry=map[username];
  if(!entry) return null;
  if(Date.now()>entry.expires) { delete map[username]; _saveOTPs(map); return null; }
  return entry;
}
function _deleteOTP(username) { const map=_loadOTPs(); delete map[username]; _saveOTPs(map); }

// ═══════════════════════════════════════════════════════════════════
//  HTTP ENTRY POINTS
// ═══════════════════════════════════════════════════════════════════
function doGet(e) {
  const action=(e&&e.parameter&&e.parameter.action)||'';
  if (action==='ping') {
    let ok=false,err='';
    try{getSpreadsheet();ok=true;}catch(ex){err=ex.message;}
    return jsonOut({ok:true,ts:new Date().toISOString(),v:'4.0',sheetOk:ok,sheetErr:err||undefined});
  }
  if (action==='version') return jsonOut({ok:true,version:'4.3',features:['sendOTP','verifyOTP','resetPassword','separatePARICS','tempPassword','migrateProperties','splitLoad','prsModule'],ts:new Date().toISOString()});
  if (action==='getData')      return jsonOut(getAllData());
  if (action==='getRefData')   return jsonOut(getRefData());
  if (action==='getHeavyData') {
    const page     = parseInt((e&&e.parameter&&e.parameter.page)||'0');
    const pageSize = parseInt((e&&e.parameter&&e.parameter.pageSize)||'500');
    return jsonOut(getHeavyData({page, pageSize}));
  }
  if (action==='diag')       return jsonOut(runDiagnostics());
  if (action==='resetSessions') {
    try{PropertiesService.getScriptProperties().deleteProperty(_SESS_KEY);}catch(_){}
    return jsonOut({ok:true,message:'Sessions cleared.'});
  }
  // Clears the sheet-init cache so ensureSheets() runs on next request.
  // Use after any Code.gs update that adds new sheets.
  // URL: <your-web-app-url>?action=resetInit
  if (action==='resetInit') {
    try{PropertiesService.getScriptProperties().deleteProperty(_INIT_KEY);}catch(_){}
    try{PropertiesService.getScriptProperties().deleteProperty('CGSO_SHEETS_INIT_V1');}catch(_){}
    return jsonOut({ok:true,message:'Init cache cleared. Sheets will be re-checked on next request.'});
  }
  if (action==='migratePasswords')   return jsonOut(migrateHashedPasswords());
  if (action==='migrateProperties')  return jsonOut(migratePropertiesToSeparateSheets());
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const body=JSON.parse(e.postData.contents);
    const action=body.action||'', payload=body.payload||{}, token=body.token||'';

    // Public actions
    if (action==='login')          return jsonOut(handleLogin(payload));
    if (action==='logout')         return jsonOut(handleLogout(token));
    if (action==='getData')        return jsonOut(getAllData());
    if (action==='changePassword') return jsonOut(doChangePassword(payload,token));
    if (action==='sendOTP')        return jsonOut(handleSendOTP(payload));
    if (action==='verifyOTP')      return jsonOut(handleVerifyOTP(payload));
    if (action==='resetPassword')  return jsonOut(handleResetPassword(payload));

    // Session-protected
    const sess=validateToken(token);
    if (!sess) return jsonOut({ok:false,error:'Session expired. Please log in again.'});

    switch(action) {
      case 'saveProperty':       return jsonOut(saveProperty(payload,sess));
      case 'deleteProperty':     return jsonOut(deleteProperty(payload.id,sess));
      case 'savePRS':            return jsonOut(savePRS(payload,sess));
      case 'deletePRS':          return jsonOut(deletePRS(payload.id,sess));
      case 'transferPRS':        return jsonOut(transferPRS(payload,sess));
      case 'saveRealProperty':   return jsonOut(saveRealProperty(payload,sess));
      case 'deleteRealProperty': return jsonOut(deleteRecord(SH.REAL_PROP,payload.id,sess,'Real Property'));
      case 'saveHistory':        return jsonOut(saveHistory(payload,sess));
      case 'getHistory':         return jsonOut(getHistory(payload.propId));
      case 'uploadPDF':          return jsonOut(uploadPDF(payload,sess));
      case 'deletePDF':          return jsonOut(deletePDF(payload,sess));
      case 'saveOfficer':        return jsonOut(saveRow(SH.OFFICERS,payload,sess,'Officer'));
      case 'deleteOfficer':      return jsonOut(deleteRecord(SH.OFFICERS,payload.id,sess,'Officer'));
      case 'saveDepartment':     return jsonOut(saveRow(SH.DEPARTMENTS,payload,sess,'Department'));
      case 'deleteDepartment':   return jsonOut(deleteRecord(SH.DEPARTMENTS,payload.id,sess,'Department'));
      case 'savePPEMain':        return jsonOut(saveRow(SH.PPE_MAIN,payload,sess,'PPE Main Category'));
      case 'deletePPEMain':      return jsonOut(deleteRecord(SH.PPE_MAIN,payload.id,sess,'PPE Main Category'));
      case 'savePPESub':         return jsonOut(saveRow(SH.PPE_SUB,payload,sess,'PPE Subcategory'));
      case 'deletePPESub':       return jsonOut(deleteRecord(SH.PPE_SUB,payload.id,sess,'PPE Subcategory'));
      case 'saveFundCluster':    return jsonOut(saveRow(SH.FUND_CLUSTERS,payload,sess,'Fund Cluster'));
      case 'deleteFundCluster':  return jsonOut(deleteRecord(SH.FUND_CLUSTERS,payload.id,sess,'Fund Cluster'));
      case 'saveCondition':      return jsonOut(saveRow(SH.CONDITIONS,payload,sess,'Condition'));
      case 'deleteCondition':    return jsonOut(deleteRecord(SH.CONDITIONS,payload.id,sess,'Condition'));
      case 'saveAcquisition':    return jsonOut(saveRow(SH.ACQUISITIONS,payload,sess,'Acquisition Mode'));
      case 'deleteAcquisition':  return jsonOut(deleteRecord(SH.ACQUISITIONS,payload.id,sess,'Acquisition Mode'));
      case 'saveAcctStatus':     return jsonOut(saveRow(SH.ACCT_STATUSES,payload,sess,'Accounting Status'));
      case 'deleteAcctStatus':   return jsonOut(deleteRecord(SH.ACCT_STATUSES,payload.id,sess,'Accounting Status'));
      case 'saveUser':           return jsonOut(saveUser(payload,sess));
      case 'deleteUser':         return jsonOut(deleteUser(payload.id,sess));
      case 'getAuditLog':        return jsonOut({ok:true,data:readSheet(SH.AUDIT_LOG)});
      case 'saveAuditLog': {
        // Called by the frontend addLog() helper to persist a frontend-generated log entry
        const p2 = payload;
        const actUser2 = (p2 && p2.user) ? p2.user
                        : (sess && sess.username)   ? sess.username : '';
        try { _auditNoFlush(p2.action||'', p2.property||'', p2.details||'', actUser2); } catch(_) {}
        return jsonOut({ok:true});
      }
      case 'saveConfig':         return jsonOut(saveConfig(payload,sess));
      case 'bulkImport':         return jsonOut(bulkImport(payload,sess));
      default: return jsonOut({ok:false,error:`Unknown action: "${action}"`});
    }
  } catch(err) { logErr('doPost',err); return jsonOut({ok:false,error:err.message}); }
}

// ═══════════════════════════════════════════════════════════════════
//  DIAGNOSTICS
// ═══════════════════════════════════════════════════════════════════
function runDiagnostics() {
  const results=[]; let allOk=true;
  function check(name,fn) {
    try { results.push({name,status:'OK',detail:String(fn())}); }
    catch(e) { results.push({name,status:'FAIL',detail:e.message}); allOk=false; }
  }
  check('SPREADSHEET_ID', ()=>SPREADSHEET_ID);
  check('Open spreadsheet', ()=>{const ss=SpreadsheetApp.openById(cleanSpreadsheetId(SPREADSHEET_ID));return'OK — "'+ss.getName()+'"';});
  check('PropertiesService', ()=>{
    const ps=PropertiesService.getScriptProperties(), k='_cgso_test_'+Date.now();
    ps.setProperty(k,'ok'); const v=ps.getProperty(k); ps.deleteProperty(k);
    if(v!=='ok') throw new Error('Read-back failed'); return 'ok';
  });
  check('ensureSheets', ()=>{ensureSheets();return'All sheets present';});
  check('PAR_Records sheet', ()=>readSheet(SH.PAR_RECORDS).length+' record(s)');
  check('ICS_Records sheet', ()=>readSheet(SH.ICS_RECORDS).length+' record(s)');
  check('Users sheet', ()=>readSheet(SH.USERS).length+' user(s)');
  check('MailApp quota', ()=>MailApp.getRemainingDailyQuota()+' emails remaining today');
  check('DRIVE_FOLDER_ID', ()=>{
    if(DRIVE_FOLDER_ID==='YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE') return '⚠ Not configured';
    return 'OK — "'+DriveApp.getFolderById(DRIVE_FOLDER_ID).getName()+'"';
  });
  return {ok:allOk,diagnostics:results};
}

// ═══════════════════════════════════════════════════════════════════
//  DATA LOAD
// ═══════════════════════════════════════════════════════════════════

/** Read all sheets in a single spreadsheet open — eliminates redundant openById() calls. */
function _readAllSheets(sheetNames) {
  const ss = getSpreadsheet();
  const result = {};
  sheetNames.forEach(name => { result[name] = _readSheetFromSS(ss, name); });
  return result;
}

function getAllData() {
  try {
    ensureSheets();
    return {ok:true, data: _buildFullDataset()};
  } catch(err) { logErr('getAllData',err); return {ok:false,error:err.message}; }
}

/** Lightweight ref data only — officers, depts, PPE, fund clusters, conditions, etc.
 *  Called first after login so the app is usable while heavy data loads. */
function getRefData() {
  try {
    ensureSheetsIfNeeded();
    // One spreadsheet open, all reads batched
    const sheets = _readAllSheets([
      SH.OFFICERS, SH.DEPARTMENTS, SH.PPE_MAIN, SH.PPE_SUB,
      SH.FUND_CLUSTERS, SH.CONDITIONS, SH.ACQUISITIONS,
      SH.ACCT_STATUSES, SH.USERS, SH.CONFIG,
    ]);
    const cfg = {};
    (sheets[SH.CONFIG]||[]).forEach(r=>{ if(r.key) cfg[r.key]=r.value; });
    return {ok:true, data:{
      officers         : sheets[SH.OFFICERS],
      departments      : sheets[SH.DEPARTMENTS],
      ppeMainCategories: sheets[SH.PPE_MAIN],
      ppeSubcategories : sheets[SH.PPE_SUB],
      fundClusters     : sheets[SH.FUND_CLUSTERS],
      conditions       : sheets[SH.CONDITIONS],
      acquisitionModes : sheets[SH.ACQUISITIONS],
      acctStatuses     : sheets[SH.ACCT_STATUSES],
      systemUsers      : sheets[SH.USERS].map(stripPassword),
      config           : cfg,
    }};
  } catch(err) { logErr('getRefData',err); return {ok:false,error:err.message}; }
}

/** Heavy data — properties, real property, history, audit log.
 *  Supports optional server-side pagination via payload.page / payload.pageSize.
 *  Called in the background after the app UI is already usable. */
function getHeavyData(payload) {
  try {
    ensureSheetsIfNeeded();
    const PAGE_SIZE = (payload && payload.pageSize) ? parseInt(payload.pageSize) : 500;
    const page      = (payload && payload.page)     ? parseInt(payload.page)     : 0; // 0 = all

    const sheets = _readAllSheets([
      SH.PAR_RECORDS, SH.ICS_RECORDS, SH.PROPERTIES,
      SH.PRS_RECORDS, SH.REAL_PROP, SH.HISTORY, SH.AUDIT_LOG,
    ]);
    const parRecords = sheets[SH.PAR_RECORDS];
    const icsRecords = sheets[SH.ICS_RECORDS];
    const parIds = new Set(parRecords.map(r=>String(r.id)));
    const icsIds = new Set(icsRecords.map(r=>String(r.id)));
    const legacyProps = sheets[SH.PROPERTIES].filter(p => {
      const id = String(p.id);
      return !parIds.has(id) && !icsIds.has(id);
    });

    let allProps = [...parRecords, ...icsRecords, ...legacyProps];
    const totalCount = allProps.length;

    // Apply pagination when page > 0
    let hasMore = false;
    if (page > 0) {
      const start = (page - 1) * PAGE_SIZE;
      allProps = allProps.slice(start, start + PAGE_SIZE);
      hasMore = (start + PAGE_SIZE) < totalCount;
    }

    return {ok:true, data:{
      properties    : allProps,
      prsRecords    : sheets[SH.PRS_RECORDS],
      realProperties: sheets[SH.REAL_PROP],
      history       : sheets[SH.HISTORY],
      auditLog      : sheets[SH.AUDIT_LOG],
    }, meta:{ totalCount, page, pageSize: PAGE_SIZE, hasMore }};
  } catch(err) { logErr('getHeavyData',err); return {ok:false,error:err.message}; }
}

function _buildFullDataset() {
  const ss = getSpreadsheet();
  const sheets = {};
  Object.values(SH).forEach(name => { sheets[name] = _readSheetFromSS(ss, name); });
  const parRecords = sheets[SH.PAR_RECORDS];
  const icsRecords = sheets[SH.ICS_RECORDS];
  const parIds = new Set(parRecords.map(r=>String(r.id)));
  const icsIds = new Set(icsRecords.map(r=>String(r.id)));
  const legacyProps = sheets[SH.PROPERTIES].filter(p => {
    const id = String(p.id);
    return !parIds.has(id) && !icsIds.has(id);
  });
  const cfg = {};
  (sheets[SH.CONFIG]||[]).forEach(r=>{ if(r.key) cfg[r.key]=r.value; });
  return {
    properties       : [...parRecords,...icsRecords,...legacyProps],
    prsRecords       : sheets[SH.PRS_RECORDS],
    realProperties   : sheets[SH.REAL_PROP],
    history          : sheets[SH.HISTORY],
    officers         : sheets[SH.OFFICERS],
    departments      : sheets[SH.DEPARTMENTS],
    ppeMainCategories: sheets[SH.PPE_MAIN],
    ppeSubcategories : sheets[SH.PPE_SUB],
    fundClusters     : sheets[SH.FUND_CLUSTERS],
    conditions       : sheets[SH.CONDITIONS],
    acquisitionModes : sheets[SH.ACQUISITIONS],
    acctStatuses     : sheets[SH.ACCT_STATUSES],
    systemUsers      : sheets[SH.USERS].map(stripPassword),
    auditLog         : sheets[SH.AUDIT_LOG],
    config           : cfg,
  };
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET BOOTSTRAP
// ═══════════════════════════════════════════════════════════════════

/** Full init — always runs all checks. Used on first load and manual triggers. */
function ensureSheets() {
  const ss=getSpreadsheet();
  Object.entries(SH).forEach(([key,name])=>{
    if (!ss.getSheetByName(name)) {
      const sh=ss.insertSheet(name), headers=HDR[key];
      if (headers&&headers.length) {
        sh.appendRow(headers);
        sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a1a18').setFontColor('#ffffff');
        sh.setFrozenRows(1);
        SpreadsheetApp.flush();
      }
    }
  });
  ensureColumns();
  seedDefaultAdmin();
  seedDefaultConfig();
  seedDefaultAcctStatuses();
  seedDefaultRefData();
  // Cache that we've initialised so subsequent requests skip the heavy work
  try { PropertiesService.getScriptProperties().setProperty(_INIT_KEY, String(Date.now())); } catch(_) {}
}

/** Fast-path init — only runs the full ensureSheets() if the cache has expired (>24 h).
 *  Cuts ~300–600 ms off every data-fetch request once the system is set up. */
function ensureSheetsIfNeeded() {
  try {
    const ts = PropertiesService.getScriptProperties().getProperty(_INIT_KEY);
    if (ts && (Date.now() - Number(ts)) < INIT_CACHE_TTL_MS) return; // already initialised recently
  } catch(_) {}
  ensureSheets(); // first time or cache expired → run full init
}

function ensureColumns() {
  const ss=getSpreadsheet();
  Object.entries(SH).forEach(([key,name])=>{
    const sh=ss.getSheetByName(name); if(!sh) return;
    const canonical=HDR[key]; if(!canonical||!canonical.length) return;
    const lastCol=sh.getLastColumn();
    const existing=lastCol>0?sh.getRange(1,1,1,lastCol).getValues()[0].map(String):[];
    const missing=canonical.filter(col=>!existing.includes(col));
    if(!missing.length) return;
    missing.forEach(col=>{
      const nc=sh.getLastColumn()+1;
      sh.getRange(1,nc).setValue(col).setFontWeight('bold').setBackground('#1a1a18').setFontColor('#ffffff');
    });
    SpreadsheetApp.flush();
  });
}

function seedDefaultAdmin() {
  const sh=getSheet(SH.USERS); if(sh.getLastRow()>1) return;
  const ALL_PERMS=JSON.stringify(['view_registry','add_property','edit_property','delete_property','print_par','view_real_property','manage_real_property','view_reports','export_data','manage_config','view_audit_log','manage_users','manage_prs']);
  sh.appendRow([Date.now(),'Administrator','admin','Admin@1234','admin@ozamiz.gov.ph','CITY GENERAL SERVICES OFFICE','System Administrator','Admin','Active',ALL_PERMS,new Date().toISOString(),'']);
  SpreadsheetApp.flush();
}

function seedDefaultConfig() {
  const sh=getSheet(SH.CONFIG); if(sh.getLastRow()>1) return;
  [['city','City of Ozamiz'],['province','Misamis Occidental'],['fiscalYear',String(new Date().getFullYear())],['parThreshold','50000'],['semester','1st Semester '+new Date().getFullYear()],['appVersion','4.0']].forEach(r=>sh.appendRow(r));
  SpreadsheetApp.flush();
}

function seedDefaultAcctStatuses() {
  const sh=getSheet(SH.ACCT_STATUSES); if(sh.getLastRow()>1) return;
  const now=Date.now();
  [[now,'Payables','Obligation incurred; payment not yet made.'],[now+1,'Paid','Payment fully settled and liquidated.']].forEach(r=>sh.appendRow(r));
  SpreadsheetApp.flush();
}

function seedDefaultRefData() {
  var now=Date.now();
  var dSh=getSheet(SH.DEPARTMENTS);
  if (dSh.getLastRow()<=1) {
    ["CITY ACCOUNTING OFFICE","CITY GENERAL SERVICES OFFICE","OFFICE OF THE CITY MAYOR","CITY PLANNING AND DEVELOPMENT OFFICE","CITY TREASURER'S OFFICE","CITY CIVIL REGISTRAR OFFICE","CITY ECONOMIC AND DEVELOPMENT OFFICE","CITY ENGINEER'S OFFICE","CITY BUDGET OFFICE","SM LAO HOSPITAL","CITY HEALTH OFFICE","CITY LIBRARY","CITY ASSESSOR'S OFFICE","CITY COUNCIL OFFICE","CITY VETERINARY OFFICE","CITY ADMINISTRATOR'S OFFICE","CITY DISASTER RISK REDUCTION MANAGEMENT OFFICE","CITY HUMAN RESOURCE MANAGEMENT OFFICE","CITY NUTRITION OFFICE","CITY SOLID WASTE AND ENVIRONMENT MANAGEMENT OFFICE","CITY SOCIAL WELFARE AND DEVELOPMENT OFFICE","PERSONS WITH DISABILITIES AFFAIRS OFFICE","COMMISSION ON AUDIT","CITY INFORMATION CHANNEL","PHILIPPINE NATIONAL POLICE (PNP)","CITY AGRICULTURIST OFFICE","OZAMIZ CITY TECHNICAL AND VOCATIONAL SCHOOL","DEPARTMENT OF INTERIOR & LOCAL GOVERNMENT","DEPARTMENT OF EDUCATION","CITY ADMINISTRATOR'S OFFICE / PERMIT & LICENSE DIVISION","CITY TOURISM OFFICE","BAC OFFICE","LIGA NG BARANGAY"].forEach(function(name,i){dSh.appendRow([now+i,name,'','']);});
    SpreadsheetApp.flush();
  }
  var cSh=getSheet(SH.CONDITIONS);
  if (cSh.getLastRow()<=1) {
    [['Serviceable','green'],['Unserviceable','red'],['For Repair','amber'],['Disposed','purple'],['Issued','blue']].forEach(function(r,i){cSh.appendRow([now+1000+i,r[0],r[1]]);});
    SpreadsheetApp.flush();
  }
  var aSh=getSheet(SH.ACQUISITIONS);
  if (aSh.getLastRow()<=1) {
    ['Purchase','Donation','Transfer','Confiscation','Construction'].forEach(function(n,i){aSh.appendRow([now+2000+i,n,'']);});
    SpreadsheetApp.flush();
  }
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET HELPERS
// ═══════════════════════════════════════════════════════════════════
function cleanSpreadsheetId(raw) {
  if(!raw) return raw;
  const s=String(raw).trim(), m=s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if(m) return m[1];
  return s.split('/')[0].split('?')[0].split('#')[0];
}

// Cache the spreadsheet object for the duration of one script execution.
// GAS creates a new execution context per request, so this cache lives
// only as long as one HTTP request — safe and avoids redundant openById calls.
let _ssCache = null;
function getSpreadsheet() {
  if (_ssCache) return _ssCache;
  if(!SPREADSHEET_ID||SPREADSHEET_ID.length<20) throw new Error('SPREADSHEET_ID not configured.');
  _ssCache = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ssCache;
}
function getSheet(name) {
  const sh=getSpreadsheet().getSheetByName(name);
  if(!sh) throw new Error(`Sheet "${name}" not found — run ensureSheets() first.`);
  return sh;
}

/** Read a sheet using a pre-opened Spreadsheet object (avoids repeated openById). */
function _readSheetFromSS(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) { Logger.log('readSheet: "'+sheetName+'" not found'); return []; }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return [];
  const raw = sh.getRange(1, 1, lastRow, lastCol).getValues();
  // Trim and normalise header names so visual capitalization ("User") doesn't break the frontend JSON map
  const headers = raw[0].map(h => {
    let s = String(h).trim();
    if (sheetName === SH.AUDIT_LOG) {
      const ls = s.toLowerCase();
      if (['id', 'ts', 'action', 'property', 'user', 'details'].includes(ls)) return ls;
      // Handle alias names common in audit logs
      if (ls === 'timestamp' || ls === 'date/time' || ls === 'date') return 'ts';
      if (ls === 'module') return 'property';
    }
    return s;
  });
  const result = [];
  for (let ri = 1; ri < raw.length; ri++) {
    const row = raw[ri];
    // Skip fully blank rows
    let hasValue = false;
    for (let ci = 0; ci < row.length; ci++) {
      const c = row[ci];
      if (c !== '' && c !== null && c !== undefined) { hasValue = true; break; }
    }
    if (!hasValue) continue;
    const obj = {};
    for (let ci = 0; ci < headers.length; ci++) {
      let v = row[ci];
      if (v instanceof Date) {
        // Fast ISO date formatting without Utilities.formatDate (saves ~1ms per cell)
        const Y=v.getFullYear(), M=String(v.getMonth()+1).padStart(2,'0'),
              D=String(v.getDate()).padStart(2,'0'),
              h=String(v.getHours()).padStart(2,'0'), m=String(v.getMinutes()).padStart(2,'0'),
              s=String(v.getSeconds()).padStart(2,'0');
        v = (h==='00' && m==='00' && s==='00')
          ? `${Y}-${M}-${D}`
          : `${Y}-${M}-${D}T${h}:${m}:${s}`;
      } else if (typeof v === 'string' && v.length > 1 && (v[0]==='['||v[0]==='{')) {
        try { v = JSON.parse(v); } catch(_) {}
      }
      obj[headers[ci]] = (v === null || v === undefined) ? '' : v;
    }
    result.push(obj);
  }
  return result;
}

/** Public readSheet — uses the cached spreadsheet object. */
function readSheet(sheetName) {
  return _readSheetFromSS(getSpreadsheet(), sheetName);
}
const _colCache = {};
function _invalidateColCache(shName) { delete _colCache[shName]; }
function upsertRow(sheetName, obj, skipFlush) {
  const sh=getSheet(sheetName), lastCol=sh.getLastColumn();
  if(lastCol<1) return;
  const headers=sh.getRange(1,1,1,lastCol).getValues()[0];
  const idCol=headers.indexOf('id');

  // For the Users sheet: if editing an existing row and no password is supplied,
  // preserve the existing password rather than overwriting it with ''.
  let existingPwdValue = null;
  const isUsers = (sheetName === SH.USERS);
  const pwdColIdx = isUsers ? headers.indexOf('password') : -1;
  const pwdMissing = isUsers && pwdColIdx >= 0 &&
                     (obj.password === undefined || obj.password === null || obj.password === '');

  if (pwdMissing && idCol >= 0 && obj.id !== undefined && obj.id !== '') {
    const ri = findRow(sh, idCol+1, obj.id);
    if (ri > 0) {
      existingPwdValue = sh.getRange(ri, pwdColIdx+1).getValue();
    }
  }

  const row=headers.map(h=>{
    // Preserve existing password when none supplied on edit
    if (isUsers && h === 'password' && pwdMissing) {
      return existingPwdValue !== null ? existingPwdValue : '';
    }
    const v=obj[h];
    if(v===undefined||v===null) return '';
    if(typeof v==='object') return JSON.stringify(v);
    return v;
  });
  const isPPE=(sheetName===SH.PPE_MAIN||sheetName===SH.PPE_SUB);
  if(idCol>=0&&obj.id!==undefined&&obj.id!=='') {
    const ri=findRow(sh,idCol+1,obj.id);
    if(ri>0) {
      const rng=sh.getRange(ri,1,1,row.length);
      if(isPPE)rng.setNumberFormat('@');
      rng.setValues([row]);
      if(!skipFlush) SpreadsheetApp.flush();
      return;
    }
  }
  const nr=sh.getLastRow()+1, rng=sh.getRange(nr,1,1,row.length);
  if(isPPE) rng.setNumberFormat('@');
  rng.setValues([row]);
  if (_colCache[sheetName] && _colCache[sheetName][idCol+1]) {
    _colCache[sheetName][idCol+1].set(String(obj.id), nr);
  }
  if(!skipFlush) SpreadsheetApp.flush();
}
function findRow(sh,col,value) {
  const shName = sh.getName();
  if (!_colCache[shName]) _colCache[shName] = {};
  if (!_colCache[shName][col]) {
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return -1;
    const vals = sh.getRange(2, col, lastRow - 1, 1).getValues();
    const map = new Map();
    for (let i = 0; i < vals.length; i++) map.set(String(vals[i][0]).trim(), i + 2);
    _colCache[shName][col] = map;
  }
  const mapped = _colCache[shName][col].get(String(value).trim());
  return mapped ? mapped : -1;
}
function deleteById(sheetName,id,skipFlush) {
  const sh=getSheet(sheetName), lastCol=sh.getLastColumn(); if(lastCol<1) return false;
  const headers=sh.getRange(1,1,1,lastCol).getValues()[0];
  const idCol=headers.indexOf('id')+1; if(idCol<1) return false;
  const ri=findRow(sh,idCol,id); if(ri<1) return false;
  sh.deleteRow(ri); _invalidateColCache(sheetName);
  if(!skipFlush) SpreadsheetApp.flush(); return true;
}
function saveRow(sheetName,obj,sess,label) {
  requirePerm(sess,['manage_config']); if(!obj.id) obj.id=Date.now();
  upsertRow(sheetName,obj); audit('EDIT',label,'Saved '+label+': '+(obj.name||obj.id), sess.username); return {ok:true};
}
function deleteRecord(sheetName,id,sess,label) {
  requirePerm(sess,['manage_config']); deleteById(sheetName,id);
  audit('DELETE',label,'Deleted '+label+' id='+id, sess.username); return {ok:true};
}
function requirePerm(sess,perms) {
  if(!sess) throw new Error('Not authenticated.');
  if(sess.role==='Admin') return;
  if(!perms.some(p=>(sess.permissions||[]).includes(p))) throw new Error('Permission denied.');
}
function audit(action,property,details,user) {
  // Use the no-flush writer — UUIDs are never pre-existing so skipping findRow scan is safe
  try { _auditNoFlush(action, property||'', details||'', user||''); } catch(_) {}
}
function logErr(ctx,err) { try{Logger.log('[ERROR] '+ctx+': '+err.message);}catch(_){} }
function jsonOut(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function readConfig() {
  try { const rows=readSheet(SH.CONFIG),cfg={}; rows.forEach(r=>{if(r.key)cfg[r.key]=r.value;}); return cfg; }
  catch(_) { return {}; }
}
function saveConfig(payload,sess) {
  if (sess.role !== 'Admin') return {ok:false,error:'Settings are restricted to Admin users only.'};
  Object.entries(payload).forEach(([key,value])=>{ upsertRow(SH.CONFIG,{key,value}); });
  return {ok:true};
}

// ═══════════════════════════════════════════════════════════════════
//  AUTHENTICATION
// ═══════════════════════════════════════════════════════════════════
function handleLogin(payload) {
  const {username,password}=payload;
  if(!username||!password) return {ok:false,error:'Username and password are required.'};

  // Fast-path init — only runs full ensureSheets() if cache expired (>24h)
  // Previously this called ensureSheets() unconditionally, adding 5–15s every login
  ensureSheetsIfNeeded();

  const users=readSheet(SH.USERS);
  const user=users.find(u=>u.username===username&&u.status==='Active');
  if(!user) return {ok:false,error:'Invalid username/password or account is inactive.'};

  const storedPwd=String(user.password||user.passwordHash||'').trim();
  const looksHashed=/^[0-9a-f]{64}$/i.test(storedPwd)||/^[0-9a-f]{40}$/i.test(storedPwd)||/^\$2[aby]\$/.test(storedPwd);
  if(!looksHashed&&!verifyPwd(password,storedPwd))
    return {ok:false,error:'Invalid username/password or account is inactive.'};

  // Update lastLogin — do it with a single targeted range write (no findRow scan)
  try {
    const sh=getSheet(SH.USERS);
    const hdr=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const llc=hdr.indexOf('lastLogin')+1;
    const ri=findRow(sh,hdr.indexOf('id')+1,user.id);
    if(llc>0&&ri>0) sh.getRange(ri,llc).setValue(new Date().toISOString());
    // No SpreadsheetApp.flush() here — let GAS batch it, saves ~500ms
  } catch(_) {}

  const token=Utilities.getUuid();
  let _perms=user.permissions;
  if(!Array.isArray(_perms)) {
    const _raw=String(_perms||'').trim();
    if(_raw.startsWith('[')) { try{_perms=JSON.parse(_raw);}catch(_){_perms=[];} }
    else if(_raw.length>0) { _perms=_raw.split(',').map(s=>s.trim()).filter(Boolean); }
    else { _perms=[]; }
  }
  // Back-fill manage_prs for Admin/Manager roles whose stored permissions
  // predate the PRS module (no sheet re-write needed — session only).
  if ((user.role==='Admin'||user.role==='Manager') && !_perms.includes('manage_prs')) {
    _perms = [..._perms, 'manage_prs'];
  }
  _sessionSet(token,{userId:user.id,username:user.username,role:user.role,permissions:_perms,expires:Date.now()+SESSION_TTL_MS});

  // Write audit log without flush — non-blocking
  try { _auditNoFlush('LOGIN', user.username, '@' + user.username + ' signed in', user.username); } catch(_) {}

  // Bundle ref data in the login response so the frontend needs ZERO extra
  // round-trips to render the app after login. This eliminates the separate
  // getRefData fetch that was adding 3–8s of perceived load time.
  let refData = null;
  try {
    const ss = getSpreadsheet();
    const sheets = _readAllSheets([
      SH.OFFICERS, SH.DEPARTMENTS, SH.PPE_MAIN, SH.PPE_SUB,
      SH.FUND_CLUSTERS, SH.CONDITIONS, SH.ACQUISITIONS,
      SH.ACCT_STATUSES, SH.CONFIG,
    ]);
    const cfg = {};
    (sheets[SH.CONFIG]||[]).forEach(r=>{ if(r.key) cfg[r.key]=r.value; });
    refData = {
      officers         : sheets[SH.OFFICERS],
      departments      : sheets[SH.DEPARTMENTS],
      ppeMainCategories: sheets[SH.PPE_MAIN],
      ppeSubcategories : sheets[SH.PPE_SUB],
      fundClusters     : sheets[SH.FUND_CLUSTERS],
      conditions       : sheets[SH.CONDITIONS],
      acquisitionModes : sheets[SH.ACQUISITIONS],
      acctStatuses     : sheets[SH.ACCT_STATUSES],
      systemUsers      : users.map(stripPassword),
      config           : cfg,
    };
  } catch(_) { /* refData stays null — frontend will fetch separately */ }

  return {ok:true,token,user:stripPassword(user),mustChangePassword:looksHashed,refData};
}

/** Audit log write without flush — faster for non-critical writes like login. */
function _auditNoFlush(action, property, details, user) {
  try {
    const sh = getSheet(SH.AUDIT_LOG), lastCol = sh.getLastColumn();
    if (lastCol < 1) return;
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const row = headers.map(h => {
      const lh = String(h || '').toLowerCase().trim();
      if (lh === 'id')       return Utilities.getUuid();
      if (lh === 'ts')       return new Date().toISOString();
      if (lh === 'action')   return action   || '';
      if (lh === 'property') return property || '';
      if (lh === 'user')     return user     || '';
      if (lh === 'details')  return details  || '';
      return '';
    });
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    // Intentionally no SpreadsheetApp.flush() — GAS will batch-commit on return
  } catch (_) {}
}
function validateToken(token) {
  if(!token) return null;
  const s=_sessionGet(token); if(!s) return null;
  s.expires=Date.now()+SESSION_TTL_MS; _sessionSet(token,s); return s;
}
function handleLogout(token) {
  if (token) {
    const sess = _sessionGet(token);
    if (sess && sess.username) {
      try { _auditNoFlush('LOGOUT', sess.username, '@' + sess.username + ' signed out', sess.username); } catch(_) {}
    }
    _sessionDel(token);
  }
  return {ok:true};
}
function verifyPwd(plain,stored) { return String(plain||'').trim()===String(stored||'').trim(); }
function stripPassword(u) { const c=Object.assign({},u); delete c.password; delete c.passwordHash; return c; }

// ═══════════════════════════════════════════════════════════════════
//  FORGOT PASSWORD  — OTP FLOW
// ═══════════════════════════════════════════════════════════════════
function handleSendOTP(payload) {
  try {
    const identifier=String(payload.identifier||'').trim().toLowerCase();
    if(!identifier) return {ok:false,error:'Please enter your username or email address.'};
    const users=readSheet(SH.USERS);
    const user=users.find(u=>u.status==='Active'&&(u.username.toLowerCase()===identifier||(u.email&&u.email.toLowerCase()===identifier)));
    if(!user) return {ok:false,error:'No active account found with that username or email.'};
    if(!user.email) return {ok:false,error:'This account has no email address on file. Please contact your administrator.'};

    const otp=String(Math.floor(100000+Math.random()*900000));
    _setOTP(user.username,otp);

    MailApp.sendEmail({
      to     : user.email,
      subject: 'CGSO Password Reset OTP',
      body   : 'Hello '+user.name+',\n\nYour One-Time Password (OTP) for resetting your CGSO Property System account password is:\n\n    '+otp+'\n\nThis OTP expires in 10 minutes. Do not share it with anyone.\n\nIf you did not request this, please contact your administrator.\n\n— CGSO Property Management System',
    });

    const parts=user.email.split('@');
    const masked=parts[0].slice(0,2)+'***@'+parts[1];
    return {ok:true,maskedEmail:masked,username:user.username};
  } catch(err) { logErr('handleSendOTP',err); return {ok:false,error:'Failed to send OTP: '+err.message}; }
}

function handleVerifyOTP(payload) {
  try {
    const username=String(payload.username||'').trim();
    const otp=String(payload.otp||'').trim();
    if(!username||!otp) return {ok:false,error:'Username and OTP are required.'};
    const entry=_getOTP(username);
    if(!entry) return {ok:false,error:'OTP expired or not issued. Please request a new one.'};
    entry.attempts=(entry.attempts||0)+1;
    if(entry.attempts>5) { _deleteOTP(username); return {ok:false,error:'Too many incorrect attempts. Please request a new OTP.'}; }
    const map=_loadOTPs(); map[username]=entry; _saveOTPs(map);
    if(entry.otp!==otp) return {ok:false,error:'Incorrect OTP. Please check and try again. ('+( 5-entry.attempts+1)+' attempts remaining)'};
    entry.verified=true; map[username]=entry; _saveOTPs(map);
    return {ok:true};
  } catch(err) { logErr('handleVerifyOTP',err); return {ok:false,error:err.message}; }
}

function handleResetPassword(payload) {
  try {
    const username=String(payload.username||'').trim();
    const newPwd=String(payload.newPwd||'').trim();
    if(!username) return {ok:false,error:'Username is required.'};
    if(!newPwd||newPwd.length<8) return {ok:false,error:'New password must be at least 8 characters.'};
    const entry=_getOTP(username);
    if(!entry||!entry.verified) return {ok:false,error:'OTP not verified. Please complete verification first.'};
    const sh=getSheet(SH.USERS), lastCol=sh.getLastColumn();
    const headers=sh.getRange(1,1,1,lastCol).getValues()[0];
    let pwdCol=headers.indexOf('password'); if(pwdCol<0) pwdCol=headers.indexOf('passwordHash');
    const unCol=headers.indexOf('username');
    if(pwdCol<0||unCol<0) return {ok:false,error:'Users sheet structure error.'};
    const data=sh.getRange(2,1,sh.getLastRow()-1,lastCol).getValues();
    let targetRow=-1;
    for(let i=0;i<data.length;i++) { if(String(data[i][unCol]||'').trim()===username){targetRow=i+2;break;} }
    if(targetRow<0) return {ok:false,error:'User not found.'};
    sh.getRange(targetRow,pwdCol+1).setValue(newPwd);
    SpreadsheetApp.flush();
    _deleteOTP(username);
    audit('EDIT',username,'@'+username+' reset password via OTP', username);
    return {ok:true};
  } catch(err) { logErr('handleResetPassword',err); return {ok:false,error:err.message}; }
}

// ═══════════════════════════════════════════════════════════════════
//  PROPERTY CRUD  — PAR → PAR_Records, ICS → ICS_Records
// ═══════════════════════════════════════════════════════════════════
function _propSheet(doctype) {
  return String(doctype||'').toUpperCase()==='ICS' ? SH.ICS_RECORDS : SH.PAR_RECORDS;
}
function saveProperty(p,sess) {
  requirePerm(sess,['add_property','edit_property']);
  const targetSheet=_propSheet(p.doctype);
  // Single SS open — read all three sheets at once to avoid repeated openById calls
  const ss=getSpreadsheet();
  const targetRows =_readSheetFromSS(ss,targetSheet);
  const legacyRows =_readSheetFromSS(ss,SH.PROPERTIES);
  const existInTarget=targetRows.find(x=>String(x.id)===String(p.id));
  const existInLegacy=legacyRows.find(x=>String(x.id)===String(p.id));
  const existing=existInTarget||existInLegacy;
  const isEdit=!!existing;
  if(!p.id) p.id=Date.now();
  if(!p.propno) p.propno=_genPropNoFromRows(p.doctype||'PAR', targetRows);
  p.updatedAt=new Date().toISOString();
  if(!isEdit){p.createdAt=p.updatedAt;p.createdBy=sess.username;}
  else{p.createdAt=existing.createdAt;p.createdBy=existing.createdBy;}
  upsertRow(targetSheet,p);
  if(existInLegacy) { try{deleteById(SH.PROPERTIES,p.id);}catch(_){} }
  // Fast date string without Utilities.formatDate
  const today=(function(){const d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');})();
  if(!isEdit&&p.custodian) {
    upsertRow(SH.HISTORY,{id:Utilities.getUuid(),propId:p.id,type:'initial',officer:p.custodian,designation:p.designation||'',dept:p.dept||'',date:p.date||today,location:p.location||'',remarks:'Initial issuance / first assignment',recordedBy:sess.username,recordedAt:new Date().toISOString()});
  }
  if(isEdit&&existing.custodian&&existing.custodian!==p.custodian&&p.custodian) {
    upsertRow(SH.HISTORY,{id:Utilities.getUuid(),propId:p.id,type:'transfer',officer:p.custodian,designation:p.designation||'',dept:p.dept||'',date:p.date||today,location:p.location||'',remarks:'Custodian updated via record edit',recordedBy:sess.username,recordedAt:new Date().toISOString()});
  }
  // Build detailed audit message with field-level diff for edits
  let auditDetails = '';
  if (isEdit && existing) {
    const TRACKED_FIELDS = [
      {key:'name', label:'Name'}, {key:'custodian', label:'Custodian'},
      {key:'dept', label:'Department'}, {key:'status', label:'Condition'},
      {key:'cost', label:'Unit Cost'}, {key:'location', label:'Location'},
      {key:'doctype', label:'Document Type'}, {key:'propno', label:'Property No.'},
      {key:'category', label:'Category'}, {key:'brand', label:'Brand'},
      {key:'serial', label:'Serial/Model'}, {key:'designation', label:'Designation'},
      {key:'fundcluster', label:'Fund Cluster'}, {key:'remarks', label:'Remarks'},
    ];
    const diffs = [];
    TRACKED_FIELDS.forEach(f => {
      const oldVal = String(existing[f.key] || '').trim();
      const newVal = String(p[f.key] || '').trim();
      if (oldVal !== newVal) diffs.push(f.label + ': "' + oldVal + '" → "' + newVal + '"');
    });
    auditDetails = 'Prop No: ' + p.propno + (diffs.length ? ' | Changes: ' + diffs.join('; ') : ' (no tracked field changes)');
  } else {
    auditDetails = 'Registered as ' + p.doctype + ' — ' + p.propno;
  }
  const actingUser = (sess && sess.username) ? sess.username : '';
  audit(isEdit ? 'EDIT' : 'ADD', p.name, auditDetails, actingUser);
  return {ok:true, record:p};
}
function deleteProperty(id, sess) {
  requirePerm(sess, ['delete_property']);
  // Read property details BEFORE deleting so we can log the name and propno
  const ss = getSpreadsheet();
  let propName = 'Property', propNo = String(id);
  try {
    const allProps = [
      ..._readSheetFromSS(ss, SH.PAR_RECORDS),
      ..._readSheetFromSS(ss, SH.ICS_RECORDS),
      ..._readSheetFromSS(ss, SH.PROPERTIES),
    ];
    const found = allProps.find(x => String(x.id) === String(id));
    if (found) {
      propName = found.name || 'Property';
      propNo   = found.propno || String(id);
    }
  } catch(_) {}
  const d1 = deleteById(SH.PAR_RECORDS, id);
  const d2 = deleteById(SH.ICS_RECORDS, id);
  const d3 = deleteById(SH.PROPERTIES,  id);
  if (!d1 && !d2 && !d3) return {ok:false, error:'Property not found.'};
  const actingUser = (sess && sess.username) ? sess.username : '';
  audit('DELETE', propName, 'Deleted Property No: ' + propNo, actingUser);
  return {ok:true};
}

// ═══════════════════════════════════════════════════════════════════
//  PRS CRUD
//  savePRS  — moves property from PAR_Records / ICS_Records into
//             PRS_Records and logs the transfer.
//  deletePRS — hard-deletes a PRS record (admin/manager only).
// ═══════════════════════════════════════════════════════════════════
function savePRS(payload, sess) {
  // Accept manage_prs (new) OR delete_property (existing Manager sessions
  // created before manage_prs was added — prevents "Permission denied" on
  // sessions that predate the permission).  Admin bypasses all checks.
  requirePerm(sess, ['manage_prs', 'delete_property']);
  if (!payload || !payload.id) return {ok:false, error:'Invalid PRS payload.'};

  // Guarantee PRS_Records sheet exists regardless of the init cache.
  // We bypass ensureSheetsIfNeeded() because the cache may be stale from
  // a deployment that pre-dated PRS_Records.
  const _ss = getSpreadsheet();
  if (!_ss.getSheetByName(SH.PRS_RECORDS)) {
    const _sh = _ss.insertSheet(SH.PRS_RECORDS);
    const _hdr = HDR.PRS_RECORDS;
    _sh.appendRow(_hdr);
    _sh.getRange(1,1,1,_hdr.length).setFontWeight('bold').setBackground('#1a1a18').setFontColor('#ffffff');
    _sh.setFrozenRows(1);
    SpreadsheetApp.flush();
    // Invalidate the init cache so ensureSheets runs fully on next request
    try { PropertiesService.getScriptProperties().deleteProperty(_INIT_KEY); } catch(_) {}
  }

  // Stamp server-side fields
  payload.prsProcessedBy = sess.username;
  payload.prsProcessedAt = new Date().toISOString();
  if (!payload.prsNo) {
    const existing = readSheet(SH.PRS_RECORDS);
    const yr = new Date().getFullYear().toString().slice(-2);
    payload.prsNo = `PRS-${String(existing.length + 1).padStart(5,'0')}-${yr}`;
  }

  // Write to PRS_Records; flush commits it before we remove from origin.
  upsertRow(SH.PRS_RECORDS, payload);
  SpreadsheetApp.flush();

  // Remove from origin sheet (PAR or ICS)
  const originSheet = String(payload.doctype||'').toUpperCase() === 'ICS'
    ? SH.ICS_RECORDS : SH.PAR_RECORDS;
  try { deleteById(originSheet, payload.id); } catch(delErr) { logErr('savePRS:del', delErr); }
  // Also remove from legacy Properties sheet if it existed there
  try { deleteById(SH.PROPERTIES, payload.id); } catch(_) {}

  // Write history entry
  try {
    upsertRow(SH.HISTORY, {
      id: Utilities.getUuid(),
      propId: payload.id,
      type: 'prs',
      officer: payload.prsReceivedBy || sess.username,
      designation: payload.prsReceivedByPos || '',
      dept: 'CGSO',
      date: payload.prsDate || new Date().toISOString().split('T')[0],
      location: '',
      remarks: `PRS ${payload.prsNo} — ${payload.prsPurpose}`,
      recordedBy: sess.username,
      recordedAt: new Date().toISOString(),
    });
  } catch(_) {}

  // Final flush: commits history + audit writes before HTTP response returns
  try { SpreadsheetApp.flush(); } catch(_) {}

  audit('PRS', payload.name,
    `PRS ${payload.prsNo}: ${payload.prsPurpose} — ${payload.propno} (${payload.doctype})`, sess.username);
  return {ok:true, prsNo: payload.prsNo, id: String(payload.id)};
}

function deletePRS(id, sess) {
  requirePerm(sess, ['manage_prs', 'delete_property']);
  const deleted = deleteById(SH.PRS_RECORDS, id);
  if (!deleted) return {ok:false, error:'PRS record not found.'};
  audit('DELETE', 'PRS Record', 'Deleted PRS record id=' + id, sess.username);
  return {ok:true};
}

function transferPRS(payload, sess) {
  requirePerm(sess, ['manage_prs']);
  if (!payload.propId || !payload.officer || !payload.dept) return {ok:false, error:'Missing transfer details.'};

  const ss = getSpreadsheet();
  const prsList = _readSheetFromSS(ss, SH.PRS_RECORDS);
  const p = prsList.find(x => String(x.id) === String(payload.propId));
  if (!p) return {ok:false, error:'PRS record not found.'};

  const originSheet = String(p.doctype||'').toUpperCase() === 'ICS' ? SH.ICS_RECORDS : SH.PAR_RECORDS;

  p.custodian = payload.officer;
  p.designation = payload.designation || '';
  p.dept = payload.dept;
  if (payload.location) p.location = payload.location;
  p.status = 'Serviceable'; // Re-issued as serviceable
  p.updatedAt = new Date().toISOString();

  // Strip PRS fields
  delete p.prsPurpose; delete p.prsDate; delete p.prsReceivedBy;
  delete p.prsReceivedByPos; delete p.prsRemarks; delete p.prsNo;
  delete p.prsProcessedBy; delete p.prsProcessedAt;

  upsertRow(originSheet, p);
  try { deleteById(SH.PRS_RECORDS, p.id); } catch(e) {}

  try {
    upsertRow(SH.HISTORY, {
      id: payload.entryId || Utilities.getUuid(),
      propId: p.id,
      type: 'transfer',
      officer: payload.officer,
      designation: payload.designation || '',
      dept: payload.dept,
      date: payload.date || new Date().toISOString().split('T')[0],
      location: payload.location || '',
      remarks: payload.remarks || 'Re-issued from PRS',
      recordedBy: sess.username,
      recordedAt: new Date().toISOString(),
    });
  } catch(e) {}

  audit('TRANSFER', p.name, `Transferred from PRS: ${p.propno} → ${p.custodian}`, sess.username);
  return {ok:true, property: p};
}
function genPropNo(doctype) {
  return _genPropNoFromRows(doctype, readSheet(_propSheet(doctype)));
}
/** Generate propno from an already-loaded rows array — avoids re-reading the sheet. */
function _genPropNoFromRows(doctype, rows) {
  const prefix=doctype==='ICS'?'ICS':'PAR', yr=new Date().getFullYear().toString().slice(-2);
  const count=rows.filter(p=>(p.propno||'').toUpperCase().startsWith(prefix+'-')).length;
  return `${prefix}-${String(count+1).padStart(5,'0')}-${yr}`;
}

// ═══════════════════════════════════════════════════════════════════
//  REAL PROPERTY CRUD
// ═══════════════════════════════════════════════════════════════════
function saveRealProperty(r,sess) {
  requirePerm(sess,['manage_real_property']);
  const allRP=readSheet(SH.REAL_PROP);
  const existing=allRP.find(x=>String(x.id)===String(r.id)), isEdit=!!existing;
  if(!r.id) r.id=Date.now();
  if(!r.recno) {
    const yr=new Date().getFullYear().toString().slice(-2);
    r.recno=`RP-${String(allRP.length+1).padStart(5,'0')}-${yr}`;
  }
  r.updatedAt=new Date().toISOString();
  if(!isEdit){r.createdAt=r.updatedAt;r.createdBy=sess.username;}
  else{r.createdAt=existing.createdAt;r.createdBy=existing.createdBy;}
  upsertRow(SH.REAL_PROP,r);
  audit(isEdit?'EDIT':'ADD',r.name,`${isEdit?'Updated':'Registered'} real property ${r.recno}`, sess.username);
  return {ok:true,record:r};
}

// ═══════════════════════════════════════════════════════════════════
//  ASSET HISTORY
// ═══════════════════════════════════════════════════════════════════
function saveHistory(entry,sess) {
  requirePerm(sess,['add_property','edit_property']);
  if(!entry.id) entry.id=Utilities.getUuid();
  entry.recordedBy=sess.username; entry.recordedAt=new Date().toISOString();
  upsertRow(SH.HISTORY,entry);
  if(entry.type==='transfer'&&entry.propId) {
    // Single SS open for all 3 lookups
    const ss=getSpreadsheet();
    const par=_readSheetFromSS(ss,SH.PAR_RECORDS).find(p=>String(p.id)===String(entry.propId));
    const ics=par?null:_readSheetFromSS(ss,SH.ICS_RECORDS).find(p=>String(p.id)===String(entry.propId));
    const leg=(par||ics)?null:_readSheetFromSS(ss,SH.PROPERTIES).find(p=>String(p.id)===String(entry.propId));
    const prop=par||ics||leg;
    if(prop) {
      prop.custodian=entry.officer;prop.designation=entry.designation;
      prop.dept=entry.dept;if(entry.location)prop.location=entry.location;
      prop.updatedAt=new Date().toISOString();
      upsertRow(par?SH.PAR_RECORDS:ics?SH.ICS_RECORDS:SH.PROPERTIES,prop);
      audit('TRANSFER',prop.name,'Transferred to @'+entry.officer+' ('+entry.dept+')', sess.username);
    }
  }
  return {ok:true};
}
function getHistory(propId) {
  if(!propId) return {ok:false,error:'propId required.'};
  return {ok:true,data:readSheet(SH.HISTORY).filter(h=>String(h.propId)===String(propId))};
}

// ═══════════════════════════════════════════════════════════════════
//  PDF UPLOAD / DELETE
// ═══════════════════════════════════════════════════════════════════
function _findProp(id) {
  // Single SS open — short-circuit once found so we never read more sheets than needed
  const ss=getSpreadsheet();
  const par=_readSheetFromSS(ss,SH.PAR_RECORDS).find(p=>String(p.id)===String(id));
  if(par) return {prop:par, sh:SH.PAR_RECORDS};
  const ics=_readSheetFromSS(ss,SH.ICS_RECORDS).find(p=>String(p.id)===String(id));
  if(ics) return {prop:ics, sh:SH.ICS_RECORDS};
  const rp=_readSheetFromSS(ss,SH.REAL_PROP).find(p=>String(p.id)===String(id));
  if(rp) return {prop:rp, sh:SH.REAL_PROP};
  const leg=_readSheetFromSS(ss,SH.PROPERTIES).find(p=>String(p.id)===String(id));
  return {prop:leg||null, sh:SH.PROPERTIES};
}
function uploadPDF(payload,sess) {
  requirePerm(sess,['add_property','edit_property']);
  const propId = payload.propId || payload.recordId;
  const fileName = payload.fileName || payload.filename || 'document.pdf';
  if(!propId) return {ok:false,error:'propId required.'};
  if(!payload.base64) return {ok:false,error:'base64 data required.'};
  const bytes=Utilities.base64Decode(payload.base64.replace(/^data:[^;]+;base64,/,''));
  if(bytes.length>MAX_UPLOAD_BYTES) return {ok:false,error:'File too large (max 10 MB).'};
  const folder=DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const blob=Utilities.newBlob(bytes,'application/pdf',fileName);
  const file=folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  const link='https://drive.google.com/file/d/'+file.getId()+'/view';
  const {prop,sh}=_findProp(propId);
  if(prop) { prop.pdfLink=link; prop.pdfName=fileName; upsertRow(sh,prop); }
  return {ok:true,url:link,link,filename:fileName,fileName};
}
function deletePDF(payload,sess) {
  requirePerm(sess,['add_property','edit_property']);
  const propId = payload.propId || payload.recordId;
  const {prop,sh}=_findProp(propId);
  if(!prop) return {ok:false,error:'Property not found.'};
  prop.pdfLink=''; prop.pdfName=''; upsertRow(sh,prop);
  return {ok:true};
}

// ═══════════════════════════════════════════════════════════════════
//  USER MANAGEMENT  — temp password + email
// ═══════════════════════════════════════════════════════════════════
function generateTempPassword() {
  const chars='ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  let pwd='';
  for(let i=0;i<8;i++) pwd+=chars[Math.floor(Math.random()*chars.length)];
  return pwd+'3!';   // always ends with digit+special to meet complexity
}

function saveUser(u,sess) {
  requirePerm(sess,['manage_users']);
  const existing=readSheet(SH.USERS).find(x=>String(x.id)===String(u.id));
  const isEdit=!!existing;
  let tempPassword=null;

  if(!isEdit) {
    if(!u.password) { tempPassword=generateTempPassword(); u.password=tempPassword; }
    else { tempPassword=u.password; }
    if(u.password.length<8) return {ok:false,error:'Password must be at least 8 characters.'};
  } else {
    if(u.password&&u.password.length<8) return {ok:false,error:'Password must be at least 8 characters.'};
    if(!u.password) delete u.password;
  }

  const all=readSheet(SH.USERS);
  if(all.some(x=>x.username===u.username&&String(x.id)!==String(u.id)))
    return {ok:false,error:`Username "@${u.username}" is already taken.`};
  if(!u.id) u.id=Date.now();
  u.createdAt=isEdit?existing.createdAt:new Date().toISOString();
  u.lastLogin=u.lastLogin||(existing?existing.lastLogin:'');
  if(Array.isArray(u.permissions)) u.permissions=JSON.stringify(u.permissions);
  upsertRow(SH.USERS,u);
  audit(isEdit?'EDIT':'ADD',u.name,(isEdit?'Updated':'Created')+' @'+u.username+' ('+u.role+')', sess.username);

  if(!isEdit&&u.email) {
    try {
      MailApp.sendEmail({
        to:u.email, subject:'Your CGSO Property System Account',
        body:'Hello '+u.name+',\n\nYour account has been created on the CGSO Property & Asset Management System.\n\nUsername          : '+u.username+'\nTemporary Password: '+tempPassword+'\n\nPlease log in and change your password immediately via My Profile → Update Password.\n\nSystem URL: '+ScriptApp.getService().getUrl()+'\n\n— CGSO Property Management System',
      });
    } catch(mailErr) { logErr('saveUser email',mailErr); }
  }
  return {ok:true,user:stripPassword(u),tempPassword:!isEdit?tempPassword:null};
}

function deleteUser(id,sess) {
  requirePerm(sess,['manage_users']);
  const users=readSheet(SH.USERS), target=users.find(x=>String(x.id)===String(id));
  if(!target) return {ok:false,error:'User not found.'};
  if(target.role==='Admin'&&users.filter(u=>u.role==='Admin').length<=1)
    return {ok:false,error:'Cannot delete the only Admin.'};
  deleteById(SH.USERS,id);
  audit('DELETE',target.name,'Deleted @'+target.username, sess.username);
  return {ok:true};
}

// ═══════════════════════════════════════════════════════════════════
//  CHANGE PASSWORD
// ═══════════════════════════════════════════════════════════════════
function doChangePassword(payload,token) {
  try {
    const username=String(payload.username||'').trim();
    const currentPwd=String(payload.currentPwd||'').trim();
    const newPwd=String(payload.newPwd||'').trim();
    if(!username) return {ok:false,error:'Username is required.'};
    if(!newPwd||newPwd.length<8) return {ok:false,error:'New password must be at least 8 characters.'};
    const sh=getSheet(SH.USERS), lastCol=sh.getLastColumn();
    const headers=sh.getRange(1,1,1,lastCol).getValues()[0];
    let pwdCol=headers.indexOf('password'); if(pwdCol<0) pwdCol=headers.indexOf('passwordHash');
    const unCol=headers.indexOf('username'), stCol=headers.indexOf('status');
    if(pwdCol<0) return {ok:false,error:'No password column found. Headers: '+headers.filter(Boolean).join(', ')};
    if(unCol<0) return {ok:false,error:'Column "username" not found.'};
    const data=sh.getRange(2,1,sh.getLastRow()-1,lastCol).getValues();
    let targetRow=-1;
    for(let i=0;i<data.length;i++){
      const ru=String(data[i][unCol]||'').trim(), rs=stCol>=0?String(data[i][stCol]||'').trim():'Active';
      if(ru===username&&rs==='Active'){targetRow=i+2;break;}
    }
    if(targetRow<0) return {ok:false,error:'User not found or inactive.'};
    const storedPwd=String(data[targetRow-2][pwdCol]||'').trim();
    const looksHashed=/^[0-9a-f]{64}$/i.test(storedPwd)||/^[0-9a-f]{40}$/i.test(storedPwd)||/^\$2[aby]\$/.test(storedPwd);
    let adminOverride=false;
    try{const s=_sessionGet(token);if(s&&s.role==='Admin')adminOverride=true;}catch(_){}
    if(!adminOverride){
      if(!currentPwd) return {ok:false,error:'Current password is required.'};
      if(!looksHashed&&currentPwd!==storedPwd) return {ok:false,error:'Current password is incorrect.'};
    }
    sh.getRange(targetRow,pwdCol+1).setValue(newPwd);
    SpreadsheetApp.flush();
    audit('EDIT',username,'@'+username+' changed password', username);
    return {ok:true};
  } catch(err){logErr('doChangePassword',err);return {ok:false,error:err.message};}
}

// ═══════════════════════════════════════════════════════════════════
//  MIGRATE HASHED PASSWORDS
// ═══════════════════════════════════════════════════════════════════
const MIGRATED_DEFAULT_PASSWORD='ChangeMe@1234';
function migrateHashedPasswords() {
  try {
    const sh=getSheet(SH.USERS), lastCol=sh.getLastColumn(), lastRow=sh.getLastRow();
    if(lastRow<2) return {ok:true,migrated:0,message:'No users.'};
    const headers=sh.getRange(1,1,1,lastCol).getValues()[0];
    let pwdCol=headers.indexOf('password'); if(pwdCol<0) pwdCol=headers.indexOf('passwordHash');
    const unCol=headers.indexOf('username');
    if(pwdCol<0) return {ok:false,error:'No password column.'};
    const isH=pwd=>{const s=String(pwd||'').trim();return /^[0-9a-f]{64}$/i.test(s)||/^[0-9a-f]{40}$/i.test(s)||/^\$2[aby]\$/.test(s);};
    const data=sh.getRange(2,1,lastRow-1,lastCol).getValues();
    let migrated=0; const affected=[];
    for(let i=0;i<data.length;i++){if(!isH(data[i][pwdCol]))continue;sh.getRange(i+2,pwdCol+1).setValue(MIGRATED_DEFAULT_PASSWORD);migrated++;if(unCol>=0)affected.push(String(data[i][unCol]||''));}
    if(migrated>0) SpreadsheetApp.flush();
    return {ok:true,migrated,affected,message:migrated>0?`Migrated ${migrated}. Temp password: ${MIGRATED_DEFAULT_PASSWORD}`:'No hashed passwords found.'};
  } catch(err){logErr('migrateHashedPasswords',err);return {ok:false,error:err.message};}
}

// ═══════════════════════════════════════════════════════════════════
//  BULK IMPORT
// ═══════════════════════════════════════════════════════════════════
function bulkImport(payload,sess) {
  requirePerm(sess,['add_property']);
  if(sess.role !== 'Admin' && sess.role !== 'Manager') return {ok:false,error:'Bulk import is restricted to Admin and Manager users.'};
  const {records,type}=payload;
  if(!Array.isArray(records)||!records.length) return {ok:false,error:'No records provided.'};
  let imported=0;

  if(type==='realProperty'){
    const existing=readSheet(SH.REAL_PROP);
    // Build set of existing recnos for duplicate check
    const usedRecnos=new Set(existing.map(r=>String(r.recno||'').toUpperCase()));
    const yr=new Date().getFullYear().toString().slice(-2);
    let rpCounter=existing.length;
    records.forEach(r=>{
      if(!r.id) r.id=Date.now()+imported;
      // Ensure unique recno
      if(!r.recno){
        do { rpCounter++; r.recno=`RP-${String(rpCounter).padStart(5,'0')}-${yr}`; }
        while(usedRecnos.has(r.recno.toUpperCase()));
      }
      usedRecnos.add(String(r.recno).toUpperCase());
      r.createdAt=r.createdAt||new Date().toISOString();
      r.createdBy=r.createdBy||sess.username;
      r.updatedAt=new Date().toISOString();
      upsertRow(SH.REAL_PROP,r,true); imported++;
    });
  } else {
    // Build a unified set of ALL existing propnos (PAR + ICS + legacy)
    const parRows=readSheet(SH.PAR_RECORDS);
    const icsRows=readSheet(SH.ICS_RECORDS);
    const legRows=readSheet(SH.PROPERTIES);
    const usedPropnos=new Set([...parRows,...icsRows,...legRows].map(p=>String(p.propno||'').toUpperCase()));
    const yr=new Date().getFullYear().toString().slice(-2);
    // Per-doctype counters for auto-generation (start from current count)
    const counters={PAR:parRows.filter(p=>(p.propno||'').toUpperCase().startsWith('PAR-')).length,
                    ICS:icsRows.filter(p=>(p.propno||'').toUpperCase().startsWith('ICS-')).length};
    records.forEach(p=>{
      if(!p.id) p.id=Date.now()+imported;
      const doctype=(p.doctype||'PAR').toUpperCase();
      const prefix=doctype==='ICS'?'ICS':'PAR';
      if(!p.propno){
        // Auto-generate unique propno
        let candidate;
        do { counters[prefix]=(counters[prefix]||0)+1;
             candidate=`${prefix}-${String(counters[prefix]).padStart(5,'0')}-${yr}`; }
        while(usedPropnos.has(candidate.toUpperCase()));
        p.propno=candidate;
      } else {
        // Provided propno — check for duplicate; if taken, auto-generate a new one
        if(usedPropnos.has(String(p.propno).toUpperCase())){
          let candidate;
          do { counters[prefix]=(counters[prefix]||0)+1;
               candidate=`${prefix}-${String(counters[prefix]).padStart(5,'0')}-${yr}`; }
          while(usedPropnos.has(candidate.toUpperCase()));
          p.propno=candidate;
        }
      }
      usedPropnos.add(String(p.propno).toUpperCase());
      p.createdAt=p.createdAt||new Date().toISOString();
      p.createdBy=p.createdBy||sess.username;
      p.updatedAt=new Date().toISOString();
      upsertRow(_propSheet(p.doctype),p,true); imported++;
    });
  }
  if(imported > 0) SpreadsheetApp.flush();
  audit('IMPORT','Bulk Import','Imported '+imported+' '+(type||'property')+' record(s) by @'+sess.username, sess.username);
  return {ok:true,imported};
}

function importInitialData() {
  ensureSheets();
  /* Paste exported JSON below and run manually from Apps Script editor.
  const data = { ... };
  */
}

// ═══════════════════════════════════════════════════════════════════
//  MIGRATE LEGACY PROPERTIES → PAR_Records + ICS_Records
//  Run once via: <url>?action=migrateProperties
//  Safe to run multiple times — only moves records not already in the
//  target sheets. Does NOT delete from the legacy Properties sheet
//  (manual cleanup after verifying migration was successful).
// ═══════════════════════════════════════════════════════════════════
function migratePropertiesToSeparateSheets() {
  try {
    ensureSheets();
    const all  = readSheet(SH.PROPERTIES);
    const pars = readSheet(SH.PAR_RECORDS);
    const icss = readSheet(SH.ICS_RECORDS);
    const parIds = new Set(pars.map(r=>String(r.id)));
    const icsIds = new Set(icss.map(r=>String(r.id)));
    let movedPAR=0, movedICS=0, skipped=0;
    all.forEach(p=>{
      const id=String(p.id);
      if(String(p.doctype||'').toUpperCase()==='ICS'){
        if(icsIds.has(id)){skipped++;return;}
        upsertRow(SH.ICS_RECORDS,p,true); movedICS++;
      } else {
        if(parIds.has(id)){skipped++;return;}
        upsertRow(SH.PAR_RECORDS,p,true); movedPAR++;
      }
    });
    if((movedPAR + movedICS) > 0) SpreadsheetApp.flush();
    return {ok:true,movedPAR,movedICS,skipped,
      message:`Migrated ${movedPAR} PAR + ${movedICS} ICS records (${skipped} already in target sheets). `+
              'Legacy Properties sheet NOT deleted — verify data then clear it manually.'};
  } catch(err){logErr('migrateProperties',err);return {ok:false,error:err.message};}
}