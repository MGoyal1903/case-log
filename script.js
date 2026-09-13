const STORAGE_KEY = 'fsl_cases_v1';
const MONTH_ORDER = ["January","February","March","April","May","June","July","August","September","October","November","December"];



let cases = [];
let deleteTargetId = null;
let sheetsUrl = null;
let sheetsEnabled = false;
let pollTimer = null;
const SHEETS_URL_KEY = 'sheets_webapp_url_v1';

/* This app runs as a downloaded file opened directly in your browser, not inside Claude's
   chat, so on-device settings (PIN, Sheets URL) use the browser's own localStorage — this
   only works once the file is actually opened on a device, not previewed inside a chat. */
function lsGet(key){
  try{ return localStorage.getItem(key); }
  catch(e){ console.error('localStorage get error', e); return null; }
}
function lsSet(key, value){
  try{ localStorage.setItem(key, value); return true; }
  catch(e){ console.error('localStorage set error', e); return false; }
}
function lsDelete(key){
  try{ localStorage.removeItem(key); }catch(e){ console.error('localStorage delete error', e); }
}

function uid(){ return 'c_' + Math.random().toString(36).slice(2,10) + Date.now().toString(36); }

function parseDate(str){
  // expects dd.mm.yyyy ; returns Date or null
  if(!str) return null;
  const m = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if(!m) return null;
  return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
}

function monthFromDate(str){
  const d = parseDate(str);
  if(!d) return '';
  return MONTH_ORDER[d.getMonth()];
}

function showToast(msg){
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(()=> t.classList.remove('show'), 2200);
}

function rerenderVisible(){
  const activePage = document.querySelector('.page.active');
  if(!activePage) return;
  if(activePage.id === 'page-dashboard') renderDashboard();
  if(activePage.id === 'page-cases') renderCasesTable();
}

/* ---------- Local (Claude-session) fallback storage ---------- */
async function loadCasesLocal(){
  const val = lsGet(STORAGE_KEY);
  if(val){
    try{ cases = JSON.parse(val); return; }catch(e){ /* corrupt value, fall through to reseed */ }
  }
  cases = SEED_CASES.map(c => ({ id: uid(), ...c }));
  await saveCasesLocal();
}

async function saveCasesLocal(){
  const ok = lsSet(STORAGE_KEY, JSON.stringify(cases));
  if(!ok) showToast('Could not save on this device — check browser storage isn\'t blocked');
}

/* ---------- Google Sheets (user's own database) ---------- */
const APPS_SCRIPT_SOURCE = `function doGet(e) {
  if (e.parameter.action === 'downloadDoc') {
    var fileId = e.parameter.fileId;
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var out = {
      fileName: file.getName(),
      mimeType: blob.getContentType(),
      base64Data: Utilities.base64Encode(blob.getBytes())
    };
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    rows.push(obj);
  }
  return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  if (body.action === 'add') {
    sheet.appendRow(headers.map(function(h){ return body.data[h] !== undefined ? body.data[h] : ''; }));
  } else if (body.action === 'update') {
    var idCol = headers.indexOf('id');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(body.id)) {
        var rowNum = i + 1;
        headers.forEach(function(h, colIdx){
          if (body.data[h] !== undefined) sheet.getRange(rowNum, colIdx + 1).setValue(body.data[h]);
        });
        break;
      }
    }
  } else if (body.action === 'delete') {
    var idColD = headers.indexOf('id');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idColD]) === String(body.id)) { sheet.deleteRow(i + 1); break; }
    }
  } else if (body.action === 'seed') {
    body.rows.forEach(function(r){
      sheet.appendRow(headers.map(function(h){ return r[h] !== undefined ? r[h] : ''; }));
    });
  } else if (body.action === 'uploadDoc') {
    var idColU = headers.indexOf('id');
    var docsColU = headers.indexOf('documents');
    var fslColU = headers.indexOf('fslNo');
    var firColU = headers.indexOf('firNo');
    var rowIndexU = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idColU]) === String(body.caseId)) { rowIndexU = i; break; }
    }
    if (rowIndexU === -1) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Case not found' })).setMimeType(ContentService.MimeType.JSON);
    }
    var caseLabel = data[rowIndexU][fslColU] || data[rowIndexU][firColU] || ('Case-' + body.caseId);
    var caseFolder = getCaseFolder_(caseLabel);
    var bytes = Utilities.base64Decode(body.base64Data);
    var blob = Utilities.newBlob(bytes, body.mimeType, body.fileName);
    var file = caseFolder.createFile(blob);
    var rowNumU = rowIndexU + 1;
    var docs;
    try { docs = JSON.parse(data[rowIndexU][docsColU] || '{}'); } catch(err) { docs = {}; }
    var old = docs[body.docType];
    if (old && old.fileId) {
      try { DriveApp.getFileById(old.fileId).setTrashed(true); } catch(err2) {}
    }
    docs[body.docType] = { fileId: file.getId(), name: body.fileName, mimeType: body.mimeType, uploadedAt: new Date().toISOString() };
    sheet.getRange(rowNumU, docsColU + 1).setValue(JSON.stringify(docs));
    return ContentService.createTextOutput(JSON.stringify({ ok: true, fileId: file.getId() })).setMimeType(ContentService.MimeType.JSON);
  } else if (body.action === 'deleteDoc') {
    var idColR = headers.indexOf('id');
    var docsColR = headers.indexOf('documents');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idColR]) === String(body.caseId)) {
        var rowNumR = i + 1;
        var docsR;
        try { docsR = JSON.parse(data[i][docsColR] || '{}'); } catch(err) { docsR = {}; }
        if (docsR[body.docType] && docsR[body.docType].fileId) {
          try { DriveApp.getFileById(docsR[body.docType].fileId).setTrashed(true); } catch(err2) {}
        }
        delete docsR[body.docType];
        sheet.getRange(rowNumR, docsColR + 1).setValue(JSON.stringify(docsR));
        break;
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
}

function getDocsRootFolder_() {
  var folders = DriveApp.getFoldersByName('CaseLogDocuments');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('CaseLogDocuments');
}

function getCaseFolder_(label) {
  var safeName = String(label).replace(/[\\\/:*?"<>|]/g, '-').trim();
  if (!safeName) safeName = 'Untitled Case';
  var root = getDocsRootFolder_();
  var existing = root.getFoldersByName(safeName);
  if (existing.hasNext()) return existing.next();
  return root.createFolder(safeName);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var baseHeaders = ['id','fslNo','firNo','policeStation','exhibitCapacity','exhibits','extractionFrom','remarks1','us','letterType','courtDoReceived','dateGiven','openingDate','reportedDate','collectedDate','status','remarks2','month','documents'];
  var sheet = ss.getSheetByName('Cases');
  if (!sheet) {
    sheet = ss.insertSheet('Cases');
    sheet.appendRow(baseHeaders);
    return sheet;
  }
  var lastCol = sheet.getLastColumn();
  var existingHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (existingHeaders.indexOf('documents') === -1) {
    sheet.getRange(1, existingHeaders.length + 1).setValue('documents');
  }
  return sheet;
}`;

function setSyncStatus(mode){
  // mode: 'checking' | 'local' | 'connected' | 'error'
  const el = document.getElementById('syncStatus');
  const banner = document.getElementById('dbBanner');
  const map = {
    checking: { color:'var(--amber)', text:'Checking database…' },
    local:    { color:'var(--amber)', text:'Not connected — using session only' },
    connected:{ color:'var(--green)', text:'Connected to your Google Sheet' },
    error:    { color:'var(--red)',   text:'Connection error — check Database tab' },
  };
  const s = map[mode] || map.local;
  el.innerHTML = `<span style="width:6px;height:6px;border-radius:50%;background:${s.color};flex-shrink:0;"></span><span>${s.text}</span>`;
  if(banner) banner.style.display = (mode === 'local' || mode === 'error') ? 'flex' : 'none';

  const connectedPanel = document.getElementById('dbConnectedPanel');
  const setupPanel = document.getElementById('dbSetupPanel');
  if(connectedPanel && setupPanel){
    if(mode === 'connected'){
      connectedPanel.style.display = 'block';
      setupPanel.style.display = 'none';
    } else {
      connectedPanel.style.display = 'none';
      setupPanel.style.display = 'block';
    }
  }
}

async function fetchSheetCases(url){
  const res = await fetch(url, { method: 'GET' });
  if(!res.ok) throw new Error('Bad response: ' + res.status);
  return res.json();
}

async function sheetsPost(url, payload){
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload)
  });
  if(!res.ok) throw new Error('Bad response: ' + res.status);
  return res.json();
}

async function refreshFromSheet(){
  const rows = await fetchSheetCases(sheetsUrl);
  cases = rows.filter(r => r && (r.id || r.fslNo || r.firNo)).map(r => ({ ...r, id: r.id || uid() }));
  rerenderVisible();
  renderDashboard();
}

function startPolling(){
  if(pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(()=>{ refreshFromSheet().catch(err=>{ console.error('Sheet poll error', err); setSyncStatus('error'); }); }, 15000);
}

async function connectSheetsUrl(url, { migrate }){
  const existing = await fetchSheetCases(url); // throws if the URL/deployment is wrong
  if(migrate && existing.length === 0){
    const migrateSource = cases.length ? cases : SEED_CASES.map(c => ({ id: uid(), ...c }));
    await sheetsPost(url, { action: 'seed', rows: migrateSource });
  }
  sheetsUrl = url;
  sheetsEnabled = true;
  await refreshFromSheet();
  startPolling();
  setSyncStatus('connected');
}

async function tryAutoConnectSheets(){
  try{
    const val = lsGet(SHEETS_URL_KEY);
    if(val){
      await connectSheetsUrl(val, { migrate: false });
      return true;
    }
  }catch(e){ console.error('Auto-connect error', e); }
  return false;
}

document.getElementById('connectDbBtn').addEventListener('click', async ()=>{
  const statusEl = document.getElementById('dbConnectStatus');
  const url = document.getElementById('f_sheetsUrl').value.trim();
  if(!url){ statusEl.textContent = 'Paste your Web app URL first.'; return; }
  statusEl.textContent = 'Connecting…';
  try{
    await connectSheetsUrl(url, { migrate: true });
    lsSet(SHEETS_URL_KEY, url);
    statusEl.textContent = '';
    showToast('Sheet connected');
  }catch(e){
    console.error(e);
    statusEl.textContent = 'Could not connect — double-check the deployment is set to "Anyone" and try again.';
  }
});

document.getElementById('disconnectDbBtn').addEventListener('click', async ()=>{
  if(pollTimer) clearInterval(pollTimer);
  sheetsEnabled = false;
  sheetsUrl = null;
  try{ lsDelete(SHEETS_URL_KEY); }catch(e){}
  await loadCasesLocal();
  setSyncStatus('local');
  rerenderVisible();
  renderDashboard();
  showToast('Disconnected — back to session-only storage');
});

document.getElementById('refreshSheetBtn').addEventListener('click', async ()=>{
  try{
    await refreshFromSheet();
    showToast('Refreshed');
  }catch(e){
    setSyncStatus('error');
    showToast('Could not refresh — check your connection');
  }
});

document.getElementById('dbBannerLink').addEventListener('click', (e)=>{
  e.preventDefault();
  document.querySelector('.nav-item[data-page="db"]').click();
});

document.getElementById('copyScriptBtn').addEventListener('click', async ()=>{
  const textarea = document.getElementById('appsScriptCode');
  try{
    await navigator.clipboard.writeText(APPS_SCRIPT_SOURCE);
    showToast('Script copied');
  }catch(e){
    textarea.removeAttribute('readonly');
    textarea.focus();
    textarea.select();
    try{ document.execCommand('copy'); showToast('Script copied'); }
    catch(e2){ showToast('Select the text above and copy manually'); }
    textarea.setAttribute('readonly', 'true');
  }
});
document.getElementById('appsScriptCode').value = APPS_SCRIPT_SOURCE;

/* Unified read/write used by the rest of the app */
async function loadCases(){
  setSyncStatus('checking');
  const connected = await tryAutoConnectSheets();
  if(!connected){
    await loadCasesLocal();
    setSyncStatus('local');
  }
}

async function persistNewCase(data){
  if(sheetsEnabled){
    const id = uid();
    await sheetsPost(sheetsUrl, { action: 'add', data: { id, ...data } });
    await refreshFromSheet();
  } else {
    cases.push({ id: uid(), ...data });
    await saveCasesLocal();
  }
}

async function persistUpdateCase(id, data){
  if(sheetsEnabled){
    await sheetsPost(sheetsUrl, { action: 'update', id, data });
    await refreshFromSheet();
  } else {
    const idx = cases.findIndex(c=>c.id===id);
    if(idx > -1) cases[idx] = { ...cases[idx], ...data };
    await saveCasesLocal();
  }
}

async function persistDeleteCase(id){
  if(sheetsEnabled){
    await sheetsPost(sheetsUrl, { action: 'delete', id });
    await refreshFromSheet();
  } else {
    cases = cases.filter(c=>c.id !== id);
    await saveCasesLocal();
  }
}

/* ---------- CSV export ---------- */
const CSV_COLUMNS = [
  ['fslNo','FSL No.'], ['firNo','FIR No.'], ['policeStation','Police Station'],
  ['us','U/S'], ['exhibitCapacity','Exhibit Capacity'], ['exhibits','Exhibits'],
  ['extractionFrom','Extraction From'], ['letterType','Letter Type'],
  ['courtDoReceived','Court/DO Received'], ['dateGiven','Date Given'],
  ['openingDate','Opening Date'], ['month','Month'], ['reportedDate','Reported Date'],
  ['collectedDate','Collected Date'], ['status','Status'], ['remarks1','Remarks']
];

function toCsvValue(v){
  const s = (v === undefined || v === null) ? '' : String(v);
  if(/[",\n]/.test(s)) return '"' + s.replace(/"/g,'""') + '"';
  return s;
}

function downloadCurrentFilterAsCsv(){
  const list = getFilteredCases();
  if(list.length === 0){ showToast('No cases match this filter'); return; }
  const header = CSV_COLUMNS.map(c=>toCsvValue(c[1])).join(',');
  const rows = list.map(c => CSV_COLUMNS.map(col => toCsvValue(c[col[0]])).join(','));
  const csv = [header, ...rows].join('\r\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const statusF = document.getElementById('statusFilter').value || 'all';
  const monthF = document.getElementById('monthFilter').value || 'all-months';
  const a = document.createElement('a');
  a.href = url;
  a.download = `cases_${statusF}_${monthF}.csv`.replace(/\s+/g,'-');
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  showToast(`Downloaded ${list.length} case${list.length===1?'':'s'}`);
}

document.getElementById('downloadCsvBtn').addEventListener('click', downloadCurrentFilterAsCsv);

/* ---------- Navigation ---------- */
document.querySelectorAll('.nav-item').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('.nav-item').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    const page = btn.dataset.page;
    document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
    document.getElementById('page-'+page).classList.add('active');
    if(page === 'dashboard') renderDashboard();
    if(page === 'cases') renderCasesTable();
  });
});


document.getElementById('addCaseBtn').addEventListener('click', ()=>{
  resetForm();
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.remove('active'));
  document.querySelector('.nav-item[data-page="add"]').classList.add('active');
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.getElementById('page-add').classList.add('active');
});

/* ---------- Dashboard ---------- */
function renderDashboard(){
  const total = cases.length;
  const byStatus = { Reported:0, Collected:0, Hold:0, Pending:0 };
  cases.forEach(c=>{
    const s = c.status && byStatus.hasOwnProperty(c.status) ? c.status : 'Pending';
    byStatus[s]++;
  });
  document.getElementById('statTotal').textContent = total;
  document.getElementById('statReported').textContent = byStatus.Reported;
  document.getElementById('statCollected').textContent = byStatus.Collected;
  document.getElementById('statHold').textContent = byStatus.Hold;
  document.getElementById('statPending').textContent = byStatus.Pending;
  document.getElementById('footCount').textContent = total + (total===1?' case on record':' cases on record');

  // monthly bars
  const monthBuckets = {};
  MONTH_ORDER.forEach(m => monthBuckets[m] = { Reported:0, Collected:0, Hold:0, Pending:0 });
  cases.forEach(c=>{
    const m = MONTH_ORDER.includes(c.month) ? c.month : null;
    if(!m) return;
    const s = c.status && monthBuckets[m].hasOwnProperty(c.status) ? c.status : 'Pending';
    monthBuckets[m][s]++;
  });
  const activeMonths = MONTH_ORDER.filter(m => Object.values(monthBuckets[m]).some(v=>v>0));
  const maxTotal = Math.max(1, ...activeMonths.map(m => Object.values(monthBuckets[m]).reduce((a,b)=>a+b,0)));

  const colors = { Reported:'var(--blue)', Collected:'var(--green)', Hold:'var(--slate)', Pending:'var(--amber)' };
  let html = '';
  if(activeMonths.length === 0){
    html = '<div class="empty-state">No dated cases yet.</div>';
  }
  activeMonths.forEach(m=>{
    const b = monthBuckets[m];
    const t = b.Reported+b.Collected+b.Hold+b.Pending;
    const widthPct = (t / maxTotal) * 100;
    html += `<div class="bar-row">
      <div class="bar-label">${m}</div>
      <div class="bar-track"><div style="display:flex; width:${widthPct}%;">
        ${b.Reported ? `<div class="bar-seg" style="background:${colors.Reported}; width:${b.Reported/t*100}%"></div>`:''}
        ${b.Collected ? `<div class="bar-seg" style="background:${colors.Collected}; width:${b.Collected/t*100}%"></div>`:''}
        ${b.Hold ? `<div class="bar-seg" style="background:${colors.Hold}; width:${b.Hold/t*100}%"></div>`:''}
        ${b.Pending ? `<div class="bar-seg" style="background:${colors.Pending}; width:${b.Pending/t*100}%"></div>`:''}
      </div></div>
      <div class="bar-total">${t} case${t===1?'':'s'}</div>
    </div>`;
  });
  document.getElementById('monthlyBars').innerHTML = html;

  // recent table (top 6 by opening date desc)
  const recent = [...cases].sort((a,b)=>{
    const da = parseDate(a.openingDate), db = parseDate(b.openingDate);
    if(!da && !db) return 0;
    if(!da) return 1;
    if(!db) return -1;
    return db - da;
  }).slice(0,6);
  document.getElementById('recentTableWrap').innerHTML = buildTable(recent, false);
  attachRowHandlers(document.getElementById('recentTableWrap'));
}

/* ---------- Case Log table ---------- */
function buildTable(list, withActions){
  if(list.length === 0){
    return '<div class="empty-state">No cases match. Try clearing filters, or add your first case.</div>';
  }
  let rows = list.map(c=>{
    const status = c.status || 'Pending';
    return `<tr data-id="${c.id}">
      <td class="mono">${escapeHtml(c.fslNo)}</td>
      <td class="mono">${escapeHtml(c.firNo)}</td>
      <td>${escapeHtml(c.policeStation)}</td>
      <td class="dim">${escapeHtml(c.us)}</td>
      <td class="dim">${escapeHtml(c.openingDate)}</td>
      <td><span class="status-pill status-${status}">${status}</span></td>
      <td class="dim">${escapeHtml(c.month)}</td>
      ${withActions ? `<td><div class="row-actions">
        <button class="icon-btn view-btn" title="View details">👁</button>
        <button class="icon-btn edit-btn">Edit</button>
        <button class="icon-btn danger delete-btn">Delete</button>
      </div></td>` : ''}
    </tr>`;
  }).join('');
  return `<table>
    <thead><tr>
      <th>FSL No.</th><th>FIR No.</th><th>Police station</th><th>U/S</th>
      <th>Opened</th><th>Status</th><th>Month</th>${withActions?'<th></th>':''}
    </tr></thead>
    <tbody>${rows}</tbody>
  </table>`;
}

function escapeHtml(s){
  if(s === undefined || s === null) return '';
  return String(s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
}

function attachRowHandlers(container){
  container.querySelectorAll('.view-btn').forEach(btn=>{
    btn.addEventListener('click', (e)=>{
      const id = e.target.closest('tr').dataset.id;
      openView(id);
    });
  });
  container.querySelectorAll('.edit-btn').forEach(btn=>{
    btn.addEventListener('click', (e)=>{
      const id = e.target.closest('tr').dataset.id;
      openEdit(id);
    });
  });
  container.querySelectorAll('.delete-btn').forEach(btn=>{
    btn.addEventListener('click', (e)=>{
      deleteTargetId = e.target.closest('tr').dataset.id;
      document.getElementById('deleteModal').classList.add('open');
    });
  });
}

function populateMonthFilter(){
  const sel = document.getElementById('monthFilter');
  const current = sel.value;
  const months = Array.from(new Set(cases.map(c=>c.month).filter(m=>MONTH_ORDER.includes(m))));
  months.sort((a,b)=> MONTH_ORDER.indexOf(a) - MONTH_ORDER.indexOf(b));
  sel.innerHTML = '<option value="">All months</option>' + months.map(m=>`<option value="${m}">${m}</option>`).join('');
  sel.value = current;
}

function getFilteredCases(){
  const q = document.getElementById('searchInput').value.trim().toLowerCase();
  const statusF = document.getElementById('statusFilter').value;
  const monthF = document.getElementById('monthFilter').value;

  let filtered = cases.filter(c=>{
    if(statusF && (c.status || 'Pending') !== statusF) return false;
    if(monthF && c.month !== monthF) return false;
    if(q){
      const hay = [c.fslNo,c.firNo,c.policeStation,c.us,c.exhibits,c.remarks1].join(' ').toLowerCase();
      if(!hay.includes(q)) return false;
    }
    return true;
  });

  filtered.sort((a,b)=>{
    const da = parseDate(a.openingDate), dbb = parseDate(b.openingDate);
    if(!da && !dbb) return 0;
    if(!da) return 1;
    if(!dbb) return -1;
    return dbb - da;
  });
  return filtered;
}

function renderCasesTable(){
  populateMonthFilter();
  const filtered = getFilteredCases();
  const wrap = document.getElementById('casesTableWrap');
  wrap.innerHTML = buildTable(filtered, true);
  attachRowHandlers(wrap);
}

document.getElementById('searchInput').addEventListener('input', renderCasesTable);
document.getElementById('statusFilter').addEventListener('change', renderCasesTable);
document.getElementById('monthFilter').addEventListener('change', renderCasesTable);

/* ---------- Delete modal ---------- */
document.getElementById('cancelDeleteBtn').addEventListener('click', ()=>{
  deleteTargetId = null;
  document.getElementById('deleteModal').classList.remove('open');
});
document.getElementById('confirmDeleteBtn').addEventListener('click', async ()=>{
  if(!deleteTargetId) return;
  await persistDeleteCase(deleteTargetId);
  deleteTargetId = null;
  document.getElementById('deleteModal').classList.remove('open');
  renderCasesTable();
  renderDashboard();
  showToast('Case deleted');
});

/* ---------- Form (add / edit) ---------- */
const form = document.getElementById('caseForm');
const FIELD_MAP = ['fslNo','firNo','policeStation','us','exhibitCapacity','exhibits','extractionFrom',
  'letterType','courtDoReceived','dateGiven','openingDate','month','reportedDate','collectedDate','status','remarks1'];

function resetForm(){
  document.getElementById('editId').value = '';
  FIELD_MAP.forEach(f=>{
    const el = document.getElementById('f_'+f);
    if(el) el.value = (f === 'status') ? 'Pending' : '';
  });
  document.getElementById('formTitle').textContent = 'Add case';
  document.getElementById('formSub').textContent = "Enter the case details below — they're saved to your case log automatically.";
  document.getElementById('saveBtn').textContent = 'Save case';
  document.getElementById('cancelEditBtn').style.display = 'none';
}

function openEdit(id){
  const c = cases.find(x=>x.id===id);
  if(!c) return;
  document.getElementById('editId').value = id;
  FIELD_MAP.forEach(f=>{
    const el = document.getElementById('f_'+f);
    if(el) el.value = c[f] || (f==='status' ? 'Pending' : '');
  });
  document.getElementById('formTitle').textContent = 'Edit case';
  document.getElementById('formSub').textContent = c.fslNo || 'Update this case\'s details.';
  document.getElementById('saveBtn').textContent = 'Save changes';
  document.getElementById('cancelEditBtn').style.display = 'inline-block';

  document.querySelectorAll('.nav-item').forEach(b=>b.classList.remove('active'));
  document.querySelector('.nav-item[data-page="add"]').classList.add('active');
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.getElementById('page-add').classList.add('active');
}

/* ---------- View case + documents ---------- */
const DOC_TYPES = [
  { key: 'examinationReport', label: 'Examination Report' },
  { key: 'intimation', label: 'Intimation' },
  { key: 'forwardingLetter', label: 'Forwarding Letter' },
  { key: 'examinationSheet', label: 'Examination Sheet' },
];
const DETAIL_FIELDS = [
  ['fslNo','FSL No.'], ['firNo','FIR No.'], ['policeStation','Police Station'],
  ['us','U/S'], ['exhibitCapacity','Exhibit Capacity'], ['exhibits','Exhibits'],
  ['extractionFrom','Extraction From'], ['letterType','Letter Type'],
  ['courtDoReceived','Court/DO Received'], ['dateGiven','Date Given'],
  ['openingDate','Opening Date'], ['month','Month'], ['reportedDate','Reported Date'],
  ['collectedDate','Collected Date'], ['status','Status'], ['remarks1','Remarks'],
];
let viewingCaseId = null;

function parseDocsField(c){
  try{ return c.documents ? JSON.parse(c.documents) : {}; }
  catch(e){ return {}; }
}

function fileToBase64(file){
  return new Promise((resolve, reject)=>{
    const reader = new FileReader();
    reader.onload = ()=> resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function openView(id){
  const c = cases.find(x=>x.id===id);
  if(!c) return;
  viewingCaseId = id;
  document.getElementById('viewTitle').textContent = c.fslNo || c.firNo || 'Case details';
  document.getElementById('viewDetailGrid').innerHTML = DETAIL_FIELDS.map(([key,label])=>{
    if(key === 'status'){
      const status = c.status || 'Pending';
      return `<div class="detail-item"><span class="dl">${label}</span><span class="dv"><span class="status-pill status-${status}">${status}</span></span></div>`;
    }
    return `<div class="detail-item${key==='remarks1'?' full':''}"><span class="dl">${label}</span><span class="dv">${escapeHtml(c[key]) || '—'}</span></div>`;
  }).join('');

  document.getElementById('viewDocsGateNote').style.display = sheetsEnabled ? 'none' : 'block';
  renderDocsList(c);
  document.getElementById('viewModal').classList.add('open');
}

function renderDocsList(c){
  const docs = parseDocsField(c);
  const wrap = document.getElementById('docsList');
  wrap.innerHTML = DOC_TYPES.map(dt=>{
    const doc = docs[dt.key];
    return `<div class="doc-row" data-doctype="${dt.key}">
      <div class="doc-info">
        <div class="doc-label">${dt.label}</div>
        ${doc ? `<div class="doc-filename">${escapeHtml(doc.name)}</div>` : `<div class="doc-empty">No file uploaded</div>`}
      </div>
      <div class="doc-actions">
        <span class="doc-status"></span>
        ${doc ? `<button class="icon-btn download-doc-btn">Download</button>` : ''}
        <label class="icon-btn" style="cursor:pointer;">${doc ? 'Replace' : 'Upload'}<input type="file" class="file-input upload-doc-input"></label>
        ${doc ? `<button class="icon-btn danger remove-doc-btn">Remove</button>` : ''}
      </div>
    </div>`;
  }).join('');

  wrap.querySelectorAll('.doc-row').forEach(row=>{
    const docType = row.dataset.doctype;
    const statusEl = row.querySelector('.doc-status');

    const fileInput = row.querySelector('.upload-doc-input');
    if(fileInput){
      fileInput.addEventListener('change', async (e)=>{
        const file = e.target.files[0];
        if(!file) return;
        if(!sheetsEnabled){
          showToast('Connect a Google Sheet first (Database tab)');
          fileInput.value = '';
          return;
        }
        if(file.size > 10 * 1024 * 1024){
          showToast('Please keep files under 10MB');
          fileInput.value = '';
          return;
        }
        statusEl.textContent = 'Uploading…';
        try{
          const base64Data = await fileToBase64(file);
          await sheetsPost(sheetsUrl, {
            action: 'uploadDoc',
            caseId: viewingCaseId,
            docType,
            fileName: file.name,
            mimeType: file.type || 'application/octet-stream',
            base64Data
          });
          await refreshFromSheet();
          const updated = cases.find(x=>x.id===viewingCaseId);
          if(updated) renderDocsList(updated);
          showToast('Document uploaded');
        }catch(err){
          console.error(err);
          statusEl.textContent = '';
          showToast('Upload failed — please try again');
        }
      });
    }

    const downloadBtn = row.querySelector('.download-doc-btn');
    if(downloadBtn){
      downloadBtn.addEventListener('click', async ()=>{
        const c = cases.find(x=>x.id===viewingCaseId);
        const docs = parseDocsField(c);
        const doc = docs[docType];
        if(!doc) return;
        statusEl.textContent = 'Preparing…';
        try{
          const url = `${sheetsUrl}?action=downloadDoc&fileId=${encodeURIComponent(doc.fileId)}`;
          const res = await fetch(url);
          const json = await res.json();
          const byteChars = atob(json.base64Data);
          const bytes = new Uint8Array(byteChars.length);
          for(let i=0;i<byteChars.length;i++) bytes[i] = byteChars.charCodeAt(i);
          const blob = new Blob([bytes], { type: json.mimeType || 'application/octet-stream' });
          const blobUrl = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = blobUrl;
          a.download = json.fileName || doc.name;
          document.body.appendChild(a);
          a.click();
          a.remove();
          URL.revokeObjectURL(blobUrl);
          statusEl.textContent = '';
        }catch(err){
          console.error(err);
          statusEl.textContent = '';
          showToast('Download failed — please try again');
        }
      });
    }

    const removeBtn = row.querySelector('.remove-doc-btn');
    if(removeBtn){
      removeBtn.addEventListener('click', async ()=>{
        statusEl.textContent = 'Removing…';
        try{
          await sheetsPost(sheetsUrl, { action: 'deleteDoc', caseId: viewingCaseId, docType });
          await refreshFromSheet();
          const updated = cases.find(x=>x.id===viewingCaseId);
          if(updated) renderDocsList(updated);
          showToast('Document removed');
        }catch(err){
          console.error(err);
          statusEl.textContent = '';
          showToast('Could not remove — please try again');
        }
      });
    }
  });
}

document.getElementById('closeViewBtn').addEventListener('click', ()=>{
  document.getElementById('viewModal').classList.remove('open');
  viewingCaseId = null;
});

document.getElementById('cancelEditBtn').addEventListener('click', resetForm);

// auto-suggest month from opening date if month field is empty
document.getElementById('f_openingDate').addEventListener('blur', ()=>{
  const monthEl = document.getElementById('f_month');
  if(!monthEl.value){
    const suggested = monthFromDate(document.getElementById('f_openingDate').value);
    if(suggested) monthEl.value = suggested;
  }
});

form.addEventListener('submit', async (e)=>{
  e.preventDefault();
  const editId = document.getElementById('editId').value;
  const data = {};
  FIELD_MAP.forEach(f=>{
    const el = document.getElementById('f_'+f);
    data[f] = el ? el.value.trim() : '';
  });
  if(!data.status) data.status = 'Pending';
  if(!data.month) data.month = monthFromDate(data.openingDate) || '';

  if(!data.fslNo && !data.firNo){
    showToast('Add at least an FSL No. or FIR No.');
    return;
  }

  if(editId){
    await persistUpdateCase(editId, data);
    showToast('Case updated');
  } else {
    await persistNewCase(data);
    showToast('Case saved');
  }
  resetForm();

  document.querySelectorAll('.nav-item').forEach(b=>b.classList.remove('active'));
  document.querySelector('.nav-item[data-page="cases"]').classList.add('active');
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.getElementById('page-cases').classList.add('active');
  renderCasesTable();
});

/* ---------- PIN lock ---------- */
const PIN_HASH_KEY = 'pin_hash_v1';
let pinMode = 'checking'; // 'checking' | 'setup' | 'unlock' | 'unlocked'
let pinAttempts = 0;

async function sha256Hex(str){
  const salted = 'caselog-salt:' + str;
  try{
    const enc = new TextEncoder().encode(salted);
    const buf = await crypto.subtle.digest('SHA-256', enc);
    return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
  }catch(e){
    // fallback for contexts without Web Crypto (e.g. non-secure preview) — not cryptographically strong,
    // but this lock is a basic gate, not high-security, so a deterministic fallback is acceptable.
    let h = 0;
    for(let i=0;i<salted.length;i++){ h = ((h<<5)-h + salted.charCodeAt(i)) | 0; }
    return 'fb' + Math.abs(h).toString(16);
  }
}

function showLockOverlay(){ document.getElementById('lockOverlay').classList.remove('hidden'); }
function hideLockOverlay(){ document.getElementById('lockOverlay').classList.add('hidden'); }

function renderLockScreen(){
  const title = document.getElementById('lockTitle');
  const sub = document.getElementById('lockSub');
  const pinInput = document.getElementById('pinInput');
  const confirmInput = document.getElementById('pinConfirmInput');
  const submitBtn = document.getElementById('pinSubmitBtn');
  const hint = document.getElementById('lockHint');
  const err = document.getElementById('lockError');
  err.textContent = '';
  pinInput.value = '';
  confirmInput.value = '';
  pinInput.style.display = 'block';
  submitBtn.style.display = 'block';

  if(pinMode === 'setup'){
    title.textContent = 'Set up a PIN';
    sub.textContent = 'Choose a 4–8 digit PIN to lock this app.';
    pinInput.placeholder = 'New PIN';
    confirmInput.style.display = 'block';
    submitBtn.textContent = 'Set PIN';
    hint.textContent = "This only gates this app on this device — it isn't recoverable if forgotten, so pick something you'll remember.";
    pinInput.focus();
  } else if(pinMode === 'unlock'){
    title.textContent = 'Enter PIN';
    sub.textContent = 'Unlock your case tracker.';
    pinInput.placeholder = '····';
    confirmInput.style.display = 'none';
    submitBtn.textContent = 'Unlock';
    hint.textContent = '';
    pinInput.focus();
  }
}

async function handlePinSubmit(){
  const err = document.getElementById('lockError');
  const pin = document.getElementById('pinInput').value.trim();
  if(!/^\d{4,8}$/.test(pin)){
    err.textContent = 'PIN must be 4–8 digits.';
    return;
  }
  if(pinMode === 'setup'){
    const confirmPin = document.getElementById('pinConfirmInput').value.trim();
    if(pin !== confirmPin){
      err.textContent = "PINs don't match.";
      return;
    }
    const hash = await sha256Hex(pin);
    const ok = lsSet(PIN_HASH_KEY, hash);
    if(!ok){
      err.textContent = 'Could not save PIN on this device — check your browser allows local storage.';
      return;
    }
    pinMode = 'unlocked';
    hideLockOverlay();
    startApp();
  } else if(pinMode === 'unlock'){
    const stored = lsGet(PIN_HASH_KEY);
    const hash = await sha256Hex(pin);
    if(stored && hash === stored){
      pinMode = 'unlocked';
      err.textContent = '';
      hideLockOverlay();
      startApp();
    } else {
      pinAttempts++;
      err.textContent = pinAttempts >= 5 ? 'Too many attempts — wait a moment and try again.' : 'Incorrect PIN.';
      document.getElementById('pinInput').value = '';
      if(pinAttempts >= 5){
        document.getElementById('pinSubmitBtn').disabled = true;
        setTimeout(()=>{
          pinAttempts = 0;
          document.getElementById('pinSubmitBtn').disabled = false;
          err.textContent = '';
        }, 15000);
      }
    }
  }
}

document.getElementById('pinSubmitBtn').addEventListener('click', handlePinSubmit);
['pinInput','pinConfirmInput'].forEach(id=>{
  document.getElementById(id).addEventListener('keydown', (e)=>{ if(e.key === 'Enter') handlePinSubmit(); });
});

document.getElementById('lockNowBtn').addEventListener('click', ()=>{
  pinMode = 'unlock';
  pinAttempts = 0;
  renderLockScreen();
  showLockOverlay();
});

document.getElementById('changePinBtn').addEventListener('click', async ()=>{
  const status = document.getElementById('pinChangeStatus');
  const current = document.getElementById('f_currentPin').value.trim();
  const next = document.getElementById('f_newPin').value.trim();
  const confirm = document.getElementById('f_newPinConfirm').value.trim();
  if(!/^\d{4,8}$/.test(next) || next !== confirm){
    status.textContent = "New PIN must be 4–8 digits and match its confirmation.";
    return;
  }
  try{
    const stored = lsGet(PIN_HASH_KEY);
    const currentHash = await sha256Hex(current);
    if(stored && currentHash !== stored){
      status.textContent = 'Current PIN is incorrect.';
      return;
    }
    const newHash = await sha256Hex(next);
    lsSet(PIN_HASH_KEY, newHash);
    status.textContent = 'PIN updated.';
    document.getElementById('f_currentPin').value = '';
    document.getElementById('f_newPin').value = '';
    document.getElementById('f_newPinConfirm').value = '';
    showToast('PIN updated');
  }catch(e){
    status.textContent = 'Could not update PIN — please try again.';
  }
});

document.getElementById('removePinBtn').addEventListener('click', async ()=>{
  const current = document.getElementById('f_currentPin').value.trim();
  const status = document.getElementById('pinChangeStatus');
  try{
    const stored = lsGet(PIN_HASH_KEY);
    const currentHash = await sha256Hex(current);
    if(stored && currentHash !== stored){
      status.textContent = 'Enter your current PIN above first to turn off the lock.';
      return;
    }
    lsDelete(PIN_HASH_KEY);
    status.textContent = 'PIN lock turned off.';
    showToast('PIN lock removed');
  }catch(e){
    status.textContent = 'Could not remove PIN — please try again.';
  }
});

async function initPinGate(){
  showLockOverlay();
  const val = lsGet(PIN_HASH_KEY);
  pinMode = val ? 'unlock' : 'setup';
  renderLockScreen();
}

async function startApp(){
  await loadCases();
  renderDashboard();
  renderCasesTable();
}

/* ---------- Init ---------- */
(async function init(){
  await initPinGate();
})();
