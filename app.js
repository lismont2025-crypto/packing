/* ============================================================
   Fabric Inspection Report System — application logic
   Everything below is a client-side app: state lives in memory,
   is auto-saved to localStorage, and is exported to real .xlsx
   (SheetJS) and real text-based .pdf (jsPDF) — never screenshots.
   ============================================================ */

const LS_KEYS = { reports:'fir_reports', templates:'fir_templates', current:'fir_current', formulas:'fir_formulas', settings:'fir_settings' };

const DEFAULT_ROLL_COLUMNS = [
  {key:'sno',    label:'S.No',        type:'text',   width:'3%',  editable:false, auto:'index'},
  {key:'dcNo',   label:'Dc No.',      type:'text',   width:'6%'},
  {key:'lotNo',  label:'Lot No.',     type:'text',   width:'6%'},
  {key:'rollNo', label:'Roll No',     type:'text',   width:'7%'},
  {key:'fWidth', label:'Fabric F.Width (")', type:'number', width:'6%'},
  {key:'cWidth', label:'Fabric C.Width (")', type:'number', width:'6%', editable:false, auto:'formula:cwidth'},
  {key:'mtrs',   label:'MTRS',        type:'number', width:'5%'},
  {key:'p1',     label:'Pt 1',        type:'number', width:'4%'},
  {key:'p2',     label:'Pt 2',        type:'number', width:'4%'},
  {key:'p3',     label:'Pt 3',        type:'number', width:'4%'},
  {key:'p4',     label:'Pt 4',        type:'number', width:'4%'},
  {key:'totalPoints', label:"Total Points", type:'number', width:'5%', editable:false, auto:'formula:points'},
  {key:'grade',  label:'Grade',       type:'select', options:['Pass','Fail','Hold'], width:'5%'},
  {key:'epi',    label:'Epi',         type:'number', width:'4%'},
  {key:'ppi',    label:'Ppi',         type:'number', width:'4%'},
  {key:'gsm',    label:'Gsm',         type:'number', width:'4%'},
  {key:'width',  label:'Width',       type:'number', width:'4%'},
  {key:'shrinkWarp', label:'Shrink Warp %', type:'text', width:'5%'},
  {key:'shrinkWeft', label:'Shrink Weft %', type:'text', width:'5%'},
  {key:'batch',  label:'Batch',       type:'text',   width:'4%'},
  {key:'remark', label:'Remark',      type:'text',   width:'12%'},
];

const DEFAULT_PARAMS = ['Ph','Csv','Hand feel','Rubbing Dry','Rubbing wet','Appearance','Front back variation',
  'Tear strength','Tensile strength','Seam slippage','Martindale abrasion','Pilling','Afterwash hand feel'];

function blankRow(i){
  const r = {};
  DEFAULT_ROLL_COLUMNS.forEach(c=>r[c.key] = c.type==='number' ? '' : (c.key==='grade' ? 'Pass' : ''));
  r._id = 'r'+Date.now()+Math.random().toString(36).slice(2,7);
  return r;
}

function defaultState(){
  return {
    meta:{ company:'LISMOUNT TEX INDIA LLP', date:new Date().toISOString().slice(0,10), customer:'', unit:'', quality:'', shade:'', remarks:'' },
    columns: JSON.parse(JSON.stringify(DEFAULT_ROLL_COLUMNS)),
    rows: [blankRow(1)],
    params: DEFAULT_PARAMS.map(name=>({name, b1:'',b2:'',b3:'',b4:'',b5:'',hold:''})),
    summary: { rejection:0, bit:0, lab:0, hold:0, courier:0, target:0 },
  };
}

let state = defaultState();
let currentReportId = null;

const DEFAULT_FORMULAS = {
  cwidth:  'fWidth - 1',
  points:  'p1 + p2*2 + p3*3 + p4*4',
  shipped: 'totalMtrs',
  inspected: 'totalMtrs + rejection + bit + lab',
  shortpct: '(target - inspected) / target * 100',
};
let formulas = JSON.parse(JSON.stringify(DEFAULT_FORMULAS));

/* ---------------- safe formula evaluation ---------------- */
function safeEval(expr, vars){
  try{
    const names = Object.keys(vars);
    const vals = names.map(n=> Number(vars[n]) || 0);
    // eslint-disable-next-line no-new-func
    const fn = new Function(...names, `"use strict"; return (${expr});`);
    const out = fn(...vals);
    return (typeof out === 'number' && isFinite(out)) ? out : 0;
  }catch(e){ return NaN; }
}

/* ================= NAVIGATION ================= */
function go(view){
  document.querySelectorAll('#main > section').forEach(s=>s.classList.add('hidden'));
  document.getElementById('view-'+view).classList.remove('hidden');
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.toggle('active', n.dataset.view===view));
  if(view==='report') renderReport();
  if(view==='history') renderHistory();
  if(view==='dashboard') renderDashboard();
  if(view==='templates') renderTemplates();
}
document.querySelectorAll('.nav-item').forEach(n=> n.addEventListener('click', ()=>go(n.dataset.view)));

/* ================= RENDER: REPORT EDITOR ================= */
function renderReport(){
  document.getElementById('f-company').value = state.meta.company;
  document.getElementById('f-date').value = state.meta.date;
  document.getElementById('f-customer').value = state.meta.customer;
  document.getElementById('f-unit').value = state.meta.unit;
  document.getElementById('f-quality').value = state.meta.quality;
  document.getElementById('f-shade').value = state.meta.shade;
  document.getElementById('f-remarks').value = state.meta.remarks || '';
  document.getElementById('s-rejection').value = state.summary.rejection;
  document.getElementById('s-bit').value = state.summary.bit;
  document.getElementById('s-lab').value = state.summary.lab;
  document.getElementById('s-hold').value = state.summary.hold;
  document.getElementById('s-courier').value = state.summary.courier;
  document.getElementById('s-target').value = state.summary.target;

  renderRollTableHead();
  renderRollTableBody();
  renderParamTable();
  recalcAll();
  updateFilenamePreview();
}

function renderRollTableHead(){
  const head = document.getElementById('rollTableHead');
  head.innerHTML = state.columns.map(c=>`<th style="width:${c.width||'auto'}">${c.label}</th>`).join('') + '<th style="width:3%"></th>';
}

function cellHTML(row, col){
  const val = row[col.key] ?? '';
  if(col.editable===false){
    return `<td class="readonly">${val===''?'':val}</td>`;
  }
  if(col.type==='select'){
    const opts = col.options.map(o=>`<option ${o===val?'selected':''}>${o}</option>`).join('');
    return `<td><select data-row="${row._id}" data-col="${col.key}" onchange="onCellChange(this)">${opts}</select></td>`;
  }
  const type = col.type==='number' ? 'number' : 'text';
  return `<td><input type="${type}" data-row="${row._id}" data-col="${col.key}" value="${val===''?'':val}" onchange="onCellChange(this)"></td>`;
}

function renderRollTableBody(){
  const body = document.getElementById('rollTableBody');
  body.innerHTML = state.rows.map((r,i)=>{
    r.sno = i+1;
    const cells = state.columns.map(c=> c.key==='sno' ? `<td class="readonly">${i+1}</td>` : cellHTML(r,c)).join('');
    return `<tr>${cells}<td><div class="row-actions"><button class="icon-btn" title="Delete row" onclick="deleteRow('${r._id}')">🗑</button></div></td></tr>`;
  }).join('');
}

function renderParamTable(){
  const body = document.getElementById('paramTableBody');
  body.innerHTML = state.params.map((p,i)=>`
    <tr>
      <td style="text-align:left;"><input data-p="${i}" data-f="name" onchange="onParamChange(this)" value="${p.name}" style="text-align:left;font-weight:600;"></td>
      <td><input data-p="${i}" data-f="b1" onchange="onParamChange(this)" value="${p.b1}"></td>
      <td><input data-p="${i}" data-f="b2" onchange="onParamChange(this)" value="${p.b2}"></td>
      <td><input data-p="${i}" data-f="b3" onchange="onParamChange(this)" value="${p.b3}"></td>
      <td><input data-p="${i}" data-f="b4" onchange="onParamChange(this)" value="${p.b4}"></td>
      <td><input data-p="${i}" data-f="b5" onchange="onParamChange(this)" value="${p.b5}"></td>
      <td><input data-p="${i}" data-f="hold" onchange="onParamChange(this)" value="${p.hold}"></td>
      <td><button class="icon-btn" onclick="deleteParamRow(${i})">🗑</button></td>
    </tr>`).join('');
}

function onCellChange(el){
  const row = state.rows.find(r=>r._id===el.dataset.row);
  if(!row) return;
  row[el.dataset.col] = el.value;
  recalcAll();
  scheduleAutosave();
}
function onParamChange(el){
  state.params[+el.dataset.p][el.dataset.f] = el.value;
  scheduleAutosave();
}

/* meta / summary field bindings */
['company','date','customer','unit','quality','shade'].forEach(k=>{
  document.addEventListener('DOMContentLoaded', ()=>{
    const el = document.getElementById('f-'+k);
    el.addEventListener('input', ()=>{ state.meta[k]=el.value; updateFilenamePreview(); scheduleAutosave(); });
  });
});
document.addEventListener('DOMContentLoaded', ()=>{
  document.getElementById('f-remarks').addEventListener('input', e=>{ state.meta.remarks = e.target.value; scheduleAutosave(); });
  ['rejection','bit','lab','hold','courier','target'].forEach(k=>{
    const el = document.getElementById('s-'+k);
    el.addEventListener('input', ()=>{ state.summary[k] = Number(el.value)||0; recalcAll(); scheduleAutosave(); });
  });
});

function addRow(){ state.rows.push(blankRow(state.rows.length+1)); renderRollTableBody(); recalcAll(); scheduleAutosave(); }
function deleteRow(id){ state.rows = state.rows.filter(r=>r._id!==id); renderRollTableBody(); recalcAll(); scheduleAutosave(); }
function addParamRow(){ state.params.push({name:'New Parameter',b1:'',b2:'',b3:'',b4:'',b5:'',hold:''}); renderParamTable(); scheduleAutosave(); }
function deleteParamRow(i){ state.params.splice(i,1); renderParamTable(); scheduleAutosave(); }

function addColumnPrompt(){
  const label = prompt('New column name:');
  if(!label) return;
  const key = 'custom_'+label.toLowerCase().replace(/[^a-z0-9]+/g,'_').replace(/^_+|_+$/g,'')+'_'+Date.now().toString(36);
  state.columns.push({key, label, type:'text', width:'6%', editable:true});
  state.rows.forEach(r=> r[key] = '');
  renderRollTableHead(); renderRollTableBody(); scheduleAutosave();
}

/* ================= CALCULATION ENGINE ================= */
function recalcAll(){
  let totalMtrs = 0, gsmSum = 0, gsmCount = 0;
  state.rows.forEach(row=>{
    const vars = { fWidth:row.fWidth, mtrs:row.mtrs, p1:row.p1, p2:row.p2, p3:row.p3, p4:row.p4 };
    const cw = row.fWidth!=='' && row.fWidth!=null ? safeEval(formulas.cwidth, vars) : '';
    row.cWidth = cw===''? '' : round2(cw);
    const hasPoints = [row.p1,row.p2,row.p3,row.p4].some(v=>v!=='' && v!=null);
    row.totalPoints = hasPoints ? round2(safeEval(formulas.points, vars)) : '';
    const m = Number(row.mtrs)||0;
    totalMtrs += m;
    if(row.gsm!=='' && row.gsm!=null){ gsmSum += Number(row.gsm)||0; gsmCount++; }
  });
  const avgGsm = gsmCount ? round2(gsmSum/gsmCount) : 0;
  const {rejection,bit,lab,hold,courier,target} = state.summary;
  const shipped = round2(safeEval(formulas.shipped, {totalMtrs,avgGsm,target,rejection,bit,lab,hold,courier}));
  const inspected = round2(safeEval(formulas.inspected, {totalMtrs,avgGsm,target,rejection,bit,lab,hold,courier}));
  const shortPct = target ? round2(safeEval(formulas.shortpct, {target, inspected, totalMtrs})) : 0;
  const shortQty = target ? round2(target - inspected) : 0;

  state._computed = {totalMtrs:round2(totalMtrs), avgGsm, shipped, inspected, shortPct, shortQty};

  // reflect readonly cells in DOM without full re-render (keeps focus)
  document.querySelectorAll('#rollTableBody tr').forEach((tr,i)=>{
    const row = state.rows[i]; if(!row) return;
    const cwCell = tr.children[state.columns.findIndex(c=>c.key==='cWidth')];
    const tpCell = tr.children[state.columns.findIndex(c=>c.key==='totalPoints')];
    if(cwCell) cwCell.textContent = row.cWidth;
    if(tpCell) tpCell.textContent = row.totalPoints;
  });

  const totMtEl = document.getElementById('calc-totalmtrs'); if(totMtEl) totMtEl.textContent = state._computed.totalMtrs;
  const avgEl = document.getElementById('calc-avggsm'); if(avgEl) avgEl.textContent = state._computed.avgGsm;
  const shipEl = document.getElementById('calc-shipped'); if(shipEl) shipEl.textContent = state._computed.shipped;
  const shortEl = document.getElementById('calc-short');
  if(shortEl){
    shortEl.textContent = target ? `${state._computed.shortQty} (${state._computed.shortPct}%)` : '—';
    document.getElementById('short-card').classList.toggle('warn', state._computed.shortPct > 2);
  }
}
function round2(n){ n = Number(n); return Math.round(n*100)/100; }

function applyFormulas(){
  const f = {
    cwidth: document.getElementById('formula-cwidth').value,
    points: document.getElementById('formula-points').value,
    shipped: document.getElementById('formula-shipped').value,
    inspected: document.getElementById('formula-inspected').value,
    shortpct: document.getElementById('formula-shortpct').value,
  };
  // validate against sample vars
  const testVars = {fWidth:10,mtrs:10,p1:1,p2:1,p3:1,p4:1,totalMtrs:10,avgGsm:10,target:10,rejection:0,bit:0,lab:0,hold:0,courier:0,inspected:10};
  let ok = true;
  Object.entries(f).forEach(([k,expr])=>{ if(isNaN(safeEval(expr, testVars))) ok = false; });
  document.getElementById('formula-status').textContent = ok ? 'valid' : 'error — check syntax';
  document.getElementById('formula-status').className = 'badge ' + (ok?'ok':'err');
  if(!ok) return;
  formulas = f;
  localStorage.setItem(LS_KEYS.formulas, JSON.stringify(formulas));
  recalcAll();
  alert('Formulas applied and report recalculated.');
}
function resetFormulas(){
  formulas = JSON.parse(JSON.stringify(DEFAULT_FORMULAS));
  document.getElementById('formula-cwidth').value = formulas.cwidth;
  document.getElementById('formula-points').value = formulas.points;
  document.getElementById('formula-shipped').value = formulas.shipped;
  document.getElementById('formula-inspected').value = formulas.inspected;
  document.getElementById('formula-shortpct').value = formulas.shortpct;
  localStorage.setItem(LS_KEYS.formulas, JSON.stringify(formulas));
  recalcAll();
}

/* ================= NEW / AUTOSAVE / RESTORE ================= */
function newReport(confirmFirst){
  if(confirmFirst && !confirm('Start a new blank report? Unsaved changes in the current editor will be lost (autosave keeps the last snapshot).')) return;
  state = defaultState();
  currentReportId = null;
  go('report');
}

let autosaveTimer = null;
function scheduleAutosave(){
  clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(()=>{
    localStorage.setItem(LS_KEYS.current, JSON.stringify(state));
    const t = document.getElementById('autosave-time');
    if(t) t.textContent = 'Last saved ' + new Date().toLocaleTimeString();
  }, 500);
}

function saveReportToHistory(){
  let list = JSON.parse(localStorage.getItem(LS_KEYS.reports) || '[]');
  const id = currentReportId || ('rep_'+Date.now());
  currentReportId = id;
  const record = {
    id, savedAt: new Date().toISOString(),
    meta: state.meta, computed: state._computed,
    state: JSON.parse(JSON.stringify(state)),
  };
  const idx = list.findIndex(r=>r.id===id);
  if(idx>=0) list[idx] = record; else list.unshift(record);
  localStorage.setItem(LS_KEYS.reports, JSON.stringify(list));
  alert('Report saved to history.');
}

function loadAllFromStorage(){
  const savedCurrent = localStorage.getItem(LS_KEYS.current);
  if(savedCurrent){ try{ state = JSON.parse(savedCurrent); }catch(e){} }
  const savedFormulas = localStorage.getItem(LS_KEYS.formulas);
  if(savedFormulas){ try{ formulas = JSON.parse(savedFormulas); }catch(e){} }
}

/* ================= HISTORY VIEW ================= */
function renderHistory(){
  const list = JSON.parse(localStorage.getItem(LS_KEYS.reports) || '[]');
  const q = (document.getElementById('hist-search').value || '').toLowerCase();
  const filtered = list.filter(r=>{
    const hay = [r.meta.customer, r.meta.quality, r.meta.shade, r.meta.company].join(' ').toLowerCase();
    return hay.includes(q);
  });
  document.getElementById('history-empty').classList.toggle('hidden', list.length>0);
  document.getElementById('historyBody').innerHTML = filtered.map(r=>`
    <tr>
      <td>${r.meta.date||''}</td><td>${esc(r.meta.customer)}</td><td>${esc(r.meta.quality)}</td><td>${esc(r.meta.shade)}</td>
      <td>${r.computed? r.computed.totalMtrs : ''}</td>
      <td>${new Date(r.savedAt).toLocaleString()}</td>
      <td>
        <button class="btn small" onclick="openHistoryReport('${r.id}')">Open</button>
        <button class="btn small ghost" onclick="duplicateHistoryReport('${r.id}')">Duplicate</button>
        <button class="btn small danger" onclick="deleteHistoryReport('${r.id}')">Delete</button>
      </td>
    </tr>`).join('');
}
function esc(s){ return (s||'').toString().replace(/[<>&]/g, c=>({'<':'&lt;','>':'&gt;','&':'&amp;'}[c])); }
function openHistoryReport(id){
  const list = JSON.parse(localStorage.getItem(LS_KEYS.reports) || '[]');
  const rec = list.find(r=>r.id===id); if(!rec) return;
  state = JSON.parse(JSON.stringify(rec.state));
  currentReportId = id;
  go('report');
}
function duplicateHistoryReport(id){
  const list = JSON.parse(localStorage.getItem(LS_KEYS.reports) || '[]');
  const rec = list.find(r=>r.id===id); if(!rec) return;
  const copy = JSON.parse(JSON.stringify(rec));
  copy.id = 'rep_'+Date.now(); copy.savedAt = new Date().toISOString();
  list.unshift(copy);
  localStorage.setItem(LS_KEYS.reports, JSON.stringify(list));
  renderHistory();
}
function deleteHistoryReport(id){
  if(!confirm('Delete this report permanently?')) return;
  let list = JSON.parse(localStorage.getItem(LS_KEYS.reports) || '[]');
  list = list.filter(r=>r.id!==id);
  localStorage.setItem(LS_KEYS.reports, JSON.stringify(list));
  renderHistory(); renderDashboard();
}
document.addEventListener('DOMContentLoaded', ()=> document.getElementById('hist-search').addEventListener('input', renderHistory));

/* ================= DASHBOARD ================= */
function renderDashboard(){
  const list = JSON.parse(localStorage.getItem(LS_KEYS.reports) || '[]');
  const totalMtrs = list.reduce((s,r)=> s + (r.computed?.totalMtrs||0), 0);
  const cards = [
    {label:'Saved Reports', val:list.length},
    {label:'Total Mtrs Inspected (all)', val: round2(totalMtrs)},
    {label:'Unique Customers', val: new Set(list.map(r=>r.meta.customer).filter(Boolean)).size},
    {label:'Templates Saved', val: JSON.parse(localStorage.getItem(LS_KEYS.templates)||'[]').length},
  ];
  document.getElementById('dash-cards').innerHTML = cards.map(c=>`<div class="sum-card"><label>${c.label}</label><div class="val">${c.val}</div></div>`).join('');
  const recent = list.slice(0,6);
  document.getElementById('dash-recent').innerHTML = recent.length ? `<table class="hist-table"><thead><tr><th>Date</th><th>Customer</th><th>Quality</th><th>Shade</th><th></th></tr></thead><tbody>${
    recent.map(r=>`<tr><td>${r.meta.date||''}</td><td>${esc(r.meta.customer)}</td><td>${esc(r.meta.quality)}</td><td>${esc(r.meta.shade)}</td><td><button class="btn small" onclick="openHistoryReport('${r.id}');go('report')">Open</button></td></tr>`).join('')
  }</tbody></table>` : `<div class="hint">No reports yet. Click "New Report" to get started.</div>`;
}

/* ================= TEMPLATES ================= */
function openSaveTemplate(){ document.getElementById('templateNameInput').value=''; openModal('templateModal'); }
function confirmSaveTemplate(){
  const name = document.getElementById('templateNameInput').value.trim();
  if(!name) return alert('Enter a template name.');
  const list = JSON.parse(localStorage.getItem(LS_KEYS.templates) || '[]');
  list.unshift({ id:'tpl_'+Date.now(), name, savedAt:new Date().toISOString(),
    meta:{...state.meta, customer:'', shade:'', remarks:''}, columns: state.columns, params: state.params.map(p=>({...p,b1:'',b2:'',b3:'',b4:'',b5:'',hold:''})) });
  localStorage.setItem(LS_KEYS.templates, JSON.stringify(list));
  closeModal('templateModal');
  alert('Template saved: '+name);
}
function renderTemplates(){
  const list = JSON.parse(localStorage.getItem(LS_KEYS.templates) || '[]');
  document.getElementById('template-empty').classList.toggle('hidden', list.length>0);
  document.getElementById('templateList').innerHTML = list.map(t=>`
    <div class="sum-card" style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
      <div><b>${esc(t.name)}</b><div class="hint">Saved ${new Date(t.savedAt).toLocaleDateString()} · ${t.columns.length} columns</div></div>
      <div class="btnrow">
        <button class="btn small" onclick="useTemplate('${t.id}')">Use for New Report</button>
        <button class="btn small danger" onclick="deleteTemplate('${t.id}')">Delete</button>
      </div>
    </div>`).join('');
}
function useTemplate(id){
  const list = JSON.parse(localStorage.getItem(LS_KEYS.templates) || '[]');
  const t = list.find(x=>x.id===id); if(!t) return;
  state = defaultState();
  state.meta = {...state.meta, ...t.meta};
  state.columns = JSON.parse(JSON.stringify(t.columns));
  state.params = JSON.parse(JSON.stringify(t.params));
  state.rows = [blankRow(1)];
  DEFAULT_ROLL_COLUMNS_SYNC();
  currentReportId = null;
  go('report');
}
function DEFAULT_ROLL_COLUMNS_SYNC(){
  // ensure existing rows have keys for every column
  state.rows.forEach(r=> state.columns.forEach(c=>{ if(!(c.key in r)) r[c.key] = ''; }));
}
function deleteTemplate(id){
  if(!confirm('Delete this template?')) return;
  let list = JSON.parse(localStorage.getItem(LS_KEYS.templates) || '[]');
  list = list.filter(t=>t.id!==id);
  localStorage.setItem(LS_KEYS.templates, JSON.stringify(list));
  renderTemplates();
}

/* ================= MODALS ================= */
function openModal(id){ document.getElementById(id).classList.remove('hidden'); }
function closeModal(id){ document.getElementById(id).classList.add('hidden'); }

/* ================= FILENAME ================= */
function updateFilenamePreview(){
  const el = document.getElementById('report-filename-preview');
  if(el) el.textContent = buildFilename();
}
function buildFilename(){
  const prefix = (document.getElementById('opt-filename')?.value) || 'Fabric_Inspection_Report';
  const parts = [prefix, state.meta.quality, state.meta.shade, state.meta.date].filter(Boolean);
  return parts.join('_').replace(/[^a-zA-Z0-9_\-]/g,'_') ;
}

/* ================= EXCEL EXPORT (real cells + formulas) ================= */
function exportExcel(){
  const wb = XLSX.utils.book_new();
  const aoa = [];
  const merges = [];
  const push = (arr)=>{ aoa.push(arr); return aoa.length-1; };

  push([state.meta.company]);
  merges.push({s:{r:0,c:0}, e:{r:0,c:state.columns.length}});
  push([`Date: ${state.meta.date}`,'','',`Customer: ${state.meta.customer}`,'','',`Unit: ${state.meta.unit}`,'','',`Quality: ${state.meta.quality}`,'','',`Shade: ${state.meta.shade}`]);
  push([]);
  const headerRowIdx = push(state.columns.map(c=>c.label));
  const firstDataRow = aoa.length; // 0-indexed
  state.rows.forEach((r,i)=>{
    const excelRow = firstDataRow + i + 1; // 1-indexed excel row number for formula refs
    const rowArr = state.columns.map(c=>{
      if(c.key==='sno') return i+1;
      if(c.key==='cWidth'){
        const fCol = colLetter(state.columns.findIndex(cc=>cc.key==='fWidth'));
        return { f: `${fCol}${excelRow}-1` };
      }
      if(c.key==='totalPoints'){
        const p1 = colLetter(state.columns.findIndex(cc=>cc.key==='p1'));
        const p2 = colLetter(state.columns.findIndex(cc=>cc.key==='p2'));
        const p3 = colLetter(state.columns.findIndex(cc=>cc.key==='p3'));
        const p4 = colLetter(state.columns.findIndex(cc=>cc.key==='p4'));
        return { f: `${p1}${excelRow}+${p2}${excelRow}*2+${p3}${excelRow}*3+${p4}${excelRow}*4` };
      }
      return r[c.key] ?? '';
    });
    push(rowArr);
  });
  const lastDataExcelRow = firstDataRow + state.rows.length; // 1-indexed last data row
  const mtrsColIdx = state.columns.findIndex(c=>c.key==='mtrs');
  const gsmColIdx = state.columns.findIndex(c=>c.key==='gsm');
  const totalRow = state.columns.map((c,ci)=>{
    if(ci===0) return 'TOTAL';
    if(ci===mtrsColIdx) return { f:`SUM(${colLetter(mtrsColIdx)}${firstDataRow+1}:${colLetter(mtrsColIdx)}${lastDataExcelRow})` };
    if(ci===gsmColIdx) return { f:`AVERAGE(${colLetter(gsmColIdx)}${firstDataRow+1}:${colLetter(gsmColIdx)}${lastDataExcelRow})` };
    return '';
  });
  push(totalRow);
  push([]);
  const summaryStart = push(['Total Rejection Qty', state.summary.rejection, '', 'Total Bit', state.summary.bit, '', 'Lab Testing', state.summary.lab]);
  push(['Total Hold', state.summary.hold, '', 'Sample Courier', state.summary.courier, '', 'Greige Issued (target)', state.summary.target]);
  const mtrsColLetter = colLetter(mtrsColIdx);
  const totalRowExcel = lastDataExcelRow + 1;
  push(['Total MTRS Inspected', {f:`${mtrsColLetter}${totalRowExcel}`}]);
  push(['Shipped Qty', {f:`${mtrsColLetter}${totalRowExcel}`}]);
  push([]);
  push(['Additional Test Parameters']);
  push(['Parameter','Batch-1','Batch-2','Batch-3','Batch-4','Batch-5','Hold']);
  state.params.forEach(p=> push([p.name,p.b1,p.b2,p.b3,p.b4,p.b5,p.hold]));
  push([]);
  push(['Remarks', state.meta.remarks||'']);

  const ws = XLSX.utils.aoa_to_sheet(aoa.map(row=> row.map(cell=>{
    if(cell && typeof cell==='object' && 'f' in cell) return {t:'n', f:cell.f};
    return cell;
  })));
  ws['!merges'] = merges;
  ws['!cols'] = state.columns.map(c=>({wch: Math.max(8, (c.label||'').length) }));
  ws['!pagesetup'] = { paperSize:9, orientation: (document.getElementById('opt-orientation')?.value==='portrait'?'portrait':'landscape'), fitToPage:true };
  ws['!fitToPage'] = true;
  ws['!margins'] = { left:0.3, right:0.3, top:0.4, bottom:0.4, header:0.2, footer:0.2 };
  if(!wb.Workbook) wb.Workbook = {};
  wb.Workbook.Views = [{RTL:false}];
  XLSX.utils.book_append_sheet(wb, ws, 'Inspection Report');
  ws['!pageSetup'] = { orientation: (document.getElementById('opt-orientation')?.value==='portrait'?'portrait':'landscape'), paperSize:9, fitToWidth:1, fitToHeight:0 };

  XLSX.writeFile(wb, buildFilename() + '.xlsx');
}
function colLetter(idx){
  let s=''; idx++;
  while(idx>0){ let m=(idx-1)%26; s=String.fromCharCode(65+m)+s; idx=Math.floor((idx-1)/26); }
  return s;
}

/* ================= EXCEL IMPORT ================= */
function importExcelFile(file){
  const reader = new FileReader();
  reader.onload = (e)=>{
    try{
      const wb = XLSX.read(e.target.result, {type:'array'});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:'', raw:true});
      mapImportedRowsToState(rows);
      alert('Excel imported: '+state.rows.length+' rows loaded into the editable report.');
      go('report');
    }catch(err){
      alert('Could not read this Excel file: '+err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function mapImportedRowsToState(rows){
  // find header row: the row with the most matches to known labels
  const knownLabels = DEFAULT_ROLL_COLUMNS.map(c=>c.label.toLowerCase());
  let headerRowIdx = -1, bestScore = 0;
  rows.forEach((row,i)=>{
    const score = row.filter(cell=> typeof cell==='string' && knownLabels.some(l=> cell.toLowerCase().includes(l.split(' ')[0].toLowerCase()))).length;
    if(score > bestScore){ bestScore = score; headerRowIdx = i; }
  });
  const newState = defaultState();
  if(headerRowIdx === -1){
    // fallback: treat first non-empty row as header
    headerRowIdx = rows.findIndex(r=> r.some(c=>c!==''));
  }
  const headerRow = rows[headerRowIdx] || [];
  const colMap = {}; // excelColIndex -> our key
  headerRow.forEach((label, ci)=>{
    if(typeof label !== 'string' || !label.trim()) return;
    const l = label.toLowerCase();
    const match = DEFAULT_ROLL_COLUMNS.find(c=> l.includes(c.label.toLowerCase().split(' ')[0].replace(/[.\(\)"]/g,'')) );
    if(match) colMap[ci] = match.key;
  });
  const dataRows = rows.slice(headerRowIdx+1).filter(r=> r.some(c=> c!==''));
  newState.rows = dataRows.slice(0, 500).map((r,i)=>{
    const nr = blankRow(i+1);
    Object.entries(colMap).forEach(([ci,key])=>{ nr[key] = r[ci] ?? ''; });
    return nr;
  });
  if(!newState.rows.length) newState.rows = [blankRow(1)];
  // try to pull header meta fields (company/customer/quality/shade) via keyword scan on early rows
  const flat = rows.slice(0, headerRowIdx+1);
  flat.forEach(r=> r.forEach(cell=>{
    if(typeof cell !== 'string') return;
    const c = cell.trim();
    const lc = c.toLowerCase();
    if(lc.startsWith('customer')) newState.meta.customer = c.split(':').slice(1).join(':').trim();
    else if(lc.startsWith('quality')) newState.meta.quality = c.split(':').slice(1).join(':').trim();
    else if(lc.startsWith('shade')) newState.meta.shade = c.split(':').slice(1).join(':').trim();
    else if(lc.startsWith('date')) newState.meta.date = c.split(':').slice(1).join(':').trim() || newState.meta.date;
    else if(lc.includes('processing unit')) newState.meta.unit = c.split(':').slice(1).join(':').trim();
  }));
  if(flat[0] && typeof flat[0][0]==='string' && flat[0][0].length>3 && !flat[0][0].includes(':')) newState.meta.company = flat[0][0];
  state = newState;
}

/* ================= PDF IMPORT (text-based only) ================= */
function importPdfFile(file, rawTextTargetId){
  const reader = new FileReader();
  reader.onload = async (e)=>{
    try{
      const pdf = await pdfjsLib.getDocument({data:e.target.result}).promise;
      let text = '';
      for(let p=1;p<=pdf.numPages;p++){
        const page = await pdf.getPage(p);
        const content = await page.getTextContent();
        text += content.items.map(it=>it.str).join(' ') + '\n';
      }
      if(rawTextTargetId){
        document.getElementById('pdfRawWrap').classList.remove('hidden');
        document.getElementById(rawTextTargetId).value = text.trim() || '(No extractable text — this PDF is likely a scanned image and needs OCR, which requires a server-side service.)';
      }
      if(!text.trim()){
        alert('No text layer found in this PDF (likely a scanned/photographed document). Browser-only OCR isn\'t reliable enough to trust for inspection data — the raw text box is left empty. Consider re-exporting from source, or add a server-side OCR service for scanned imports.');
        return;
      }
      alert('Text extracted from PDF. Review it in the box below, then enter the data into the report editor (automatic table-structure detection from freeform PDF text is unreliable, so this is a manual-assist step rather than a silent auto-import).');
    }catch(err){
      alert('Could not read this PDF: '+err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}
if(window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/* ================= PDF EXPORT (real text, A4 auto-fit) ================= */
function buildA4HTML(){
  const orient = resolveOrientation();
  const cols = state.columns;
  const rowsHtml = state.rows.map((r,i)=>`<tr>${cols.map(c=>`<td>${c.key==='sno'? i+1 : (r[c.key]??'')}</td>`).join('')}</tr>`).join('');
  const headHtml = cols.map(c=>`<th>${c.label}</th>`).join('');
  const paramRows = state.params.map(p=>`<tr><td style="text-align:left;">${p.name}</td><td>${p.b1}</td><td>${p.b2}</td><td>${p.b3}</td><td>${p.b4}</td><td>${p.b5}</td><td>${p.hold}</td></tr>`).join('');
  return `
    <div class="a4-sheet" id="a4sheet-inner" style="width:${orient==='landscape'?'1123px':'794px'};">
      <h1>${esc(state.meta.company)}</h1>
      <div class="a4-meta">
        <span>Date: ${esc(state.meta.date)}</span><span>Customer: ${esc(state.meta.customer)}</span>
        <span>Unit: ${esc(state.meta.unit)}</span><span>Quality: ${esc(state.meta.quality)}</span>
        <span>Shade: ${esc(state.meta.shade)}</span><span>Shipped Qty: ${state._computed?.shipped ?? 0}</span>
      </div>
      <table class="pdf-tbl"><thead><tr>${headHtml}</tr></thead><tbody>${rowsHtml}
        <tr style="font-weight:700;background:#f0ead9;"><td colspan="${cols.findIndex(c=>c.key==='mtrs')}">TOTAL</td><td>${state._computed?.totalMtrs ?? 0}</td><td colspan="${cols.length - cols.findIndex(c=>c.key==='mtrs') - 2}"></td><td>Avg GSM: ${state._computed?.avgGsm ?? 0}</td></tr>
      </tbody></table>
      <div class="a4-meta">
        <span>Total Rejection Qty: ${state.summary.rejection}</span><span>Total Bit: ${state.summary.bit}</span>
        <span>Lab Testing: ${state.summary.lab}</span><span>Total Hold: ${state.summary.hold}</span>
        <span>Sample Courier: ${state.summary.courier}</span><span>Greige Issued: ${state.summary.target}</span>
        <span>Short: ${state._computed?.shortQty ?? 0} (${state._computed?.shortPct ?? 0}%)</span>
      </div>
      <table class="pdf-tbl"><thead><tr><th style="text-align:left;">Parameter</th><th>Batch-1</th><th>Batch-2</th><th>Batch-3</th><th>Batch-4</th><th>Batch-5</th><th>Hold</th></tr></thead><tbody>${paramRows}</tbody></table>
      ${state.meta.remarks? `<div style="font-size:9px;margin-top:6px;"><b>Remarks:</b> ${esc(state.meta.remarks)}</div>`:''}
    </div>`;
}
function resolveOrientation(){
  const opt = document.getElementById('opt-orientation')?.value || 'landscape';
  if(opt!=='auto') return opt;
  return state.columns.length > 14 ? 'landscape' : 'portrait';
}
function openPreview(){
  recalcAll();
  document.getElementById('a4PreviewSheet').innerHTML = buildA4HTML();
  document.getElementById('preview-orient-label').textContent = '('+resolveOrientation()+')';
  openModal('previewModal');
}

function exportPDF(){
  recalcAll();
  const orient = resolveOrientation();
  const doc = new jspdf.jsPDF({ orientation: orient, unit:'mm', format:'a4' });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const margin = 8;

  doc.setFont('times','bold'); doc.setFontSize(16);
  doc.text(state.meta.company || '', pageW/2, margin+4, {align:'center'});
  doc.setFontSize(8); doc.setFont('times','normal');
  const metaLine = `Date: ${state.meta.date}    Customer: ${state.meta.customer}    Unit: ${state.meta.unit}    Quality: ${state.meta.quality}    Shade: ${state.meta.shade}    Shipped Qty: ${state._computed.shipped}`;
  doc.text(metaLine, margin, margin+10);

  const cols = state.columns;
  const head = [cols.map(c=>c.label)];
  const body = state.rows.map((r,i)=> cols.map(c=> c.key==='sno' ? (i+1) : (r[c.key] ?? '')));
  const totalRowIdx = cols.findIndex(c=>c.key==='mtrs');
  const totals = cols.map((c,ci)=> ci===0?'TOTAL': ci===totalRowIdx? state._computed.totalMtrs : (c.key==='gsm'? 'Avg:'+state._computed.avgGsm : ''));
  body.push(totals);

  const fit = (document.getElementById('opt-fit')?.value ?? 'yes') === 'yes';
  let fontSize = 7;
  const availW = pageW - margin*2;
  // shrink font as column count grows to try to keep it on one page width-wise
  if(cols.length > 18) fontSize = 5.5;
  else if(cols.length > 12) fontSize = 6.5;

  doc.autoTable({
    head, body, startY: margin+13, margin:{left:margin, right:margin},
    styles:{ fontSize, cellPadding:1, halign:'center', lineColor:[60,60,60], lineWidth:0.1, font:'times' },
    headStyles:{ fillColor:[28,43,58], textColor:255, fontStyle:'bold' },
    tableWidth: availW,
    theme:'grid',
  });

  let y = doc.lastAutoTable.finalY + 5;
  doc.setFontSize(8);
  const s = state.summary, c = state._computed;
  const sumLine1 = `Total Rejection Qty: ${s.rejection}   Total Bit: ${s.bit}   Lab Testing: ${s.lab}   Total Hold: ${s.hold}   Sample Courier: ${s.courier}   Greige Issued: ${s.target}`;
  const sumLine2 = `Total MTRS Inspected: ${c.totalMtrs}   Average GSM: ${c.avgGsm}   Short Qty: ${c.shortQty}  (${c.shortPct}%)`;
  doc.text(sumLine1, margin, y); y+=5;
  doc.text(sumLine2, margin, y); y+=7;

  if(y < pageH - margin - 20){
    doc.autoTable({
      head:[['Parameter','Batch-1','Batch-2','Batch-3','Batch-4','Batch-5','Hold']],
      body: state.params.map(p=>[p.name,p.b1,p.b2,p.b3,p.b4,p.b5,p.hold]),
      startY:y, margin:{left:margin,right:margin},
      styles:{fontSize:6.5,cellPadding:1,halign:'center',font:'times'},
      headStyles:{fillColor:[28,43,58],textColor:255},
      tableWidth: availW, theme:'grid',
    });
    y = doc.lastAutoTable.finalY + 5;
  }
  if(state.meta.remarks){
    doc.setFontSize(7.5);
    doc.text('Remarks: '+state.meta.remarks, margin, Math.min(y, pageH-margin));
  }

  doc.save(buildFilename()+'.pdf');
}

/* ================= WIRE UP FILE INPUTS ================= */
document.addEventListener('DOMContentLoaded', ()=>{
  document.getElementById('excelInput').addEventListener('change', e=>{ if(e.target.files[0]) importExcelFile(e.target.files[0]); e.target.value=''; });
  document.getElementById('excelInput2').addEventListener('change', e=>{ if(e.target.files[0]) importExcelFile(e.target.files[0]); e.target.value=''; });
  document.getElementById('pdfInput').addEventListener('change', e=>{ if(e.target.files[0]) importPdfFile(e.target.files[0]); e.target.value=''; });
  document.getElementById('pdfInput2').addEventListener('change', e=>{ if(e.target.files[0]) importPdfFile(e.target.files[0], 'pdfRawText'); e.target.value=''; });

  loadAllFromStorage();
  document.getElementById('formula-cwidth').value = formulas.cwidth;
  document.getElementById('formula-points').value = formulas.points;
  document.getElementById('formula-shipped').value = formulas.shipped;
  document.getElementById('formula-inspected').value = formulas.inspected;
  document.getElementById('formula-shortpct').value = formulas.shortpct;

  renderDashboard();
  go('dashboard');
});
