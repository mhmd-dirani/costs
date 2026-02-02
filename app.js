/* Workshop Payments - client-side XLSX viewer/editor
   - Reads XLSX using SheetJS
   - Lets you choose sheet, view rows, add/delete, sort, and export XLSX/CSV
*/
const state = {
  workbook: null,           // SheetJS workbook
  sheets: {},               // { sheetName: [ {who, why, amount} ] }
  activeSheet: null,
  sort: { key: null, asc: true }
};

const $ = (id) => document.getElementById(id);

const fileInput = $("fileInput");
const sheetSelect = $("sheetSelect");
const totalAmount = $("totalAmount");
const rowCount = $("rowCount");
const tableBody = $("paymentsTable").querySelector("tbody");
const downloadXlsx = $("downloadXlsx");
const downloadCsv = $("downloadCsv");
const addForm = $("addForm");
const addBtn = $("addBtn");
const clearLocal = $("clearLocal");

// --- helpers ---
function normalizeHeader(h){
  if(!h) return "";
  return String(h).trim().toLowerCase().replace(/\s+/g," ");
}
function coerceAmount(v){
  if(v === null || v === undefined || v === "") return 0;
  if(typeof v === "number") return v;
  const s = String(v).replace(/[, ]+/g,"").trim();
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}
function formatNumber(n){
  try { return new Intl.NumberFormat().format(n); } catch { return String(n); }
}
function clone(obj){ return JSON.parse(JSON.stringify(obj)); }

function computeTotal(rows){
  return rows.reduce((sum, r) => sum + coerceAmount(r.amount), 0);
}

function enableUI(enabled){
  sheetSelect.disabled = !enabled;
  downloadXlsx.disabled = !enabled;
  downloadCsv.disabled = !enabled;
  addBtn.disabled = !enabled;
}

function toRowsFromSheet(sheet){
  // Convert worksheet to rows with normalized fields.
  // Accepts various column name spellings:
  // who / Who, why / Why, how much / How Much / amount, etc.
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  if(!raw.length) return [];

  // If headers are weird, SheetJS still returns objects; normalize keys
  const rows = raw.map((r) => {
    const out = { who: "", why: "", amount: 0 };
    for(const [k, v] of Object.entries(r)){
      const nk = normalizeHeader(k);
      if(nk === "who" || nk === "to" || nk === "paid to" || nk === "name"){
        out.who = String(v).trim();
      } else if(nk === "why" || nk === "reason" || nk === "description" || nk === "for"){
        out.why = String(v).trim();
      } else if(nk === "how much" || nk === "amount" || nk === "value" || nk === "cost"){
        out.amount = coerceAmount(v);
      }
    }
    return out;
  }).filter(r => (r.who || r.why || r.amount)); // drop empty lines
  return rows;
}

function render(){
  const sheetName = state.activeSheet;
  const rows = sheetName ? state.sheets[sheetName] : [];
  tableBody.innerHTML = "";

  if(!sheetName){
    totalAmount.textContent = "—";
    rowCount.textContent = "—";
    return;
  }

  // Sort
  const sorted = clone(rows);
  const { key, asc } = state.sort;
  if(key){
    sorted.sort((a,b) => {
      const av = key === "amount" ? coerceAmount(a[key]) : String(a[key] ?? "").toLowerCase();
      const bv = key === "amount" ? coerceAmount(b[key]) : String(b[key] ?? "").toLowerCase();
      if(av < bv) return asc ? -1 : 1;
      if(av > bv) return asc ? 1 : -1;
      return 0;
    });
  }

  sorted.forEach((r, idx) => {
    const tr = document.createElement("tr");

    const tdWho = document.createElement("td");
    tdWho.textContent = r.who || "";
    tr.appendChild(tdWho);

    const tdWhy = document.createElement("td");
    tdWhy.textContent = r.why || "";
    tr.appendChild(tdWhy);

    const tdAmt = document.createElement("td");
    tdAmt.className = "right";
    tdAmt.textContent = formatNumber(coerceAmount(r.amount));
    tr.appendChild(tdAmt);

    const tdAct = document.createElement("td");
    tdAct.className = "right";
    const wrap = document.createElement("div");
    wrap.className = "row-actions";

    const del = document.createElement("button");
    del.type = "button";
    del.className = "iconbtn";
    del.textContent = "Delete";
    del.title = "Delete this row";
    del.addEventListener("click", () => {
      // delete by identity: find first matching row in original list
      const originalIndex = state.sheets[sheetName].findIndex(x =>
        x.who === r.who && x.why === r.why && coerceAmount(x.amount) === coerceAmount(r.amount)
      );
      if(originalIndex >= 0){
        state.sheets[sheetName].splice(originalIndex, 1);
        persistToLocal();
        render();
      }
    });

    wrap.appendChild(del);
    tdAct.appendChild(wrap);
    tr.appendChild(tdAct);

    tableBody.appendChild(tr);
  });

  rowCount.textContent = String(rows.length);
  totalAmount.textContent = formatNumber(computeTotal(rows));
}

function populateSheetSelect(){
  sheetSelect.innerHTML = "";
  const opt0 = document.createElement("option");
  opt0.value = "";
  opt0.textContent = "—";
  sheetSelect.appendChild(opt0);

  Object.keys(state.sheets).forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    sheetSelect.appendChild(opt);
  });
}

function setActiveSheet(name){
  state.activeSheet = name || null;
  sheetSelect.value = name || "";
  render();
}

function workbookFromState(){
  // Create a new workbook from current state (all sheets).
  const wb = XLSX.utils.book_new();
  Object.entries(state.sheets).forEach(([name, rows]) => {
    const aoa = [["Who","Why","How Much"]];
    rows.forEach(r => aoa.push([r.who, r.why, coerceAmount(r.amount)]));
    // Add a total row (nice in Excel)
    aoa.push([]);
    aoa.push(["", "TOTAL", computeTotal(rows)]);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, name);
  });
  return wb;
}

function downloadBlob(blob, filename){
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1500);
}

function exportXlsx(){
  const wb = workbookFromState();
  const out = XLSX.write(wb, { bookType:"xlsx", type:"array" });
  const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const date = new Date().toISOString().slice(0,10);
  downloadBlob(blob, `workshop_payments_${date}.xlsx`);
}

function exportCsvCurrent(){
  if(!state.activeSheet) return;
  const rows = state.sheets[state.activeSheet] || [];
  const aoa = [["Who","Why","How Much"], ...rows.map(r => [r.who, r.why, coerceAmount(r.amount)])];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const csv = XLSX.utils.sheet_to_csv(ws);
  downloadBlob(new Blob([csv], { type:"text/csv;charset=utf-8" }), `${state.activeSheet}.csv`);
}

// --- local persistence (optional, so you don't lose edits if you refresh) ---
const LS_KEY = "workshop_payments_state_v1";
function persistToLocal(){
  try{
    const payload = { sheets: state.sheets, activeSheet: state.activeSheet, sort: state.sort };
    localStorage.setItem(LS_KEY, JSON.stringify(payload));
  } catch(e){ /* ignore */ }
}
function restoreFromLocal(){
  try{
    const raw = localStorage.getItem(LS_KEY);
    if(!raw) return false;
    const payload = JSON.parse(raw);
    if(payload && payload.sheets){
      state.sheets = payload.sheets;
      state.activeSheet = payload.activeSheet || null;
      state.sort = payload.sort || state.sort;
      populateSheetSelect();
      enableUI(Object.keys(state.sheets).length > 0);
      setActiveSheet(state.activeSheet || Object.keys(state.sheets)[0] || null);
      return true;
    }
  } catch(e){ /* ignore */ }
  return false;
}
function clearLocalState(){
  try{ localStorage.removeItem(LS_KEY); } catch(e){}
  location.reload();
}

// --- events ---
fileInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if(!file) return;

  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:"array" });

  state.workbook = wb;
  state.sheets = {};
  wb.SheetNames.forEach((name) => {
    const ws = wb.Sheets[name];
    state.sheets[name] = toRowsFromSheet(ws);
  });

  populateSheetSelect();
  enableUI(true);
  const first = wb.SheetNames[0] || null;
  setActiveSheet(first);
  persistToLocal();
});

sheetSelect.addEventListener("change", (e) => {
  setActiveSheet(e.target.value);
  persistToLocal();
});

downloadXlsx.addEventListener("click", exportXlsx);
downloadCsv.addEventListener("click", exportCsvCurrent);
clearLocal.addEventListener("click", clearLocalState);

addForm.addEventListener("submit", (e) => {
  e.preventDefault();
  if(!state.activeSheet) return;

  const who = $("whoInput").value.trim();
  const why = $("whyInput").value.trim();
  const amount = coerceAmount($("amountInput").value);

  if(!who || !why) return;

  state.sheets[state.activeSheet].push({ who, why, amount });
  $("whoInput").value = "";
  $("whyInput").value = "";
  $("amountInput").value = "";
  $("whoInput").focus();

  persistToLocal();
  render();
});

// column sort
document.querySelectorAll("th[data-key]").forEach(th => {
  th.addEventListener("click", () => {
    const key = th.dataset.key;
    if(state.sort.key === key){
      state.sort.asc = !state.sort.asc;
    } else {
      state.sort.key = key;
      state.sort.asc = true;
    }
    persistToLocal();
    render();
  });
});

// init
window.addEventListener("DOMContentLoaded", () => {
  const restored = restoreFromLocal();
  if(!restored){
    enableUI(false);
    render();
  }
});
