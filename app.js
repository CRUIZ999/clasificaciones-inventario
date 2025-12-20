/* =========================================================
   Maestro Inventarios - El Cedro (GitHub Pages)
   - Lee XLSX/CSV local (SheetJS)
   - Lee 3 hojas si existen:
     OUTPUT, RESUMEN_MES_SUC_CATEG, RESUMEN_MES_SUC_MARCA
   - Mantiene Resumen/Detalle y agrega:
     Buscador SKU, Oportunidades (Quiebres), Oportunidades (Sobreinv),
     Histórico Categorías, Histórico Marcas (con pivots $ y litros)
========================================================= */

/* ===================== CONFIG ===================== */
const BRANCHES = [
  { key: "adelitas", label: "ADELITAS", colKeyHint: "adelitas" },
  { key: "express", label: "EXPRESS", colKeyHint: "express" },
  { key: "general", label: "GENERAL", colKeyHint: "general" },
  { key: "ilustres", label: "H ILUSTRES", colKeyHint: "ilustres" },
  { key: "san_agust", label: "SAN AGUST", colKeyHint: "san agustin" },
];

const THRESH_RISK_DAYS = 15;
const THRESH_OVER_DAYS = 60;

// Para “Oportunidades” (recomendación)
const TARGET_RISK_DAYS = 15;     // objetivo mínimo
const TARGET_OVER_DAYS = 60;     // objetivo máximo (para definir exceso)

const SHEET_OUTPUT = "OUTPUT";
const SHEET_CATEG = "RESUMEN_MES_SUC_CATEG";
const SHEET_MARCA = "RESUMEN_MES_SUC_MARCA";

/* ===================== STATE ===================== */
const state = {
  fileName: null,
  loaded: false,

  outputRows: [],       // base SKU rows
  monthsUsed: null,

  // históricos
  histCateg: [],        // rows: {mes, sucursal, item, monto, litros}
  histMarca: [],

  // UI
  currentView: "summary",
  charts: {
    stacked: null,
    riskOver: null,
    dist: null,
  },

  // paginación (varias tablas)
  detail: { page: 1, pageSize: 100, sortKey: null, sortDir: "desc" },
  finder: { page: 1, pageSize: 100, sortKey: null, sortDir: "asc" },
  oppBreak: { page: 1, pageSize: 100, sortKey: null, sortDir: "desc" },
  oppOver: { page: 1, pageSize: 100, sortKey: null, sortDir: "desc" },
  histCategUI: { page: 1, pageSize: 100, sortKey: null, sortDir: "desc" },
  histMarcaUI: { page: 1, pageSize: 100, sortKey: null, sortDir: "desc" },
};

/* ===================== HELPERS ===================== */
function $(id){ return document.getElementById(id); }

function normKey(s){
  return String(s ?? "")
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"")
    .replace(/[^a-z0-9]+/g,"")
    .trim();
}

function toNumber(x){
  if (x === null || x === undefined || x === "") return 0;
  if (typeof x === "number") return isFinite(x) ? x : 0;
  const s = String(x).trim();
  // currency or formatted
  const clean = s.replace(/[^0-9.\-]/g,"");
  const n = parseFloat(clean);
  return isFinite(n) ? n : 0;
}

function toInt(x){
  const n = Math.round(toNumber(x));
  return isFinite(n) ? n : 0;
}

function formatInt(n){
  return new Intl.NumberFormat("es-MX", { maximumFractionDigits: 0 }).format(toNumber(n));
}

function formatMoney(n){
  return new Intl.NumberFormat("es-MX", { style:"currency", currency:"MXN", maximumFractionDigits: 2 }).format(toNumber(n));
}

// Promedio Vta Mes: 1 decimal solo si >0 y no entero. Si es entero o 0, sin decimal.
function formatProm(n){
  const v = toNumber(n);
  if (!isFinite(v) || v === 0) return "0";
  const isInt = Math.abs(v - Math.round(v)) < 1e-9;
  if (isInt) return String(Math.round(v));
  return new Intl.NumberFormat("es-MX", { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(v);
}

// Cobertura (Mes) y Cobertura Dias: sin decimales (redondeo)
function formatNoDecimals(n){
  return String(Math.round(toNumber(n)));
}

function pct(n){
  return new Intl.NumberFormat("es-MX", { style:"percent", maximumFractionDigits: 1 }).format(n);
}

function median(arr){
  const a = arr.filter(v => isFinite(v)).sort((x,y)=>x-y);
  if (!a.length) return 0;
  const mid = Math.floor(a.length/2);
  return a.length % 2 ? a[mid] : (a[mid-1]+a[mid])/2;
}

function clamp(v, min, max){ return Math.max(min, Math.min(max, v)); }

function normalizeSucursalLabel(s){
  const t = String(s||"").trim().toUpperCase();
  // ya vienen correctas según tú, pero dejamos tolerancia:
  if (t.includes("ADEL")) return "ADELITAS";
  if (t.includes("EXP")) return "EXPRESS";
  if (t.includes("GEN")) return "GENERAL";
  if (t.includes("ILUST")) return "H ILUSTRES";
  if (t.includes("SAN")) return "SAN AGUST";
  return t;
}

function isRisk(covDays){
  const d = toNumber(covDays);
  return d > 0 && d < THRESH_RISK_DAYS;
}
function isOver(covDays){
  const d = toNumber(covDays);
  return d > THRESH_OVER_DAYS;
}

function safeClass(x){
  const s = String(x||"").trim();
  if (!s) return "Sin Mov";
  if (s.toUpperCase() === "SIN MOV") return "Sin Mov";
  if (s === "A" || s === "B" || s === "C") return s;
  return s;
}

function findSheetName(wb, target){
  // match exact or case-insensitive or normalized
  const targetN = normKey(target);
  const names = wb.SheetNames || [];
  for (const n of names){
    if (normKey(n) === targetN) return n;
  }
  // contains
  for (const n of names){
    if (normKey(n).includes(targetN)) return n;
  }
  return null;
}

function sheetToRows(wb, sheetName){
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function mesYYYYMM(value){
  // Acepta: Date, number (excel date), string
  if (!value) return "";
  if (value instanceof Date && !isNaN(value.getTime())){
    const y = value.getFullYear();
    const m = String(value.getMonth()+1).padStart(2,"0");
    return `${y}-${m}`;
  }
  // SheetJS suele traer date como string o number según parsing.
  const s = String(value).trim();
  // si ya viene YYYY-MM, úsalo
  if (/^\d{4}-\d{2}$/.test(s)) return s;

  // si viene tipo "Sun Jun 01 2025 ..." o similar:
  const dt = new Date(s);
  if (!isNaN(dt.getTime())){
    const y = dt.getFullYear();
    const m = String(dt.getMonth()+1).padStart(2,"0");
    return `${y}-${m}`;
  }

  // fallback: intenta partir por /
  // 01/08/2025 -> 2025-08 (si fuera)
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m1){
    const y = m1[3];
    const mm = String(m1[2]).padStart(2,"0");
    return `${y}-${mm}`;
  }

  return s; // último recurso
}

/* ===================== PARSE OUTPUT ===================== */
function parseOutput(rows){
  if (!rows.length) return { rows: [], monthsUsed: 0 };

  // Build normalized key map per row for robust access
  const parsed = [];
  for (const r of rows){
    const keys = Object.keys(r);
    const nk = {};
    for (const k of keys){
      nk[normKey(k)] = k;
    }

    const codeKey = nk[normKey("Codigo")] || nk[normKey("Cve_prod")] || nk["codigo"] || nk["cveprod"];
    const descKey = nk[normKey("desc_prod")] || nk[normKey("Desc_prod")] || nk["descprod"] || nk["desc_prod"];
    const mesesKey = nk[normKey("MesesUsados")] || nk["mesesusados"];

    const codigo = String(r[codeKey] ?? "").trim();
    if (!codigo) continue;

    const obj = {
      codigo,
      desc: String(r[descKey] ?? "").trim(),
      mesesUsados: toInt(r[mesesKey]),
      byBranch: {}
    };

    for (const b of BRANCHES){
      const invK = nk[normKey(`Inv-${b.colKeyHint}`)] || nk[normKey(`Inv ${b.colKeyHint}`)] || nk[normKey(`Inv-${b.key}`)];
      const clsK = nk[normKey(`Clasificacion-${b.colKeyHint}`)] || nk[normKey(`Clasificación-${b.colKeyHint}`)] || nk[normKey(`Clasificacion ${b.colKeyHint}`)];
      const promK = nk[normKey(`Promedio Vta Mes-${b.colKeyHint}`)] || nk[normKey(`Promedio Vta Mes ${b.colKeyHint}`)];
      const covMesK = nk[normKey(`Cobertura (Mes)-${b.colKeyHint}`)] || nk[normKey(`Cobertura (Mes) -${b.colKeyHint}`)] || nk[normKey(`Cobertura Mes-${b.colKeyHint}`)];
      const covDiaK = nk[normKey(`Cobertura Dias (30)-${b.colKeyHint}`)] ||
                      nk[normKey(`Cobertura Dias (30) -${b.colKeyHint}`)] ||
                      nk[normKey(`Cobertura Dias (30) -${b.colKeyHint}`)] ||
                      nk[normKey(`Cobertura Dias (30)-${b.key}`)] ||
                      nk[normKey(`Cobertura Dias (30) -${b.key}`)];

      const inv = toInt(r[invK]);
      const cls = safeClass(r[clsK]);
      const prom = toNumber(r[promK]);
      const covMes = toNumber(r[covMesK]);
      const covDays = toNumber(r[covDiaK]);

      obj.byBranch[b.key] = { inv, cls, prom, covMes, covDays };
    }

    parsed.push(obj);
  }

  const monthsUsed = parsed.length ? (parsed[0].mesesUsados || 0) : 0;
  return { rows: parsed, monthsUsed };
}

/* ===================== PARSE HIST ===================== */
function parseHist(rows, mode /* "CATEG" | "MARCA" */){
  const out = [];
  for (const r of rows){
    const keys = Object.keys(r);
    const nk = {};
    for (const k of keys) nk[normKey(k)] = k;

    const mesK = nk[normKey("Mes")] || nk["mes"];
    const sucK = nk[normKey("Sucursal")] || nk["sucursal"];
    const itemK = mode === "CATEG"
      ? (nk[normKey("Categoria")] || nk["categoria"])
      : (nk[normKey("Marca")] || nk["marca"]);
    const montoK = nk[normKey("Monto_venta")] || nk[normKey("Monto venta")] || nk["montoventa"] || nk["monto_venta"];
    const litrosK = nk[normKey("Litros Vendidos")] || nk[normKey("LitrosVendidos")] || nk["litrosvendidos"] || nk["litrosvendidos"];

    const mes = mesYYYYMM(r[mesK]);
    const suc = normalizeSucursalLabel(r[sucK]);
    const item = String(r[itemK] ?? "").trim();
    if (!mes || !suc || !item) continue;

    const monto = toNumber(r[montoK]);
    const litros = toNumber(r[litrosK]);

    out.push({ mes, sucursal: suc, item, monto, litros });
  }
  return out;
}

/* ===================== UI INIT ===================== */
function setStatus(msg, meta=""){
  $("statusMsg").textContent = msg;
  $("statusMeta").textContent = meta || "";
}

function enableUI(enable){
  const ids = [
    "tabSummary","tabDetail","tabFinder","tabOppBreak","tabOppOver","tabHistCateg","tabHistMarca",
    "btnClear",
    "warehouseSelect","classSelect","searchInput","onlyInvToggle","covMin","covMax","onlySinMovInv","onlyABZero","btnExport",
    "detailPageSize","detailPrev","detailNext",
    "finderSearch","finderOnlyInv","finderExport","finderPageSize","finderPrev","finderNext",
    "oppBreakTarget","oppBreakOnlyRisk","oppBreakExport","oppBreakPageSize","oppBreakPrev","oppBreakNext",
    "oppOverTarget","oppOverOnlyOver","oppOverExport","oppOverPageSize","oppOverPrev","oppOverNext",
    "histCategSucursal","histCategSearch","histCategFrom","histCategTo","histCategExport","histCategPageSize","histCategPrev","histCategNext",
    "histMarcaSucursal","histMarcaSearch","histMarcaFrom","histMarcaTo","histMarcaExport","histMarcaPageSize","histMarcaPrev","histMarcaNext",
    "summaryWarehouseSelect"
  ];
  for (const id of ids){
    const el = $(id);
    if (el) el.disabled = !enable;
  }
}

function setActiveTab(tabId, viewId){
  const tabIds = ["tabSummary","tabDetail","tabFinder","tabOppBreak","tabOppOver","tabHistCateg","tabHistMarca"];
  const viewIds = ["summaryView","detailView","finderView","oppBreakView","oppOverView","histCategView","histMarcaView"];
  tabIds.forEach(id => $(id).classList.remove("active"));
  viewIds.forEach(id => $(id).classList.add("hidden"));
  $(tabId).classList.add("active");
  $(viewId).classList.remove("hidden");
}

function bindTabs(){
  $("tabSummary").addEventListener("click", () => { setActiveTab("tabSummary","summaryView"); renderSummary(); });
  $("tabDetail").addEventListener("click", () => { setActiveTab("tabDetail","detailView"); renderDetail(); });
  $("tabFinder").addEventListener("click", () => { setActiveTab("tabFinder","finderView"); renderFinder(); });
  $("tabOppBreak").addEventListener("click", () => { setActiveTab("tabOppBreak","oppBreakView"); renderOppBreak(); });
  $("tabOppOver").addEventListener("click", () => { setActiveTab("tabOppOver","oppOverView"); renderOppOver(); });
  $("tabHistCateg").addEventListener("click", () => { setActiveTab("tabHistCateg","histCategView"); renderHistCateg(); });
  $("tabHistMarca").addEventListener("click", () => { setActiveTab("tabHistMarca","histMarcaView"); renderHistMarca(); });
}

/* ===================== TABLE UTILS ===================== */
function sortRows(rows, key, dir){
  if (!key) return rows;
  const sgn = dir === "asc" ? 1 : -1;
  return [...rows].sort((a,b)=>{
    const va = a[key];
    const vb = b[key];
    const na = typeof va === "number";
    const nb = typeof vb === "number";
    if (na && nb) return (va - vb) * sgn;
    return String(va ?? "").localeCompare(String(vb ?? ""), "es", { sensitivity:"base" }) * sgn;
  });
}

function paginate(rows, page, pageSize){
  const total = rows.length;
  const pages = Math.max(1, Math.ceil(total / pageSize));
  const p = clamp(page, 1, pages);
  const start = (p-1)*pageSize;
  const end = start + pageSize;
  return { page: p, pages, total, slice: rows.slice(start,end) };
}

function renderTable(tableEl, columns, rows, onSort){
  // columns: [{key,label,fmt?,className?}]
  const thead = tableEl.querySelector("thead");
  const tbody = tableEl.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const trh = document.createElement("tr");
  for (const c of columns){
    const th = document.createElement("th");
    th.textContent = c.label;
    if (onSort){
      th.addEventListener("click", ()=> onSort(c.key));
      th.title = "Clic para ordenar";
    }
    trh.appendChild(th);
  }
  thead.appendChild(trh);

  const frag = document.createDocumentFragment();
  for (const r of rows){
    const tr = document.createElement("tr");
    for (const c of columns){
      const td = document.createElement("td");
      const v = r[c.key];
      td.textContent = c.fmt ? c.fmt(v, r) : (v ?? "");
      if (c.className) td.className = c.className;
      tr.appendChild(td);
    }
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
}

function downloadCSV(filename, columns, rows){
  const header = columns.map(c => `"${String(c.label).replace(/"/g,'""')}"`).join(",");
  const lines = rows.map(r => columns.map(c=>{
    const raw = r[c.key];
    const val = raw === null || raw === undefined ? "" : String(raw);
    return `"${val.replace(/"/g,'""')}"`;
  }).join(","));
  const csv = [header, ...lines].join("\n");
  const blob = new Blob([csv], { type:"text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ===================== SUMMARY COMPUTE ===================== */
function branchListForSummary(selected){
  if (selected && selected !== "ALL"){
    return BRANCHES.filter(b => b.label === selected);
  }
  return BRANCHES;
}

function computeBranchStats(branchKey, filteredRows){
  // filteredRows are outputRows (sku objects). Compute using branchKey
  let skus = 0, invSum = 0;
  const clsCount = { A:0, B:0, C:0, "Sin Mov":0 };
  const clsInv = { A:0, B:0, C:0, "Sin Mov":0 };
  const covDaysArr = [];
  let promSum = 0;

  for (const sku of filteredRows){
    const b = sku.byBranch[branchKey];
    if (!b) continue;
    skus += 1;
    invSum += b.inv;
    promSum += b.prom;

    const c = safeClass(b.cls);
    clsCount[c] = (clsCount[c] || 0) + 1;
    clsInv[c] = (clsInv[c] || 0) + b.inv;

    if (b.covDays > 0 && isFinite(b.covDays)) covDaysArr.push(b.covDays);
  }

  const covMed = median(covDaysArr);
  const riskCount = filteredRows.filter(s => isRisk(s.byBranch[branchKey]?.covDays)).length;
  const overCount = filteredRows.filter(s => isOver(s.byBranch[branchKey]?.covDays)).length;
  const sinMovInv = filteredRows.filter(s => safeClass(s.byBranch[branchKey]?.cls)==="Sin Mov" && (s.byBranch[branchKey]?.inv||0)>0).length;

  return {
    skus, invSum, promSum,
    clsCount, clsInv,
    covMed,
    riskCount, overCount, sinMovInv,
    riskPct: skus ? riskCount/skus : 0,
    overPct: skus ? overCount/skus : 0,
  };
}

function computeGlobal(selectedSucursalLabel){
  const rows = state.outputRows;
  const monthsUsed = state.monthsUsed || 0;

  if (!rows.length){
    return {
      monthsUsed, skus:0, inv:0, prom:0, covMed:0, riskPct:0, overPct:0, sinMovInv:0,
      perBranch:[]
    };
  }

  // executive selector:
  if (selectedSucursalLabel && selectedSucursalLabel !== "ALL"){
    const b = BRANCHES.find(x => x.label === selectedSucursalLabel);
    const st = computeBranchStats(b.key, rows);
    return {
      monthsUsed,
      skus: st.skus,
      inv: st.invSum,
      prom: st.promSum,
      covMed: st.covMed,
      riskPct: st.riskPct,
      overPct: st.overPct,
      sinMovInv: st.sinMovInv,
      perBranch: [{ branchLabel: b.label, ...st }]
    };
  }

  // ALL: compute per branch and also global combined
  let invAll=0, promAll=0;
  const covAll = [];
  let riskAll=0, overAll=0, sinMovInvAll=0;
  const skuCountAll = rows.length;

  const perBranch = BRANCHES.map(b=>{
    const st = computeBranchStats(b.key, rows);
    invAll += st.invSum;
    promAll += st.promSum;
    if (st.covMed>0) covAll.push(st.covMed);
    riskAll += st.riskCount;
    overAll += st.overCount;
    sinMovInvAll += st.sinMovInv;

    return { branchLabel: b.label, ...st };
  });

  const covMedGlobal = median(covAll);

  // riskPct global basado en skus totales * sucursales? NO. Usaremos promedio simple de sucursales (más útil ejecutivo).
  const riskPct = perBranch.length ? perBranch.reduce((a,x)=>a+x.riskPct,0)/perBranch.length : 0;
  const overPct = perBranch.length ? perBranch.reduce((a,x)=>a+x.overPct,0)/perBranch.length : 0;

  return {
    monthsUsed,
    skus: skuCountAll,
    inv: invAll,
    prom: promAll,
    covMed: covMedGlobal,
    riskPct, overPct,
    sinMovInv: sinMovInvAll,
    perBranch
  };
}

/* ===================== CHARTS ===================== */
function destroyChart(ch){
  if (ch && typeof ch.destroy === "function") ch.destroy();
}

function renderSummaryCharts(perBranch){
  // stacked classification 100% (by sku count)
  const labels = perBranch.map(x=>x.branchLabel);
  const A = perBranch.map(x => x.skus ? (x.clsCount.A/x.skus)*100 : 0);
  const B = perBranch.map(x => x.skus ? (x.clsCount.B/x.skus)*100 : 0);
  const C = perBranch.map(x => x.skus ? (x.clsCount.C/x.skus)*100 : 0);
  const S = perBranch.map(x => x.skus ? (x.clsCount["Sin Mov"]/x.skus)*100 : 0);

  destroyChart(state.charts.stacked);
  state.charts.stacked = new Chart($("chartStacked"), {
    type:"bar",
    data:{
      labels,
      datasets:[
        { label:"A", data:A, stack:"s" },
        { label:"B", data:B, stack:"s" },
        { label:"C", data:C, stack:"s" },
        { label:"Sin Mov", data:S, stack:"s" },
      ]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      scales:{
        x:{ stacked:true, ticks:{ color:"#9fb1d1" }, grid:{ color:"rgba(255,255,255,.06)" } },
        y:{ stacked:true, ticks:{ color:"#9fb1d1", callback:(v)=>v+"%" }, grid:{ color:"rgba(255,255,255,.06)" }, max:100 }
      },
      plugins:{
        legend:{ labels:{ color:"#e9f0ff" } },
        tooltip:{ callbacks:{ label:(ctx)=>`${ctx.dataset.label}: ${ctx.parsed.y.toFixed(1)}%` } }
      }
    }
  });

  // risk/over
  const risk = perBranch.map(x => x.riskPct*100);
  const over = perBranch.map(x => x.overPct*100);

  destroyChart(state.charts.riskOver);
  state.charts.riskOver = new Chart($("chartRiskOver"), {
    type:"bar",
    data:{
      labels,
      datasets:[
        { label:"Riesgo (<15)", data:risk },
        { label:"Sobreinv (>60)", data:over }
      ]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      scales:{
        x:{ ticks:{ color:"#9fb1d1" }, grid:{ color:"rgba(255,255,255,.06)" } },
        y:{ ticks:{ color:"#9fb1d1", callback:(v)=>v+"%" }, grid:{ color:"rgba(255,255,255,.06)" } }
      },
      plugins:{ legend:{ labels:{ color:"#e9f0ff" } } }
    }
  });
}

function renderDetailChart(distBuckets){
  destroyChart(state.charts.dist);
  state.charts.dist = new Chart($("chartHist"), {
    type:"bar",
    data:{
      labels: distBuckets.map(b=>b.label),
      datasets:[{ label:"SKUs", data: distBuckets.map(b=>b.count) }]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      scales:{
        x:{ ticks:{ color:"#9fb1d1" }, grid:{ color:"rgba(255,255,255,.06)" } },
        y:{ ticks:{ color:"#9fb1d1" }, grid:{ color:"rgba(255,255,255,.06)" } },
      },
      plugins:{ legend:{ labels:{ color:"#e9f0ff" } } }
    }
  });
}

/* ===================== RENDER SUMMARY ===================== */
function renderHeatTable(perBranch){
  const table = $("heatTable");
  const columns = [
    { key:"branchLabel", label:"Sucursal" },
    { key:"skus", label:"SKUs" , fmt:(v)=>formatInt(v)},
    { key:"invSum", label:"Inv (pzs)", fmt:(v)=>formatInt(v) },
    { key:"covMed", label:"Cob Med (días)", fmt:(v)=>formatNoDecimals(v) },
    { key:"A", label:"A (SKUs / pzs)", fmt:(v,r)=>`${formatInt(r.clsCount.A)} | ${formatInt(r.clsInv.A)}` },
    { key:"B", label:"B (SKUs / pzs)", fmt:(v,r)=>`${formatInt(r.clsCount.B)} | ${formatInt(r.clsInv.B)}` },
    { key:"C", label:"C (SKUs / pzs)", fmt:(v,r)=>`${formatInt(r.clsCount.C)} | ${formatInt(r.clsInv.C)}` },
    { key:"S", label:"Sin Mov (SKUs / pzs)", fmt:(v,r)=>`${formatInt(r.clsCount["Sin Mov"])} | ${formatInt(r.clsInv["Sin Mov"])}` },
    { key:"riskPct", label:"Riesgo %", fmt:(v)=>pct(v) },
    { key:"overPct", label:"Sobreinv %", fmt:(v)=>pct(v) },
  ];

  const rows = perBranch.map(x=>({
    branchLabel:x.branchLabel,
    skus:x.skus,
    invSum:x.invSum,
    covMed:x.covMed,
    clsCount:x.clsCount,
    clsInv:x.clsInv,
    riskPct:x.riskPct,
    overPct:x.overPct,
  }));

  renderTable(table, columns, rows, null);
}

function renderRankings(perBranch){
  const riskRank = [...perBranch].sort((a,b)=>b.riskPct-a.riskPct).slice(0,5);
  const overRank = [...perBranch].sort((a,b)=>b.overPct-a.overPct).slice(0,5);

  $("rankRisk").innerHTML = riskRank.map(x=>`<li>${x.branchLabel}: <b>${pct(x.riskPct)}</b> (SKUs: ${formatInt(x.riskCount)})</li>`).join("");
  $("rankOver").innerHTML = overRank.map(x=>`<li>${x.branchLabel}: <b>${pct(x.overPct)}</b> (SKUs: ${formatInt(x.overCount)})</li>`).join("");
}

function renderInsights(global){
  const items = [];
  if (!global.perBranch.length){
    $("insights").innerHTML = "";
    return;
  }

  if (global.perBranch.length === 1){
    const b = global.perBranch[0];
    items.push(`Sucursal <b>${b.branchLabel}</b>: Riesgo <b>${pct(b.riskPct)}</b> • Sobreinv <b>${pct(b.overPct)}</b> • Sin Mov con inv <b>${formatInt(b.sinMovInv)}</b>.`);
    items.push(`Cobertura mediana <b>${formatNoDecimals(b.covMed)} días</b> con inventario total <b>${formatInt(b.invSum)}</b> piezas.`);
  } else {
    const topRisk = [...global.perBranch].sort((a,b)=>b.riskPct-a.riskPct)[0];
    const topOver = [...global.perBranch].sort((a,b)=>b.overPct-a.overPct)[0];
    items.push(`Mayor Riesgo (<15 días): <b>${topRisk.branchLabel}</b> con <b>${pct(topRisk.riskPct)}</b>.`);
    items.push(`Mayor Sobreinventario (>60 días): <b>${topOver.branchLabel}</b> con <b>${pct(topOver.overPct)}</b>.`);
    items.push(`Cobertura global (mediana de sucursales): <b>${formatNoDecimals(global.covMed)} días</b>.`);
  }

  $("insights").innerHTML = items.map(t=>`<li>${t}</li>`).join("");
}

function renderSummary(){
  if (!state.loaded) return;

  const selected = $("summaryWarehouseSelect").value || "ALL";
  const g = computeGlobal(selected);

  $("gMeses").textContent = String(g.monthsUsed || 0);
  $("gSkus").textContent = formatInt(g.skus || 0);
  $("gInv").textContent = formatInt(g.inv || 0);
  $("gCobDias").textContent = formatNoDecimals(g.covMed || 0);
  $("gPromMes").textContent = formatProm(g.prom || 0);
  $("gRiskPct").textContent = pct(g.riskPct || 0);
  $("gOverPct").textContent = pct(g.overPct || 0);
  $("gSinMovInv").textContent = formatInt(g.sinMovInv || 0);

  renderSummaryCharts(g.perBranch);
  renderHeatTable(g.perBranch);
  renderRankings(g.perBranch);
  renderInsights(g);
}

/* ===================== DETAIL VIEW ===================== */
function getDetailFilters(){
  const branch = $("warehouseSelect").value;
  const cls = $("classSelect").value;
  const q = String($("searchInput").value||"").trim().toUpperCase();
  const onlyInv = $("onlyInvToggle").checked;
  const min = $("covMin").value ? toNumber($("covMin").value) : null;
  const max = $("covMax").value ? toNumber($("covMax").value) : null;
  const onlySinMovInv = $("onlySinMovInv").checked;
  const onlyABZero = $("onlyABZero").checked;

  return { branch, cls, q, onlyInv, min, max, onlySinMovInv, onlyABZero };
}

function buildDetailRows(){
  const f = getDetailFilters();
  const b = BRANCHES.find(x=>x.label === f.branch) || BRANCHES[0];

  let rows = state.outputRows.map(sku=>{
    const bb = sku.byBranch[b.key];
    return {
      codigo: sku.codigo,
      desc: sku.desc,
      inv: bb.inv,
      cls: safeClass(bb.cls),
      prom: bb.prom,
      covMes: bb.covMes,
      covDays: bb.covDays,
      risk: isRisk(bb.covDays) ? 1 : 0,
      over: isOver(bb.covDays) ? 1 : 0,
    };
  });

  if (f.q){
    rows = rows.filter(r => (r.codigo||"").toUpperCase().includes(f.q) || (r.desc||"").toUpperCase().includes(f.q));
  }
  if (f.cls !== "ALL"){
    rows = rows.filter(r => r.cls === f.cls);
  }
  if (f.onlyInv){
    rows = rows.filter(r => r.inv > 0);
  }
  if (f.onlySinMovInv){
    rows = rows.filter(r => r.cls === "Sin Mov" && r.inv > 0);
  }
  if (f.onlyABZero){
    rows = rows.filter(r => (r.cls === "A" || r.cls === "B") && r.inv === 0);
  }
  if (f.min !== null){
    rows = rows.filter(r => toNumber(r.covDays) >= f.min);
  }
  if (f.max !== null){
    rows = rows.filter(r => toNumber(r.covDays) <= f.max);
  }

  // default sort by covDays desc
  const sortKey = state.detail.sortKey || "covDays";
  rows = sortRows(rows, sortKey, state.detail.sortDir || "desc");
  return rows;
}

function renderDetailKPIs(rows){
  const meses = state.monthsUsed || 0;
  $("kpiMeses").textContent = String(meses);
  $("kpiSkus").textContent = formatInt(rows.length);
  $("kpiInv").textContent = formatInt(rows.reduce((a,r)=>a+r.inv,0));

  const covArr = rows.map(r=>r.covDays).filter(x=>x>0 && isFinite(x));
  $("kpiCobMed").textContent = formatNoDecimals(median(covArr));

  $("kpiProm").textContent = formatProm(rows.reduce((a,r)=>a+r.prom,0));
  const risk = rows.filter(r=>isRisk(r.covDays)).length;
  const over = rows.filter(r=>isOver(r.covDays)).length;
  $("kpiRisk").textContent = rows.length ? pct(risk/rows.length) : "0%";
  $("kpiOver").textContent = rows.length ? pct(over/rows.length) : "0%";

  const sinMovInv = rows.filter(r=>r.cls==="Sin Mov" && r.inv>0).length;
  $("kpiSinMovInv").textContent = formatInt(sinMovInv);

  $("kpiA").textContent = formatInt(rows.filter(r=>r.cls==="A").length);
  $("kpiB").textContent = formatInt(rows.filter(r=>r.cls==="B").length);
  $("kpiC").textContent = formatInt(rows.filter(r=>r.cls==="C").length);
  $("kpiS").textContent = formatInt(rows.filter(r=>r.cls==="Sin Mov").length);
}

function renderDetailTop(rows){
  const over = [...rows].sort((a,b)=>b.covDays-a.covDays).slice(0,15);
  const risk = [...rows].filter(r=>r.covDays>0).sort((a,b)=>a.covDays-b.covDays).slice(0,15);

  $("topOver").innerHTML = over.map(r=>`<li><b>${r.codigo}</b> — ${r.desc} <span class="muted">(${formatNoDecimals(r.covDays)} días, inv ${formatInt(r.inv)})</span></li>`).join("");
  $("topRisk").innerHTML = risk.map(r=>`<li><b>${r.codigo}</b> — ${r.desc} <span class="muted">(${formatNoDecimals(r.covDays)} días, inv ${formatInt(r.inv)})</span></li>`).join("");
}

function renderDetail(){
  if (!state.loaded) return;

  const rows = buildDetailRows();
  renderDetailKPIs(rows);

  const buckets = [
    { label:"0–15", min:0, max:15, count:0 },
    { label:"16–30", min:16, max:30, count:0 },
    { label:"31–60", min:31, max:60, count:0 },
    { label:"61–120", min:61, max:120, count:0 },
    { label:">120", min:121, max:1e12, count:0 },
  ];
  for (const r of rows){
    const d = toNumber(r.covDays);
    if (!d || d <= 0) continue;
    const b = buckets.find(x=>d>=x.min && d<=x.max) || buckets[buckets.length-1];
    b.count += 1;
  }
  renderDetailChart(buckets);
  renderDetailTop(rows);

  const cols = [
    { key:"codigo", label:"Código" },
    { key:"desc", label:"Descripción" },
    { key:"cls", label:"Clasificación" },
    { key:"inv", label:"Inv", fmt:(v)=>formatInt(v) },
    { key:"prom", label:"Prom Vta Mes", fmt:(v)=>formatProm(v) },
    { key:"covMes", label:"Cobertura (Mes)", fmt:(v)=>formatNoDecimals(v) },
    { key:"covDays", label:"Cobertura (Días)", fmt:(v)=>formatNoDecimals(v) },
  ];

  // pagination
  state.detail.pageSize = parseInt($("detailPageSize").value || "100",10);
  const pg = paginate(rows, state.detail.page, state.detail.pageSize);
  state.detail.page = pg.page;

  $("detailPageInfo").textContent = `Página ${pg.page} de ${pg.pages} • ${formatInt(pg.total)} filas`;
  $("detailPrev").disabled = pg.page <= 1;
  $("detailNext").disabled = pg.page >= pg.pages;

  const onSort = (key)=>{
    if (state.detail.sortKey === key){
      state.detail.sortDir = state.detail.sortDir === "asc" ? "desc" : "asc";
    } else {
      state.detail.sortKey = key;
      state.detail.sortDir = "asc";
    }
    state.detail.page = 1;
    renderDetail();
  };

  renderTable($("dataTable"), cols, pg.slice, onSort);

  $("tableHint").textContent = `Sucursal ${$("warehouseSelect").value} • Filtros aplicados`;

  // export current filtered rows (no pagination)
  $("btnExport").onclick = ()=> downloadCSV(`detalle_${$("warehouseSelect").value}.csv`, cols, rows);
}

/* ===================== FINDER VIEW ===================== */
function buildFinderRows(){
  const q = String($("finderSearch").value||"").trim().toUpperCase();
  const onlyInv = $("finderOnlyInv").checked;

  let rows = state.outputRows.map(sku=>{
    const r = { codigo: sku.codigo, desc: sku.desc };
    let anyInv = false;
    for (const b of BRANCHES){
      const bb = sku.byBranch[b.key];
      r[`cls_${b.key}`] = safeClass(bb.cls);
      r[`inv_${b.key}`] = bb.inv;
      if (bb.inv > 0) anyInv = true;
    }
    r.anyInv = anyInv ? 1 : 0;
    return r;
  });

  if (q){
    rows = rows.filter(r => r.codigo.toUpperCase().includes(q) || r.desc.toUpperCase().includes(q));
  }
  if (onlyInv){
    rows = rows.filter(r => r.anyInv === 1);
  }

  const sortKey = state.finder.sortKey || "codigo";
  rows = sortRows(rows, sortKey, state.finder.sortDir || "asc");
  return rows;
}

function renderFinder(){
  const rows = buildFinderRows();

  const cols = [
    { key:"codigo", label:"Código" },
    { key:"desc", label:"Descripción" },
    ...BRANCHES.flatMap(b=>[
      { key:`cls_${b.key}`, label:`Clasif ${b.label}` },
      { key:`inv_${b.key}`, label:`Inv ${b.label}`, fmt:(v)=>formatInt(v) },
    ]),
  ];

  state.finder.pageSize = parseInt($("finderPageSize").value || "100",10);
  const pg = paginate(rows, state.finder.page, state.finder.pageSize);
  state.finder.page = pg.page;

  $("finderPageInfo").textContent = `Página ${pg.page} de ${pg.pages} • ${formatInt(pg.total)} filas`;
  $("finderPrev").disabled = pg.page <= 1;
  $("finderNext").disabled = pg.page >= pg.pages;

  const onSort = (key)=>{
    if (state.finder.sortKey === key){
      state.finder.sortDir = state.finder.sortDir === "asc" ? "desc" : "asc";
    } else {
      state.finder.sortKey = key;
      state.finder.sortDir = "asc";
    }
    state.finder.page = 1;
    renderFinder();
  };

  renderTable($("finderTable"), cols, pg.slice, onSort);

  $("finderExport").onclick = ()=> downloadCSV("buscador_sku.csv", cols, rows);
}

/* ===================== OPORTUNIDADES (QUiebres) ===================== */
function computeOppBreak(targetLabel, onlyRisk){
  const target = BRANCHES.find(b=>b.label===targetLabel) || BRANCHES[0];

  const out = [];

  for (const sku of state.outputRows){
    const t = sku.byBranch[target.key];
    const tCls = safeClass(t.cls);
    if (!(tCls==="A" || tCls==="B")) continue;

    const tCov = toNumber(t.covDays);
    const tRisk = isRisk(tCov) || t.inv === 0;

    if (onlyRisk && !tRisk) continue;

    // deficit to reach TARGET_RISK_DAYS
    const daily = (t.prom > 0) ? (t.prom/30) : 0;
    if (daily <= 0) continue; // sin velocidad de venta no sugerimos

    const required = daily * TARGET_RISK_DAYS;
    const deficit = Math.max(0, required - t.inv);
    if (deficit <= 0) continue;

    const donors = [];
    for (const b of BRANCHES){
      if (b.key === target.key) continue;
      const bb = sku.byBranch[b.key];
      const c = safeClass(bb.cls);
      if ((c==="C" || c==="Sin Mov") && bb.inv > 0){
        donors.push({
          branch: b.label,
          inv: bb.inv,
          cls: c,
          covDays: bb.covDays
        });
      }
    }
    if (!donors.length) continue;

    // allocate from donors (simple: highest inv first)
    donors.sort((a,b)=>b.inv-a.inv);
    let remaining = deficit;
    const alloc = {};
    for (const d of donors){
      if (remaining <= 0) break;
      const take = Math.min(d.inv, remaining);
      alloc[d.branch] = take;
      remaining -= take;
    }
    const suggested = deficit - remaining;

    out.push({
      codigo: sku.codigo,
      desc: sku.desc,
      targetSucursal: target.label,
      targetCls: tCls,
      targetInv: t.inv,
      targetCov: tCov,
      targetProm: t.prom,
      deficit: deficit,
      sugerido: suggested,
      ...donors.reduce((acc,d)=>{
        acc[`don_cls_${d.branch}`] = d.cls;
        acc[`don_inv_${d.branch}`] = d.inv;
        acc[`don_take_${d.branch}`] = alloc[d.branch] || 0;
        return acc;
      },{})
    });
  }

  // sort: highest deficit first
  return out.sort((a,b)=>b.deficit-a.deficit);
}

function renderOppBreak(){
  const target = $("oppBreakTarget").value;
  const onlyRisk = $("oppBreakOnlyRisk").checked;

  const rows = computeOppBreak(target, onlyRisk);

  // columns
  const donorBranches = BRANCHES.filter(b=>b.label!==target);
  const cols = [
    { key:"codigo", label:"Código" },
    { key:"desc", label:"Descripción" },
    { key:"targetCls", label:"Cls Obj" },
    { key:"targetInv", label:"Inv Obj", fmt:(v)=>formatInt(v) },
    { key:"targetCov", label:"Cob Obj (días)", fmt:(v)=>formatNoDecimals(v) },
    { key:"deficit", label:"Déficit (pzs)", fmt:(v)=>formatInt(v) },
    { key:"sugerido", label:"Sugerido traslado", fmt:(v)=>formatInt(v) },
    ...donorBranches.flatMap(b=>[
      { key:`don_cls_${b.label}`, label:`Cls ${b.label}` },
      { key:`don_inv_${b.label}`, label:`Inv ${b.label}`, fmt:(v)=>formatInt(v) },
      { key:`don_take_${b.label}`, label:`Tomar ${b.label}`, fmt:(v)=>formatInt(v) },
    ]),
  ];

  state.oppBreak.pageSize = parseInt($("oppBreakPageSize").value || "100",10);
  const pg = paginate(rows, state.oppBreak.page, state.oppBreak.pageSize);
  state.oppBreak.page = pg.page;

  $("oppBreakPageInfo").textContent = `Página ${pg.page} de ${pg.pages} • ${formatInt(pg.total)} filas`;
  $("oppBreakPrev").disabled = pg.page <= 1;
  $("oppBreakNext").disabled = pg.page >= pg.pages;

  const onSort = (key)=>{
    if (state.oppBreak.sortKey === key){
      state.oppBreak.sortDir = state.oppBreak.sortDir === "asc" ? "desc" : "asc";
    } else {
      state.oppBreak.sortKey = key;
      state.oppBreak.sortDir = "asc";
    }
    state.oppBreak.page = 1;
    const sorted = sortRows(rows, state.oppBreak.sortKey, state.oppBreak.sortDir);
    const pg2 = paginate(sorted, state.oppBreak.page, state.oppBreak.pageSize);
    renderTable($("oppBreakTable"), cols, pg2.slice, onSort);
  };

  renderTable($("oppBreakTable"), cols, pg.slice, onSort);

  $("oppBreakExport").onclick = ()=> downloadCSV(`oportunidades_quiebres_${target}.csv`, cols, rows);
}

/* ===================== OPORTUNIDADES (Sobreinv) ===================== */
function computeOppOver(targetLabel, onlyOver){
  const target = BRANCHES.find(b=>b.label===targetLabel) || BRANCHES[0];
  const out = [];

  for (const sku of state.outputRows){
    const t = sku.byBranch[target.key];
    const tCls = safeClass(t.cls);
    if (!(tCls==="C" || tCls==="Sin Mov")) continue;

    const tCov = toNumber(t.covDays);
    const isOverHere = isOver(tCov);

    if (onlyOver && !isOverHere) continue;
    if (t.inv <= 0) continue;

    // exceso objetivo: inv - daily*TARGET_OVER_DAYS
    const daily = (t.prom > 0) ? (t.prom/30) : 0;
    const excess = daily > 0 ? Math.max(0, t.inv - daily*TARGET_OVER_DAYS) : t.inv;
    if (excess <= 0) continue;

    // receptores: sucursales con A/B (idealmente riesgo)
    const receivers = [];
    for (const b of BRANCHES){
      if (b.key === target.key) continue;
      const bb = sku.byBranch[b.key];
      const c = safeClass(bb.cls);
      if (c==="A" || c==="B"){
        const dailyR = (bb.prom > 0) ? (bb.prom/30) : 0;
        const deficitR = dailyR > 0 ? Math.max(0, dailyR*TARGET_RISK_DAYS - bb.inv) : 0;
        receivers.push({
          branch: b.label,
          cls: c,
          inv: bb.inv,
          covDays: bb.covDays,
          deficit: deficitR
        });
      }
    }
    if (!receivers.length) continue;

    // allocate excess to highest deficit first
    receivers.sort((a,b)=>b.deficit-a.deficit);
    let remaining = excess;
    const alloc = {};
    for (const r of receivers){
      if (remaining <= 0) break;
      const take = Math.min(remaining, Math.max(0, r.deficit || 0));
      if (take > 0){
        alloc[r.branch] = take;
        remaining -= take;
      }
    }
    const suggested = excess - remaining;

    out.push({
      codigo: sku.codigo,
      desc: sku.desc,
      targetSucursal: target.label,
      targetCls: tCls,
      targetInv: t.inv,
      targetCov: tCov,
      excess: excess,
      sugerido: suggested,
      ...receivers.reduce((acc,r)=>{
        acc[`rec_cls_${r.branch}`] = r.cls;
        acc[`rec_inv_${r.branch}`] = r.inv;
        acc[`rec_need_${r.branch}`] = r.deficit || 0;
        acc[`rec_send_${r.branch}`] = alloc[r.branch] || 0;
        return acc;
      },{})
    });
  }

  return out.sort((a,b)=>b.excess-a.excess);
}

function renderOppOver(){
  const target = $("oppOverTarget").value;
  const onlyOver = $("oppOverOnlyOver").checked;

  const rows = computeOppOver(target, onlyOver);

  const receiverBranches = BRANCHES.filter(b=>b.label!==target);
  const cols = [
    { key:"codigo", label:"Código" },
    { key:"desc", label:"Descripción" },
    { key:"targetCls", label:"Cls Obj" },
    { key:"targetInv", label:"Inv Obj", fmt:(v)=>formatInt(v) },
    { key:"targetCov", label:"Cob Obj (días)", fmt:(v)=>formatNoDecimals(v) },
    { key:"excess", label:"Exceso (pzs)", fmt:(v)=>formatInt(v) },
    { key:"sugerido", label:"Sugerido mover", fmt:(v)=>formatInt(v) },
    ...receiverBranches.flatMap(b=>[
      { key:`rec_cls_${b.label}`, label:`Cls ${b.label}` },
      { key:`rec_inv_${b.label}`, label:`Inv ${b.label}`, fmt:(v)=>formatInt(v) },
      { key:`rec_need_${b.label}`, label:`Necesita ${b.label}`, fmt:(v)=>formatInt(v) },
      { key:`rec_send_${b.label}`, label:`Enviar ${b.label}`, fmt:(v)=>formatInt(v) },
    ]),
  ];

  state.oppOver.pageSize = parseInt($("oppOverPageSize").value || "100",10);
  const pg = paginate(rows, state.oppOver.page, state.oppOver.pageSize);
  state.oppOver.page = pg.page;

  $("oppOverPageInfo").textContent = `Página ${pg.page} de ${pg.pages} • ${formatInt(pg.total)} filas`;
  $("oppOverPrev").disabled = pg.page <= 1;
  $("oppOverNext").disabled = pg.page >= pg.pages;

  const onSort = (key)=>{
    if (state.oppOver.sortKey === key){
      state.oppOver.sortDir = state.oppOver.sortDir === "asc" ? "desc" : "asc";
    } else {
      state.oppOver.sortKey = key;
      state.oppOver.sortDir = "asc";
    }
    state.oppOver.page = 1;
    const sorted = sortRows(rows, state.oppOver.sortKey, state.oppOver.sortDir);
    const pg2 = paginate(sorted, state.oppOver.page, state.oppOver.pageSize);
    renderTable($("oppOverTable"), cols, pg2.slice, onSort);
  };

  renderTable($("oppOverTable"), cols, pg.slice, onSort);

  $("oppOverExport").onclick = ()=> downloadCSV(`oportunidades_sobreinv_${target}.csv`, cols, rows);
}

/* ===================== HISTORICOS ===================== */
function filterHist(rows, sucursalSel, text, from, to){
  const q = String(text||"").trim().toUpperCase();
  const f = String(from||"").trim();
  const t = String(to||"").trim();

  return rows.filter(r=>{
    if (sucursalSel && sucursalSel !== "ALL" && r.sucursal !== sucursalSel) return false;
    if (q && !r.item.toUpperCase().includes(q)) return false;
    if (f && r.mes < f) return false;
    if (t && r.mes > t) return false;
    return true;
  });
}

function buildPivot(rows){
  // rows: {mes, sucursal, monto, litros} already filtered
  const meses = [...new Set(rows.map(r=>r.mes))].sort();
  const sucs = [...new Set(rows.map(r=>r.sucursal))].sort((a,b)=>a.localeCompare(b,"es",{sensitivity:"base"}));

  const money = {};
  const litros = {};

  for (const m of meses){
    money[m] = {};
    litros[m] = {};
    for (const s of sucs){
      money[m][s] = 0;
      litros[m][s] = 0;
    }
  }

  for (const r of rows){
    if (!money[r.mes]) continue;
    money[r.mes][r.sucursal] += toNumber(r.monto);
    litros[r.mes][r.sucursal] += toNumber(r.litros);
  }

  return { meses, sucs, money, litros };
}

function renderPivotTable(tableEl, pivot, kind /* "money"|"litros" */){
  const { meses, sucs } = pivot;
  const data = kind === "money" ? pivot.money : pivot.litros;

  const thead = tableEl.querySelector("thead");
  const tbody = tableEl.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const trh = document.createElement("tr");
  const th0 = document.createElement("th");
  th0.textContent = "Mes";
  trh.appendChild(th0);
  for (const s of sucs){
    const th = document.createElement("th");
    th.textContent = s;
    trh.appendChild(th);
  }
  const thT = document.createElement("th");
  thT.textContent = "TOTAL";
  trh.appendChild(thT);
  thead.appendChild(trh);

  const frag = document.createDocumentFragment();
  for (const m of meses){
    const tr = document.createElement("tr");
    const td0 = document.createElement("td");
    td0.textContent = m;
    tr.appendChild(td0);

    let total = 0;
    for (const s of sucs){
      const v = toNumber(data[m][s] || 0);
      total += v;
      const td = document.createElement("td");
      td.textContent = kind === "money" ? formatMoney(v) : formatInt(v);
      tr.appendChild(td);
    }
    const tdT = document.createElement("td");
    tdT.textContent = kind === "money" ? formatMoney(total) : formatInt(total);
    tr.appendChild(tdT);

    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
}

function renderHistDetail(tableEl, rows, mode, onSort){
  const cols = [
    { key:"mes", label:"Mes" },
    { key:"sucursal", label:"Sucursal" },
    { key:"item", label: mode === "CATEG" ? "Categoría" : "Marca" },
    { key:"monto", label:"Monto_venta", fmt:(v)=>formatMoney(v) },
    { key:"litros", label:"Litros Vendidos", fmt:(v)=>formatInt(v) },
  ];
  renderTable(tableEl, cols, rows, onSort);
  return cols;
}

function renderHistCateg(){
  const suc = $("histCategSucursal").value;
  const q = $("histCategSearch").value;
  const from = $("histCategFrom").value;
  const to = $("histCategTo").value;

  const filtered = filterHist(state.histCateg, suc, q, from, to);

  // pivots (dependen del filtro)
  const pivot = buildPivot(filtered);
  renderPivotTable($("histCategPivotMoney"), pivot, "money");
  renderPivotTable($("histCategPivotLitros"), pivot, "litros");

  // detalle + pagination
  const sortKey = state.histCategUI.sortKey || "monto";
  const sorted = sortRows(filtered.map(r=>({ ...r })), sortKey, state.histCategUI.sortDir || "desc");

  state.histCategUI.pageSize = parseInt($("histCategPageSize").value || "100",10);
  const pg = paginate(sorted, state.histCategUI.page, state.histCategUI.pageSize);
  state.histCategUI.page = pg.page;

  $("histCategPageInfo").textContent = `Página ${pg.page} de ${pg.pages} • ${formatInt(pg.total)} filas`;
  $("histCategPrev").disabled = pg.page <= 1;
  $("histCategNext").disabled = pg.page >= pg.pages;

  const onSort = (key)=>{
    if (state.histCategUI.sortKey === key){
      state.histCategUI.sortDir = state.histCategUI.sortDir === "asc" ? "desc" : "asc";
    } else {
      state.histCategUI.sortKey = key;
      state.histCategUI.sortDir = "asc";
    }
    state.histCategUI.page = 1;
    renderHistCateg();
  };

  const cols = renderHistDetail($("histCategDetail"), pg.slice, "CATEG", onSort);

  $("histCategExport").onclick = ()=> downloadCSV("historico_categorias_detalle.csv", cols, filtered);
}

function renderHistMarca(){
  const suc = $("histMarcaSucursal").value;
  const q = $("histMarcaSearch").value;
  const from = $("histMarcaFrom").value;
  const to = $("histMarcaTo").value;

  const filtered = filterHist(state.histMarca, suc, q, from, to);

  const pivot = buildPivot(filtered);
  renderPivotTable($("histMarcaPivotMoney"), pivot, "money");
  renderPivotTable($("histMarcaPivotLitros"), pivot, "litros");

  const sortKey = state.histMarcaUI.sortKey || "monto";
  const sorted = sortRows(filtered.map(r=>({ ...r })), sortKey, state.histMarcaUI.sortDir || "desc");

  state.histMarcaUI.pageSize = parseInt($("histMarcaPageSize").value || "100",10);
  const pg = paginate(sorted, state.histMarcaUI.page, state.histMarcaUI.pageSize);
  state.histMarcaUI.page = pg.page;

  $("histMarcaPageInfo").textContent = `Página ${pg.page} de ${pg.pages} • ${formatInt(pg.total)} filas`;
  $("histMarcaPrev").disabled = pg.page <= 1;
  $("histMarcaNext").disabled = pg.page >= pg.pages;

  const onSort = (key)=>{
    if (state.histMarcaUI.sortKey === key){
      state.histMarcaUI.sortDir = state.histMarcaUI.sortDir === "asc" ? "desc" : "asc";
    } else {
      state.histMarcaUI.sortKey = key;
      state.histMarcaUI.sortDir = "asc";
    }
    state.histMarcaUI.page = 1;
    renderHistMarca();
  };

  const cols = renderHistDetail($("histMarcaDetail"), pg.slice, "MARCA", onSort);

  $("histMarcaExport").onclick = ()=> downloadCSV("historico_marcas_detalle.csv", cols, filtered);
}

/* ===================== EVENTS ===================== */
function bindDetailEvents(){
  const rerender = ()=> { state.detail.page = 1; renderDetail(); };
  $("warehouseSelect").addEventListener("change", rerender);
  $("classSelect").addEventListener("change", rerender);
  $("searchInput").addEventListener("input", rerender);
  $("onlyInvToggle").addEventListener("change", rerender);
  $("covMin").addEventListener("input", rerender);
  $("covMax").addEventListener("input", rerender);
  $("onlySinMovInv").addEventListener("change", rerender);
  $("onlyABZero").addEventListener("change", rerender);

  $("detailPageSize").addEventListener("change", ()=>{ state.detail.page=1; renderDetail(); });
  $("detailPrev").addEventListener("click", ()=>{ state.detail.page--; renderDetail(); });
  $("detailNext").addEventListener("click", ()=>{ state.detail.page++; renderDetail(); });
}

function bindFinderEvents(){
  const rerender = ()=>{ state.finder.page=1; renderFinder(); };
  $("finderSearch").addEventListener("input", rerender);
  $("finderOnlyInv").addEventListener("change", rerender);
  $("finderPageSize").addEventListener("change", ()=>{ state.finder.page=1; renderFinder(); });
  $("finderPrev").addEventListener("click", ()=>{ state.finder.page--; renderFinder(); });
  $("finderNext").addEventListener("click", ()=>{ state.finder.page++; renderFinder(); });
}

function bindOppEvents(){
  $("oppBreakTarget").addEventListener("change", ()=>{ state.oppBreak.page=1; renderOppBreak(); });
  $("oppBreakOnlyRisk").addEventListener("change", ()=>{ state.oppBreak.page=1; renderOppBreak(); });
  $("oppBreakPageSize").addEventListener("change", ()=>{ state.oppBreak.page=1; renderOppBreak(); });
  $("oppBreakPrev").addEventListener("click", ()=>{ state.oppBreak.page--; renderOppBreak(); });
  $("oppBreakNext").addEventListener("click", ()=>{ state.oppBreak.page++; renderOppBreak(); });

  $("oppOverTarget").addEventListener("change", ()=>{ state.oppOver.page=1; renderOppOver(); });
  $("oppOverOnlyOver").addEventListener("change", ()=>{ state.oppOver.page=1; renderOppOver(); });
  $("oppOverPageSize").addEventListener("change", ()=>{ state.oppOver.page=1; renderOppOver(); });
  $("oppOverPrev").addEventListener("click", ()=>{ state.oppOver.page--; renderOppOver(); });
  $("oppOverNext").addEventListener("click", ()=>{ state.oppOver.page++; renderOppOver(); });
}

function bindHistEvents(){
  const rerC = ()=>{ state.histCategUI.page=1; renderHistCateg(); };
  $("histCategSucursal").addEventListener("change", rerC);
  $("histCategSearch").addEventListener("input", rerC);
  $("histCategFrom").addEventListener("input", rerC);
  $("histCategTo").addEventListener("input", rerC);
  $("histCategPageSize").addEventListener("change", ()=>{ state.histCategUI.page=1; renderHistCateg(); });
  $("histCategPrev").addEventListener("click", ()=>{ state.histCategUI.page--; renderHistCateg(); });
  $("histCategNext").addEventListener("click", ()=>{ state.histCategUI.page++; renderHistCateg(); });

  const rerM = ()=>{ state.histMarcaUI.page=1; renderHistMarca(); };
  $("histMarcaSucursal").addEventListener("change", rerM);
  $("histMarcaSearch").addEventListener("input", rerM);
  $("histMarcaFrom").addEventListener("input", rerM);
  $("histMarcaTo").addEventListener("input", rerM);
  $("histMarcaPageSize").addEventListener("change", ()=>{ state.histMarcaUI.page=1; renderHistMarca(); });
  $("histMarcaPrev").addEventListener("click", ()=>{ state.histMarcaUI.page--; renderHistMarca(); });
  $("histMarcaNext").addEventListener("click", ()=>{ state.histMarcaUI.page++; renderHistMarca(); });
}

function bindSummaryEvents(){
  $("summaryWarehouseSelect").addEventListener("change", renderSummary);
}

function populateBranchSelects(){
  // detail view
  const ws = $("warehouseSelect");
  ws.innerHTML = "";
  for (const b of BRANCHES){
    const opt = document.createElement("option");
    opt.value = b.label;
    opt.textContent = b.label;
    ws.appendChild(opt);
  }
  ws.value = "EXPRESS"; // default (como tu screenshot) si existe
  if (![...ws.options].some(o=>o.value==="EXPRESS")) ws.value = BRANCHES[0].label;

  // summary executive selector
  const ss = $("summaryWarehouseSelect");
  // keep first "Todas"
  for (const b of BRANCHES){
    const opt = document.createElement("option");
    opt.value = b.label;
    opt.textContent = b.label;
    ss.appendChild(opt);
  }
  ss.value = "ALL";

  // opp selects
  const fill = (sel)=>{
    sel.innerHTML = "";
    for (const b of BRANCHES){
      const opt = document.createElement("option");
      opt.value = b.label;
      opt.textContent = b.label;
      sel.appendChild(opt);
    }
    sel.value = BRANCHES[0].label;
  };
  fill($("oppBreakTarget"));
  fill($("oppOverTarget"));

  // hist selects
  const fillHist = (sel)=>{
    // keep ALL
    for (const b of BRANCHES){
      const opt = document.createElement("option");
      opt.value = b.label;
      opt.textContent = b.label;
      sel.appendChild(opt);
    }
    sel.value = "ALL";
  };
  fillHist($("histCategSucursal"));
  fillHist($("histMarcaSucursal"));
}

function clearAll(){
  state.loaded = false;
  state.fileName = null;
  state.outputRows = [];
  state.histCateg = [];
  state.histMarca = [];
  state.monthsUsed = 0;

  setStatus("Carga un archivo para comenzar.", "");
  enableUI(false);

  destroyChart(state.charts.stacked);
  destroyChart(state.charts.riskOver);
  destroyChart(state.charts.dist);
  state.charts.stacked = null;
  state.charts.riskOver = null;
  state.charts.dist = null;

  setActiveTab("tabSummary","summaryView");
}

/* ===================== FILE LOAD ===================== */
async function handleFile(file){
  if (!file) return;

  setStatus("Procesando archivo…", "Leyendo XLSX/CSV localmente.");
  const t0 = performance.now();

  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:"array", cellDates:true });

  const outName = findSheetName(wb, SHEET_OUTPUT);
  const categName = findSheetName(wb, SHEET_CATEG);
  const marcaName = findSheetName(wb, SHEET_MARCA);

  if (!outName){
    setStatus("ERROR: No encontré la hoja OUTPUT.", "Asegúrate que tu archivo tenga una pestaña llamada OUTPUT.");
    return;
  }

  const outputRaw = sheetToRows(wb, outName);
  const parsed = parseOutput(outputRaw);

  state.outputRows = parsed.rows;
  state.monthsUsed = parsed.monthsUsed;

  // Históricos (si existen)
  if (categName){
    state.histCateg = parseHist(sheetToRows(wb, categName), "CATEG");
  } else {
    state.histCateg = [];
  }
  if (marcaName){
    state.histMarca = parseHist(sheetToRows(wb, marcaName), "MARCA");
  } else {
    state.histMarca = [];
  }

  state.loaded = true;
  state.fileName = file.name;

  const t1 = performance.now();
  const meta = [
    `Archivo: ${file.name}`,
    `SKUs: ${formatInt(state.outputRows.length)}`,
    `Hojas: ${outName}${categName?`, ${categName}`:""}${marcaName?`, ${marcaName}`:""}`,
    `Tiempo: ${formatNoDecimals((t1-t0)/1000)}s`
  ].join(" • ");

  setStatus("✅ Archivo cargado correctamente", `Sucursales detectadas: ${BRANCHES.map(b=>b.key.replace("_"," ")).join(", ")} • ${meta}`);

  enableUI(true);
  populateBranchSelects();

  // default renders
  renderSummary();
  renderDetail();
  renderFinder();
  renderOppBreak();
  renderOppOver();
  renderHistCateg();
  renderHistMarca();
}

/* ===================== INIT ===================== */
function init(){
  bindTabs();
  bindDetailEvents();
  bindFinderEvents();
  bindOppEvents();
  bindHistEvents();
  bindSummaryEvents();

  $("fileInput").addEventListener("change", (e)=>{
    const file = e.target.files && e.target.files[0];
    if (file) handleFile(file);
  });

  $("btnClear").addEventListener("click", ()=>{
    $("fileInput").value = "";
    clearAll();
  });

  clearAll();
}

document.addEventListener("DOMContentLoaded", init);
