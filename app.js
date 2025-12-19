/* ============================
   Maestro Inventarios - Client-side (GitHub Pages)
   - Carga XLSX/CSV
   - Resumen general: KPIs + stacked 100% + heatmap + rankings + insights
   - Detalle por sucursal: filtros + histograma + top lists + tabla
   - Umbrales: Riesgo < 15 días | Sobreinv > 60 días
============================ */

const el = (id) => document.getElementById(id);

// Thresholds (pedido)
const TH_RISK = 15;
const TH_OVER = 60;

// UI - common
const fileInput = el("fileInput");
const btnClear = el("btnClear");
const statusMsg = el("statusMsg");
const statusMeta = el("statusMeta");

// Tabs
const tabSummary = el("tabSummary");
const tabDetail = el("tabDetail");
const summaryView = el("summaryView");
const detailView = el("detailView");

// Summary KPIs
const gMeses = el("gMeses");
const gSkus = el("gSkus");
const gInv = el("gInv");
const gCobDias = el("gCobDias");
const gPromMes = el("gPromMes");
const gRiskPct = el("gRiskPct");
const gOverPct = el("gOverPct");
const gSinMovInv = el("gSinMovInv");

// Summary visuals + tables
const heatTable = el("heatTable");
const heatHead = heatTable.querySelector("thead");
const heatBody = heatTable.querySelector("tbody");
const rankRisk = el("rankRisk");
const rankOver = el("rankOver");
const insights = el("insights");

// Detail UI
const btnExport = el("btnExport");
const warehouseSelect = el("warehouseSelect");
const classSelect = el("classSelect");
const searchInput = el("searchInput");
const onlyInvToggle = el("onlyInvToggle");
const covMin = el("covMin");
const covMax = el("covMax");
const onlySinMovInv = el("onlySinMovInv");
const onlyABZero = el("onlyABZero");

const tableHint = el("tableHint");
const dataTable = el("dataTable");
const thead = dataTable.querySelector("thead");
const tbody = dataTable.querySelector("tbody");

// Detail KPIs
const kpiMeses = el("kpiMeses");
const kpiSkus = el("kpiSkus");
const kpiInv = el("kpiInv");
const kpiCobMed = el("kpiCobMed");
const kpiProm = el("kpiProm");
const kpiRisk = el("kpiRisk");
const kpiOver = el("kpiOver");
const kpiSinMovInv = el("kpiSinMovInv");
const kpiA = el("kpiA");
const kpiB = el("kpiB");
const kpiC = el("kpiC");
const kpiS = el("kpiS");

// Detail tops
const topOver = el("topOver");
const topRisk = el("topRisk");

// Charts
let chartStacked = null;
let chartRiskOver = null;
let chartHist = null;

// Data state
let MASTER = [];        // array of rows (objects)
let WAREHOUSES = [];    // list of warehouse keys as found in columns (e.g. "adelitas", "express"...)
let MONTHS_USED = null; // from MesesUsados (first row)
let sortKey = null;
let sortDir = "asc"; // asc|desc

// ======= Helpers =======
function setStatus(msg, meta=""){
  statusMsg.textContent = msg;
  statusMeta.textContent = meta;
}

function setEnabled(enabled){
  btnClear.disabled = !enabled;
  tabSummary.disabled = !enabled;
  tabDetail.disabled = !enabled;

  btnExport.disabled = !enabled;
  warehouseSelect.disabled = !enabled;
  classSelect.disabled = !enabled;
  searchInput.disabled = !enabled;
  onlyInvToggle.disabled = !enabled;
  covMin.disabled = !enabled;
  covMax.disabled = !enabled;
  onlySinMovInv.disabled = !enabled;
  onlyABZero.disabled = !enabled;
}

function fmtNum(n, dec=0){
  if (n === null || n === undefined || n === "") return "—";
  const x = Number(n);
  if (!Number.isFinite(x)) return "—";
  return x.toLocaleString("es-MX", { maximumFractionDigits: dec, minimumFractionDigits: dec });
}

function pct(n, d, dec=1){
  if (!d) return "0%";
  return (100 * (n/d)).toLocaleString("es-MX", { maximumFractionDigits: dec, minimumFractionDigits: dec }) + "%";
}

function safeNum(x){
  const n = Number(x);
  return Number.isFinite(n) ? n : 0;
}

function median(nums){
  const arr = nums.filter(n => Number.isFinite(n)).slice().sort((a,b)=>a-b);
  if (!arr.length) return null;
  const mid = Math.floor(arr.length/2);
  return arr.length % 2 ? arr[mid] : (arr[mid-1] + arr[mid]) / 2;
}

function debounce(fn, ms){
  let t = null;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

function colMap(wh){
  return {
    inv: `Inv-${wh}`,
    cls: `Clasificacion-${wh}`,
    prom: `Promedio Vta Mes-${wh}`,
    cobMes: `Cobertura (Mes)-${wh}`,
    cobDias: `Cobertura Dias (30) -${wh}`,
  };
}

function normalizeColumns(rows){
  return rows.map(r => {
    const out = {};
    Object.keys(r).forEach(k => out[String(k).trim()] = r[k]);
    return out;
  });
}

// ======= Parsing =======
function splitCSVLine(line){
  const out = [];
  let cur = "";
  let inQ = false;
  for (let i=0; i<line.length; i++){
    const ch = line[i];
    if (ch === '"'){
      if (inQ && line[i+1] === '"'){ cur += '"'; i++; }
      else inQ = !inQ;
    } else if (ch === "," && !inQ){
      out.push(cur);
      cur = "";
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out;
}

function parseCSV(text){
  const lines = text.split(/\r?\n/).filter(l => l.trim().length);
  if (!lines.length) return [];
  const headers = splitCSVLine(lines[0]).map(h => h.trim());
  const rows = [];
  for (let i=1; i<lines.length; i++){
    const vals = splitCSVLine(lines[i]);
    const obj = {};
    headers.forEach((h, idx) => obj[h] = vals[idx] ?? "");
    rows.push(obj);
  }
  return rows;
}

function parseXLSX(arrayBuffer){
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

// ======= Detection/validation =======
function detectWarehouses(rows){
  if (!rows.length) return [];
  const cols = Object.keys(rows[0]);
  const invCols = cols.filter(c => c.startsWith("Inv-"));
  return invCols.map(c => c.replace("Inv-", "").trim());
}

function validateMaster(rows){
  if (!rows.length) throw new Error("El archivo está vacío.");
  const cols = Object.keys(rows[0]);

  for (const c of ["Codigo", "desc_prod"]) {
    if (!cols.includes(c)) throw new Error(`Falta columna obligatoria: ${c}`);
  }

  const warehouses = detectWarehouses(rows);
  if (!warehouses.length) throw new Error("No encontré columnas tipo 'Inv-ALMACEN'.");

  for (const wh of warehouses){
    const m = colMap(wh);
    for (const col of Object.values(m)){
      if (!cols.includes(col)) throw new Error(`Falta columna '${col}' para sucursal '${wh}'.`);
    }
  }

  return warehouses;
}

function coerceTypes(rows, warehouses){
  const out = rows.map(r => ({...r}));

  MONTHS_USED = null;
  if ("MesesUsados" in out[0]){
    MONTHS_USED = safeNum(out[0]["MesesUsados"]) || null;
  }

  for (const r of out){
    r["Codigo"] = String(r["Codigo"] ?? "").trim();
    r["desc_prod"] = String(r["desc_prod"] ?? "").trim();

    for (const wh of warehouses){
      const m = colMap(wh);
      r[m.inv] = safeNum(r[m.inv]);
      r[m.prom] = safeNum(r[m.prom]);
      r[m.cobMes] = safeNum(r[m.cobMes]);
      r[m.cobDias] = safeNum(r[m.cobDias]);
      r[m.cls] = String(r[m.cls] ?? "").trim();
    }
  }
  return out;
}

function fillWarehouseSelect(warehouses){
  warehouseSelect.innerHTML = "";
  warehouses.forEach(wh => {
    const opt = document.createElement("option");
    opt.value = wh;
    opt.textContent = wh;
    warehouseSelect.appendChild(opt);
  });
  warehouseSelect.value = warehouses[0] || "";
}

// ======= Summary computations =======
function computeSummaryByWarehouse(){
  // returns map: wh -> metrics
  const map = {};
  const skuCount = MASTER.length; // each row is a SKU
  for (const wh of WAREHOUSES){
    const m = colMap(wh);

    const A = MASTER.filter(r => r[m.cls] === "A").length;
    const B = MASTER.filter(r => r[m.cls] === "B").length;
    const C = MASTER.filter(r => r[m.cls] === "C").length;
    const S = MASTER.filter(r => r[m.cls] === "Sin Mov").length;

    const invTotal = MASTER.reduce((acc,r)=>acc + safeNum(r[m.inv]), 0);
    const promTotal = MASTER.reduce((acc,r)=>acc + safeNum(r[m.prom]), 0);

    const covArr = MASTER.map(r => safeNum(r[m.cobDias])).filter(x => Number.isFinite(x) && x >= 0);
    const covMed = median(covArr);

    const riskCount = MASTER.filter(r => safeNum(r[m.cobDias]) > 0 && safeNum(r[m.cobDias]) < TH_RISK).length;
    const overCount = MASTER.filter(r => safeNum(r[m.cobDias]) > TH_OVER).length;

    const sinMovInv = MASTER.filter(r => r[m.cls] === "Sin Mov" && safeNum(r[m.inv]) > 0).length;

    // inventory pieces in risk/over bands (piezas)
    const riskInv = MASTER.reduce((acc,r)=> {
      const d = safeNum(r[m.cobDias]);
      if (d > 0 && d < TH_RISK) return acc + safeNum(r[m.inv]);
      return acc;
    }, 0);
    const overInv = MASTER.reduce((acc,r)=> {
      const d = safeNum(r[m.cobDias]);
      if (d > TH_OVER) return acc + safeNum(r[m.inv]);
      return acc;
    }, 0);

    map[wh] = {
      wh,
      skuCount,
      A,B,C,S,
      invTotal,
      promTotal,
      covMed,
      riskCount,
      overCount,
      sinMovInv,
      riskInv,
      overInv
    };
  }
  return map;
}

function computeGlobalSummary(summaryMap){
  const skuCount = MASTER.length; // unique SKUs
  let invAll = 0;
  let promAll = 0;
  let riskPairs = 0;
  let overPairs = 0;
  let sinMovInvPairs = 0;

  for (const wh of WAREHOUSES){
    const m = colMap(wh);
    invAll += MASTER.reduce((a,r)=>a + safeNum(r[m.inv]), 0);
    promAll += MASTER.reduce((a,r)=>a + safeNum(r[m.prom]), 0);
    riskPairs += MASTER.filter(r => safeNum(r[m.cobDias]) > 0 && safeNum(r[m.cobDias]) < TH_RISK).length;
    overPairs += MASTER.filter(r => safeNum(r[m.cobDias]) > TH_OVER).length;
    sinMovInvPairs += MASTER.filter(r => r[m.cls] === "Sin Mov" && safeNum(r[m.inv]) > 0).length;
  }

  const pairsTotal = MASTER.length * WAREHOUSES.length;

  const cobDiasGlobal = promAll > 0 ? (invAll / (promAll / 30)) : null;

  return {
    months: MONTHS_USED,
    skus: skuCount,
    invAll,
    promAll,
    cobDiasGlobal,
    riskPct: pairsTotal ? (riskPairs / pairsTotal) : 0,
    overPct: pairsTotal ? (overPairs / pairsTotal) : 0,
    sinMovInvPairs
  };
}

// ======= Rendering Summary =======
function renderSummary(){
  const byWh = computeSummaryByWarehouse();
  const global = computeGlobalSummary(byWh);

  // KPIs global
  gMeses.textContent = global.months ?? "—";
  gSkus.textContent = fmtNum(global.skus, 0);
  gInv.textContent = fmtNum(global.invAll, 0);
  gPromMes.textContent = fmtNum(global.promAll, 2);
  gCobDias.textContent = global.cobDiasGlobal === null ? "—" : fmtNum(global.cobDiasGlobal, 1);

  gRiskPct.textContent = (100*global.riskPct).toLocaleString("es-MX", {maximumFractionDigits:1, minimumFractionDigits:1}) + "%";
  gOverPct.textContent = (100*global.overPct).toLocaleString("es-MX", {maximumFractionDigits:1, minimumFractionDigits:1}) + "%";
  gSinMovInv.textContent = fmtNum(global.sinMovInvPairs, 0);

  // Charts
  renderStackedChart(byWh);
  renderRiskOverChart(byWh);

  // Heatmap table
  renderHeatmap(byWh);

  // Rankings
  renderRankings(byWh);

  // Insights
  renderInsights(byWh, global);
}

function destroyChart(ch){
  if (ch) { ch.destroy(); }
}

function renderStackedChart(byWh){
  const labels = WAREHOUSES.slice();
  const skuN = MASTER.length || 1;

  const A = labels.map(wh => byWh[wh].A / skuN * 100);
  const B = labels.map(wh => byWh[wh].B / skuN * 100);
  const C = labels.map(wh => byWh[wh].C / skuN * 100);
  const S = labels.map(wh => byWh[wh].S / skuN * 100);

  const ctx = document.getElementById("chartStacked").getContext("2d");
  destroyChart(chartStacked);

  chartStacked = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        { label: "A", data: A, stack: "stack1" },
        { label: "B", data: B, stack: "stack1" },
        { label: "C", data: C, stack: "stack1" },
        { label: "Sin Mov", data: S, stack: "stack1" },
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: "y",
      plugins: {
        legend: { position: "bottom" },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.raw.toFixed(1)}%`
          }
        }
      },
      scales: {
        x: {
          stacked: true,
          max: 100,
          ticks: { callback: (v)=> v + "%" }
        },
        y: { stacked: true }
      }
    }
  });
}

function renderRiskOverChart(byWh){
  const labels = WAREHOUSES.slice();
  const skuN = MASTER.length || 1;

  const risk = labels.map(wh => byWh[wh].riskCount / skuN * 100);
  const over = labels.map(wh => byWh[wh].overCount / skuN * 100);

  const ctx = document.getElementById("chartRiskOver").getContext("2d");
  destroyChart(chartRiskOver);

  chartRiskOver = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        { label: `Riesgo <${TH_RISK} días`, data: risk },
        { label: `Sobreinv >${TH_OVER} días`, data: over },
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: "bottom" } },
      scales: {
        y: { ticks: { callback:(v)=> v + "%" } }
      }
    }
  });
}

function heatAlpha(p){
  // p in [0..1] -> alpha 0.06..0.35
  return 0.06 + Math.min(0.29, p * 0.29);
}

function renderHeatmap(byWh){
  heatHead.innerHTML = "";
  heatBody.innerHTML = "";

  const trh = document.createElement("tr");
  ["Sucursal","A","B","C","Sin Mov","Inv total","Cobertura (días) mediana","% Riesgo","% Sobreinv","Sin Mov c/inv"].forEach(h=>{
    const th = document.createElement("th");
    th.textContent = h;
    trh.appendChild(th);
  });
  heatHead.appendChild(trh);

  const skuN = MASTER.length || 1;

  for (const wh of WAREHOUSES){
    const m = byWh[wh];

    const tr = document.createElement("tr");

    const td0 = document.createElement("td");
    td0.textContent = wh;
    td0.classList.add("mono");
    tr.appendChild(td0);

    const cells = [
      { label:"A", val:m.A, p:m.A/skuN },
      { label:"B", val:m.B, p:m.B/skuN },
      { label:"C", val:m.C, p:m.C/skuN },
      { label:"Sin Mov", val:m.S, p:m.S/skuN },
    ];

    for (const c of cells){
      const td = document.createElement("td");
      const box = document.createElement("div");
      box.classList.add("heatCell");
      box.style.background = `rgba(122,162,255,${heatAlpha(c.p)})`;
      box.innerHTML = `<b class="mono">${fmtNum(c.val,0)}</b><span class="pct">${pct(c.val, skuN,1)}</span>`;
      td.appendChild(box);
      tr.appendChild(td);
    }

    const tdInv = document.createElement("td");
    tdInv.textContent = fmtNum(m.invTotal,0);
    tdInv.classList.add("mono");
    tr.appendChild(tdInv);

    const tdMed = document.createElement("td");
    tdMed.textContent = m.covMed === null ? "—" : fmtNum(m.covMed,1);
    tdMed.classList.add("mono");
    tr.appendChild(tdMed);

    const tdRisk = document.createElement("td");
    tdRisk.textContent = pct(m.riskCount, skuN,1);
    tdRisk.classList.add("mono");
    tr.appendChild(tdRisk);

    const tdOver = document.createElement("td");
    tdOver.textContent = pct(m.overCount, skuN,1);
    tdOver.classList.add("mono");
    tr.appendChild(tdOver);

    const tdSM = document.createElement("td");
    tdSM.innerHTML = `<b class="mono">${fmtNum(m.sinMovInv,0)}</b> <span class="heatSmall">(${pct(m.sinMovInv, skuN,1)})</span>`;
    tr.appendChild(tdSM);

    heatBody.appendChild(tr);
  }
}

function renderRankings(byWh){
  rankRisk.innerHTML = "";
  rankOver.innerHTML = "";

  const skuN = MASTER.length || 1;
  const arr = WAREHOUSES.map(wh => ({
    wh,
    riskPct: byWh[wh].riskCount / skuN,
    overPct: byWh[wh].overCount / skuN,
    riskInv: byWh[wh].riskInv,
    overInv: byWh[wh].overInv
  }));

  const riskSorted = arr.slice().sort((a,b)=> b.riskPct - a.riskPct);
  const overSorted = arr.slice().sort((a,b)=> b.overPct - a.overPct);

  for (const r of riskSorted){
    const li = document.createElement("li");
    li.innerHTML = `<b>${r.wh}</b> — ${pct(r.riskPct,1,1)} • <span class="heatSmall">piezas en riesgo: ${fmtNum(r.riskInv,0)}</span>`;
    rankRisk.appendChild(li);
  }

  for (const r of overSorted){
    const li = document.createElement("li");
    li.innerHTML = `<b>${r.wh}</b> — ${pct(r.overPct,1,1)} • <span class="heatSmall">piezas sobreinv: ${fmtNum(r.overInv,0)}</span>`;
    rankOver.appendChild(li);
  }
}

function renderInsights(byWh, global){
  insights.innerHTML = "";
  const skuN = MASTER.length || 1;

  const arr = WAREHOUSES.map(wh => ({
    wh,
    sinMovPct: byWh[wh].S / skuN,
    riskPct: byWh[wh].riskCount / skuN,
    overPct: byWh[wh].overCount / skuN,
    inv: byWh[wh].invTotal,
    prom: byWh[wh].promTotal,
    sinMovInv: byWh[wh].sinMovInv
  }));

  const maxSinMov = arr.slice().sort((a,b)=> b.sinMovPct - a.sinMovPct)[0];
  const maxRisk = arr.slice().sort((a,b)=> b.riskPct - a.riskPct)[0];
  const maxOver = arr.slice().sort((a,b)=> b.overPct - a.overPct)[0];
  const maxInv = arr.slice().sort((a,b)=> b.inv - a.inv)[0];

  const bullets = [
    `La sucursal con mayor proporción de <b>Sin Mov</b> es <b>${maxSinMov.wh}</b> con ${pct(maxSinMov.sinMovPct,1,1)} del catálogo (SKUs).`,
    `Mayor <b>riesgo de quiebre</b> (&lt;${TH_RISK} días) en <b>${maxRisk.wh}</b>: ${pct(maxRisk.riskPct,1,1)} de SKUs en riesgo.`,
    `Mayor <b>sobreinventario</b> (&gt;${TH_OVER} días) en <b>${maxOver.wh}</b>: ${pct(maxOver.overPct,1,1)} de SKUs sobrestock.`,
    `<b>${maxInv.wh}</b> concentra el mayor inventario: <b>${fmtNum(maxInv.inv,0)}</b> piezas.`,
    `Cobertura global estimada: <b>${global.cobDiasGlobal===null?"—":fmtNum(global.cobDiasGlobal,1)}</b> días, con prom. mensual total <b>${fmtNum(global.promAll,2)}</b>.`,
    `Sin Mov con inventario (sumando sucursales): <b>${fmtNum(global.sinMovInvPairs,0)}</b> casos SKU-sucursal (ojo: inmoviliza capital).`
  ];

  for (const b of bullets){
    const li = document.createElement("li");
    li.innerHTML = b;
    insights.appendChild(li);
  }
}

// ======= Detail view computations =======
function getViewColumns(wh){
  const m = colMap(wh);
  return [
    { key: "Codigo", label: "Codigo", type: "text" },
    { key: "desc_prod", label: "desc_prod", type: "text" },
    { key: m.inv, label: `Inv-${wh}`, type: "num", dec: 0 },
    { key: m.cls, label: `Clasificacion-${wh}`, type: "text" },
    { key: m.prom, label: `Promedio Vta Mes-${wh}`, type: "num", dec: 2 },
    { key: m.cobMes, label: `Cobertura (Mes)-${wh}`, type: "num", dec: 2 },
    { key: m.cobDias, label: `Cobertura Dias (30) -${wh}`, type: "num", dec: 2 },
  ];
}

function applyFilters(rows){
  const wh = warehouseSelect.value;
  const m = colMap(wh);

  const cls = classSelect.value;
  const q = (searchInput.value || "").trim().toLowerCase();
  const onlyInv = !!onlyInvToggle.checked;
  const minD = covMin.value !== "" ? Number(covMin.value) : null;
  const maxD = covMax.value !== "" ? Number(covMax.value) : null;

  const smInv = !!onlySinMovInv.checked;
  const abZero = !!onlyABZero.checked;

  return rows.filter(r => {
    if (!wh) return false;

    if (cls !== "ALL" && r[m.cls] !== cls) return false;

    if (q){
      const ok = (String(r.Codigo).toLowerCase().includes(q) ||
                  String(r.desc_prod).toLowerCase().includes(q));
      if (!ok) return false;
    }

    if (onlyInv && safeNum(r[m.inv]) <= 0) return false;

    const d = safeNum(r[m.cobDias]);
    if (minD !== null && d < minD) return false;
    if (maxD !== null && d > maxD) return false;

    if (smInv){
      if (!(r[m.cls] === "Sin Mov" && safeNum(r[m.inv]) > 0)) return false;
    }

    if (abZero){
      const isAB = (r[m.cls] === "A" || r[m.cls] === "B");
      if (!(isAB && safeNum(r[m.inv]) === 0)) return false;
    }

    return true;
  });
}

function sortRows(rows, columns){
  if (!sortKey) return rows;
  const col = columns.find(c => c.key === sortKey);
  if (!col) return rows;

  const dir = sortDir === "asc" ? 1 : -1;
  const copy = [...rows];

  copy.sort((a,b) => {
    const va = a[sortKey];
    const vb = b[sortKey];
    if (col.type === "num") return (safeNum(va) - safeNum(vb)) * dir;
    return String(va ?? "").localeCompare(String(vb ?? ""), "es", { sensitivity:"base" }) * dir;
  });

  return copy;
}

function renderTable(rows){
  const wh = warehouseSelect.value;
  const columns = getViewColumns(wh);

  // header
  thead.innerHTML = "";
  const trh = document.createElement("tr");
  columns.forEach(col => {
    const th = document.createElement("th");
    const arrow = (sortKey === col.key) ? (sortDir === "asc" ? " ▲" : " ▼") : "";
    th.textContent = col.label + arrow;
    th.addEventListener("click", () => {
      if (sortKey === col.key) sortDir = (sortDir === "asc" ? "desc" : "asc");
      else { sortKey = col.key; sortDir = "asc"; }
      refreshDetail();
    });
    trh.appendChild(th);
  });
  thead.appendChild(trh);

  // body
  tbody.innerHTML = "";
  rows.forEach(r => {
    const tr = document.createElement("tr");
    columns.forEach(col => {
      const td = document.createElement("td");
      const v = r[col.key];

      if (col.type === "num"){
        const n = safeNum(v);
        td.textContent = fmtNum(n, col.dec ?? 0);
        td.classList.add("mono");
        if (n < 0) td.classList.add("bad");
      } else {
        td.textContent = String(v ?? "");
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  // hint
  tableHint.textContent = `Sucursal: ${wh} • Registros: ${rows.length.toLocaleString("es-MX")}`;

  // KPIs + charts + tops
  updateDetailKPIs(rows);
  renderHist(rows);
  renderTopLists(rows);
}

function updateDetailKPIs(viewRows){
  const wh = warehouseSelect.value;
  const m = colMap(wh);

  const invTot = viewRows.reduce((a,r)=>a + safeNum(r[m.inv]), 0);
  const promTot = viewRows.reduce((a,r)=>a + safeNum(r[m.prom]), 0);

  const covVals = viewRows.map(r=>safeNum(r[m.cobDias])).filter(n=>Number.isFinite(n) && n>=0);
  const covMed = median(covVals);

  const countA = viewRows.filter(r => r[m.cls] === "A").length;
  const countB = viewRows.filter(r => r[m.cls] === "B").length;
  const countC = viewRows.filter(r => r[m.cls] === "C").length;
  const countS = viewRows.filter(r => r[m.cls] === "Sin Mov").length;

  const riskCount = viewRows.filter(r => {
    const d = safeNum(r[m.cobDias]);
    return d > 0 && d < TH_RISK;
  }).length;

  const overCount = viewRows.filter(r => safeNum(r[m.cobDias]) > TH_OVER).length;

  const smInv = viewRows.filter(r => r[m.cls] === "Sin Mov" && safeNum(r[m.inv]) > 0).length;

  kpiMeses.textContent = MONTHS_USED ?? "—";
  kpiSkus.textContent = fmtNum(viewRows.length,0);
  kpiInv.textContent = fmtNum(invTot,0);
  kpiProm.textContent = fmtNum(promTot,2);
  kpiCobMed.textContent = covMed === null ? "—" : fmtNum(covMed,1);

  kpiRisk.textContent = pct(riskCount, viewRows.length,1);
  kpiOver.textContent = pct(overCount, viewRows.length,1);
  kpiSinMovInv.textContent = fmtNum(smInv,0);

  kpiA.textContent = fmtNum(countA,0);
  kpiB.textContent = fmtNum(countB,0);
  kpiC.textContent = fmtNum(countC,0);
  kpiS.textContent = fmtNum(countS,0);
}

function renderHist(viewRows){
  const wh = warehouseSelect.value;
  const m = colMap(wh);

  const buckets = [
    { name: `0-${TH_RISK}`, from: 0, to: TH_RISK },
    { name: "16-30", from: 15.00001, to: 30 },
    { name: "31-60", from: 30.00001, to: 60 },
    { name: "61-120", from: 60.00001, to: 120 },
    { name: ">120", from: 120.00001, to: Infinity },
  ];
  const counts = buckets.map(()=>0);

  for (const r of viewRows){
    const d = safeNum(r[m.cobDias]);
    for (let i=0; i<buckets.length; i++){
      const b = buckets[i];
      if (d >= b.from && d <= b.to){
        counts[i] += 1;
        break;
      }
    }
  }

  const ctx = document.getElementById("chartHist").getContext("2d");
  destroyChart(chartHist);

  chartHist = new Chart(ctx, {
    type: "bar",
    data: {
      labels: buckets.map(b=>b.name),
      datasets: [{ label: "SKUs", data: counts }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } }
    }
  });
}

function renderTopLists(viewRows){
  const wh = warehouseSelect.value;
  const m = colMap(wh);

  // Sobreinventario: altos días con inventario >0
  const over = viewRows
    .filter(r => safeNum(r[m.inv]) > 0)
    .slice()
    .sort((a,b)=> safeNum(b[m.cobDias]) - safeNum(a[m.cobDias]))
    .slice(0, 15);

  // Riesgo: días más bajos (incluye inv=0 para detectar quiebre)
  const risk = viewRows
    .slice()
    .sort((a,b)=> safeNum(a[m.cobDias]) - safeNum(b[m.cobDias]))
    .slice(0, 15);

  topOver.innerHTML = "";
  topRisk.innerHTML = "";

  for (const r of over){
    const li = document.createElement("li");
    li.innerHTML = `<b class="mono">${r.Codigo}</b> — ${r.desc_prod} <span class="heatSmall">| inv ${fmtNum(r[m.inv],0)} | ${fmtNum(r[m.cobDias],1)} días</span>`;
    topOver.appendChild(li);
  }

  for (const r of risk){
    const li = document.createElement("li");
    li.innerHTML = `<b class="mono">${r.Codigo}</b> — ${r.desc_prod} <span class="heatSmall">| inv ${fmtNum(r[m.inv],0)} | ${fmtNum(r[m.cobDias],1)} días</span>`;
    topRisk.appendChild(li);
  }
}

function exportViewToCSV(viewRows){
  const wh = warehouseSelect.value;
  const cols = getViewColumns(wh);

  const header = cols.map(c => c.label);
  const lines = [header.join(",")];

  for (const r of viewRows){
    const row = cols.map(c => {
      let val = r[c.key];
      if (val === null || val === undefined) val = "";
      const s = String(val).replace(/"/g,'""');
      return (s.includes(",") || s.includes("\n")) ? `"${s}"` : s;
    });
    lines.push(row.join(","));
  }

  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `vista_${wh}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ======= Refresh Detail Pipeline =======
function refreshDetail(){
  if (!MASTER.length) return;
  const filtered = applyFilters(MASTER);
  const cols = getViewColumns(warehouseSelect.value);
  const sorted = sortRows(filtered, cols);
  renderTable(sorted);
}

// ======= Tabs =======
function showSummary(){
  summaryView.classList.remove("hidden");
  detailView.classList.add("hidden");
  tabSummary.classList.add("active");
  tabDetail.classList.remove("active");
}

function showDetail(){
  summaryView.classList.add("hidden");
  detailView.classList.remove("hidden");
  tabSummary.classList.remove("active");
  tabDetail.classList.add("active");
}

// ======= Events =======
tabSummary.addEventListener("click", () => showSummary());
tabDetail.addEventListener("click", () => showDetail());

fileInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  setStatus("Leyendo archivo…", `${file.name} • ${Math.round(file.size/1024).toLocaleString("es-MX")} KB`);

  try{
    let rows = [];
    if (file.name.toLowerCase().endsWith(".csv")){
      const text = await file.text();
      rows = parseCSV(text);
    } else {
      const buf = await file.arrayBuffer();
      rows = parseXLSX(buf);
    }

    rows = normalizeColumns(rows);
    const warehouses = validateMaster(rows);

    WAREHOUSES = warehouses;
    MASTER = coerceTypes(rows, warehouses);

    fillWarehouseSelect(WAREHOUSES);

    // enable
    setEnabled(true);
    setStatus("✅ Archivo cargado correctamente",
      `Sucursales detectadas: ${WAREHOUSES.join(", ")} • Filas (SKUs): ${MASTER.length.toLocaleString("es-MX")}`);

    // reset filters
    classSelect.value = "ALL";
    searchInput.value = "";
    onlyInvToggle.checked = false;
    covMin.value = "";
    covMax.value = "";
    onlySinMovInv.checked = false;
    onlyABZero.checked = false;

    sortKey = null;
    sortDir = "asc";

    // render both views
    renderSummary();
    refreshDetail();

    // default view
    showSummary();

  } catch(err){
    console.error(err);
    setEnabled(false);
    setStatus("❌ No se pudo cargar el archivo", String(err?.message || err));
  }
});

btnClear.addEventListener("click", () => {
  MASTER = [];
  WAREHOUSES = [];
  MONTHS_USED = null;

  fileInput.value = "";
  setEnabled(false);
  setStatus("Carga un archivo para comenzar.", "");

  // clear tables
  heatHead.innerHTML = "";
  heatBody.innerHTML = "";
  thead.innerHTML = "";
  tbody.innerHTML = "";
  rankRisk.innerHTML = "";
  rankOver.innerHTML = "";
  insights.innerHTML = "";
  topOver.innerHTML = "";
  topRisk.innerHTML = "";
  tableHint.textContent = "—";

  // clear kpis
  [gMeses,gSkus,gInv,gCobDias,gPromMes,gRiskPct,gOverPct,gSinMovInv].forEach(x=>x.textContent="—");
  [kpiMeses,kpiSkus,kpiInv,kpiCobMed,kpiProm,kpiRisk,kpiOver,kpiSinMovInv,kpiA,kpiB,kpiC,kpiS].forEach(x=>x.textContent="—");

  destroyChart(chartStacked); chartStacked=null;
  destroyChart(chartRiskOver); chartRiskOver=null;
  destroyChart(chartHist); chartHist=null;

  showSummary();
});

warehouseSelect.addEventListener("change", () => { sortKey=null; sortDir="asc"; refreshDetail(); });
classSelect.addEventListener("change", refreshDetail);
searchInput.addEventListener("input", debounce(refreshDetail, 140));
onlyInvToggle.addEventListener("change", refreshDetail);
covMin.addEventListener("input", debounce(refreshDetail, 180));
covMax.addEventListener("input", debounce(refreshDetail, 180));

onlySinMovInv.addEventListener("change", () => {
  // si activas SinMovInv, apaga ABZero para evitar conflicto
  if (onlySinMovInv.checked) onlyABZero.checked = false;
  refreshDetail();
});
onlyABZero.addEventListener("change", () => {
  if (onlyABZero.checked) onlySinMovInv.checked = false;
  refreshDetail();
});

btnExport.addEventListener("click", () => {
  const view = sortRows(applyFilters(MASTER), getViewColumns(warehouseSelect.value));
  exportViewToCSV(view);
});

// Initial state
setEnabled(false);
showSummary();
