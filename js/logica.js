// logica.js - utilidades y lógica de fechas/parsers/export
window.APP = window.APP || {
  workbook: null,
  sourceFileName: "",
  parsedRows: [],
  headerRowIndex: null,
  sheetUsedName: "",
  seasons: []
};

const REQUIRED_HEADERS = [
  "STATUS\nSeason",
  "Season",
  "HOD\nLULULEMON",
  "HOD\nCOFACO\nPCP",
  "PO #",
  "OP",
  "Style",
  "Style Name",
  "Color Code",
  "Color Description",
  "Color Name",
];

const OUT_COLUMNS = [
  "STATUS\nSeason",
  "Season",
  "PO #",
  "OP",
  "HOD\nCOFACO\nPCP",
  "HOD\nLULULEMON",
  "Style",
  "Style Name",
  "Color Code",
  "Color Description",
  "Color Name",
  "TOP DATE",
  "REAL DATE"
];

function normHeader(s){
  if (s === null || s === undefined) return "";
  return String(s)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .trim();
}

function stripAccents(s){
  return String(s || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function normKey(s){
  return stripAccents(String(s || "").trim().toLowerCase());
}

function isValidDate(d){
  return d instanceof Date && !isNaN(d.getTime());
}

function parseExcelDate(v){
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) return isValidDate(v) ? v : null;
  if (typeof v === "number" && window.XLSX && XLSX.SSF && XLSX.SSF.parse_date_code){
    const dc = XLSX.SSF.parse_date_code(v);
    if (dc){
      const d = new Date(Date.UTC(dc.y, dc.m - 1, dc.d, dc.H || 0, dc.M || 0, Math.floor(dc.S || 0)));
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
  }
  const s = String(v).trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m){
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return isValidDate(d) ? d : null;
  }
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m){
    const mm = Number(m[1]), dd = Number(m[2]), yy = Number(m[3]);
    const year = (yy < 100) ? (2000 + yy) : yy;
    const d = new Date(year, mm - 1, dd);
    return isValidDate(d) ? d : null;
  }
  m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
  if (m){
    const dd = Number(m[1]);
    const mon = m[2].toLowerCase();
    const yy = Number(m[3]);
    const year = (yy < 100) ? (2000 + yy) : yy;
    const map = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    if (map.hasOwnProperty(mon)){
      const d = new Date(year, map[mon], dd);
      return isValidDate(d) ? d : null;
    }
  }
  const d = new Date(s);
  return isValidDate(d) ? d : null;
}

function fmtYYYYMMDD(d){
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,"0");
  const day = String(d.getDate()).padStart(2,"0");
  return `${y}-${m}-${day}`;
}

function subtractBusinessDays(dateObj, n){
  const d = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  let remaining = n;
  while (remaining > 0){
    d.setDate(d.getDate() - 1);
    const day = d.getDay();
    // Días hábiles: lun, mar, jue, vie (no sáb, dom ni miércoles)
    if (day !== 0 && day !== 6 && day !== 3) remaining--;
  }
  return d;
}

function adjustNoWednesday(dateObj){
  const d = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  if (d.getDay() === 3){ d.setDate(d.getDate() - 1); }
  return d;
}

function detectHeaderRow(aoa){
  const maxScan = Math.min(50, aoa.length);
  let bestIdx = null;
  let bestScore = -1;
  const requiredNorm = REQUIRED_HEADERS.map(normHeader);
  for (let r=0; r<maxScan; r++){
    const row = aoa[r] || [];
    const set = new Set(row.map(normHeader));
    let score = 0;
    for (const req of requiredNorm){ if (set.has(req)) score++; }
    if (score > bestScore){ bestScore = score; bestIdx = r; }
  }
  if (bestScore < 4){ return (aoa.length > 5) ? 5 : 0; }
  return bestIdx;
}

function aoaToObjects(aoa, hdrRow){
  const headers = (aoa[hdrRow] || []).map(normHeader);
  const headerIndex = new Map();
  headers.forEach((h, i) => { if (h) headerIndex.set(h, i); });
  const missing = [];
  for (const req of REQUIRED_HEADERS){ if (!headerIndex.has(normHeader(req))) missing.push(req); }
  if (missing.length){ throw new Error("Faltan columnas requeridas: " + missing.join(", ")); }
  const rows = [];
  for (let r = hdrRow + 1; r < aoa.length; r++){
    const row = aoa[r];
    if (!row) continue;
    const hasAny = row.some(v => v !== null && v !== undefined && String(v).trim() !== "");
    if (!hasAny) continue;
    const obj = {};
    headers.forEach((h, i) => { if (!h) return; obj[h] = row[i]; });
    rows.push(obj);
  }
  return rows;
}

function toSafeIntOrBlank(v){
  if (v === null || v === undefined || v === "") return "";
  const s = String(v).trim();
  if (/^0\d+$/.test(s)) return s;
  const n = Number(s);
  if (!isFinite(n)) return s;
  const i = Math.trunc(n);
  return (Math.abs(n - i) < 1e-9) ? i : n;
}

function computeColWidths(aoa){
  const cols = aoa[0]?.length || 0;
  const w = new Array(cols).fill(10);
  for (let c=0; c<cols; c++){
    let mx = 10;
    for (let r=0; r<aoa.length; r++){
      let v = aoa[r][c];
      let s = "";
      if (v instanceof Date){ s = fmtYYYYMMDD(v); }
      else if (v === null || v === undefined){ s = ""; }
      else { s = String(v); }
      mx = Math.max(mx, Math.min(45, s.length + 2));
    }
    w[c] = mx;
  }
  return w.map(x => ({ wch: x }));
}
