// api/src/generate-report.js
// Azure version — fully unified report builder.
// Sample meta from SP "Archived Intake" list (was Google Sheets).
// Instrument data from SharePoint files (same as Netlify generate-report-sp.js).
// No cross-cloud split needed — Graph API calls take <1s from Azure Functions.
//
// Required env vars:
//   MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, SP_SITE_ID
//   SP_CONTROL_FOLDER  (/sites/Laboratory/Shared Documents/Documents/Test C)
//   SP_ICPMS_FOLDER    (/sites/Laboratory/Shared Documents/Documents/Test M)
//   SP_ACID_FOLDER     (/sites/Laboratory/Shared Documents/Documents/metals prep)
//   ANTHROPIC_API_KEY  (for Results Cache fallback if needed)

const { app }   = require('@azure/functions');
const XLSX      = require('xlsx');
const { getToken } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

// ─────────────────────────────────────────────────────────────────────────────
// PARAMETER CONFIG — identical to Netlify version
// ─────────────────────────────────────────────────────────────────────────────
const PARAM_CONFIG = [
  { name:'Chloride, Total',              rl:2.00,    epa:250,       unit:'mg/L', method:'SM4500Cl',         source:'gallery', colKey:'Chloride mg/L',                packages:['Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive','Pro Plus'] },
  { name:'Fluoride, Total',              rl:0.20,    epa:4,         unit:'mg/L', method:'SM4500F',          source:'gallery', colKey:'Fluoride mg/L',                packages:['Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Nitrite-Nitrogen, Total',      rl:0.20,    epa:1,         unit:'mg/L', method:'EPA 354.1',        source:'gallery', colKey:'Nitrite mg/L',                 packages:['Basic Safety (FHA)','Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Nitrate-Nitrogen, Total',      rl:1.00,    epa:10,        unit:'mg/L', method:'SM4500NO3',        source:'gallery', colKey:'Nitrate mg/L',                 packages:['Basic Safety (FHA)','Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Arsenic, Total',               rl:1.00,    epa:10,        unit:'ug/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'As 75 (ug/L)',               packages:['Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Lead, Total',                  rl:1.00,    epa:15,        unit:'ug/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'Pb 208 (ug/L)',              packages:['Basic Safety (FHA)','Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Uranium, Total',               rl:1.00,    epa:30,        unit:'ug/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'U 238 (ug/L)',               packages:['Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Copper, Total',                rl:0.001,   epa:1.3,       unit:'mg/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'Cu 63 (mg/L)',               packages:['Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Iron, Total',                  rl:0.05,    epa:0.3,       unit:'mg/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'Fe 54 (mg/L)',               packages:['Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Manganese, Total',             rl:0.001,   epa:0.05,      unit:'mg/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'Mn 55 (mg/L)',               packages:['Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Sodium, Total',                rl:1.00,    epa:null,      unit:'mg/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'Na 23 (mg/L)',               packages:['Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Hardness by calculation',      rl:0.91,    epa:null,      unit:'mg/L', method:'',                 source:'calc',                                           packages:['Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Calcium, Total',               rl:0.2,     epa:null,      unit:'mg/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'Ca 43 (mg/L)',               packages:['Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Magnesium, Total',             rl:0.1,     epa:null,      unit:'mg/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'Mg 24 (mg/L)',               packages:['Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'Antimony, Total',              rl:0.0005,  epa:0.006,     unit:'mg/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'Sb 121 (mg/L)',              packages:['Comprehensive'] },
  { name:'Cadmium, Total',               rl:0.002,   epa:0.005,     unit:'mg/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'Cd 111 (mg/L)',              packages:['Comprehensive'] },
  { name:'Chromium, Total',              rl:0.002,   epa:0.1,       unit:'mg/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'Cr 52 (mg/L)',               packages:['Comprehensive'] },
  { name:'Cobalt',                       rl:null,    epa:null,      unit:'mg/L', method:'EPA 200.8',        source:'icpms',   icpmsCol:'Co 59 (mg/L)',               packages:['Comprehensive'] },
  { name:'pH Electrometric',             rl:null,    epa:'6.5-8.5', unit:'',     method:'SM4500H+B',        source:'ph',      decimals:2,                            packages:['Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive','Pro Plus'] },
  { name:'Alkalinity',                   rl:40.00,   epa:null,      unit:'mg/L', method:'',                 source:'gallery', colKey:'Alkalinity mg/L',              packages:['Comprehensive','Pro Plus'] },
  { name:'Sulfate',                      rl:40.00,   epa:250,       unit:'mg/L', method:'SM4500-SO4',       source:'gallery', colKey:'Sulfate mg/L',                 packages:['Comprehensive'] },
  { name:'Tannins',                      rl:null,    epa:null,      unit:'',     method:'Hach Method 8193', source:'gallery', colKey:'Tannins mg/L',                 packages:['Pro Plus'] },
  { name:'Total Dissolved Solids (TDS)', rl:null,    epa:null,      unit:'ppm',  method:'SM4500C1E',        source:'gallery', colKey:'Total Dissolved Solids (TDS)', packages:['Pro Plus'] },
  { name:'Bromide',                      rl:null,    epa:null,      unit:'mg/L', method:'HI 93716',         source:'gallery', colKey:'Bromide',                      packages:['Pro Plus'] },
  { name:'Total Coliform',               rl:null,    epa:1,         unit:'MPN',  method:'SM9223 B',         source:'control', colKey:'Coliform MPN',                 packages:['Basic Safety (FHA)','Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
  { name:'E. Coli',                      rl:null,    epa:1,         unit:'MPN',  method:'SM9223 B',         source:'control', colKey:'Ecoli MPN',                    packages:['Basic Safety (FHA)','Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'] },
];

const FHA_PARAM_NAMES = ['Nitrite-Nitrogen, Total','Nitrate-Nitrogen, Total','Lead, Total','Total Coliform','E. Coli'];
const NEEDS_FHA_TYPES = ['Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'];
const MONTH_TABS      = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sept','Oct','Nov','Dec'];
const MONTH_LONG      = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const ICPMS_COLS      = ['Na 23 (mg/L)','Mg 24 (mg/L)','Ca 43 (mg/L)','Cr 52 (mg/L)',
                         'Fe 54 (mg/L)','Mn 55 (mg/L)','Co 59 (mg/L)','Cu 63 (mg/L)',
                         'As 75 (ug/L)','Cd 111 (mg/L)','Sb 121 (mg/L)','Pb 208 (ug/L)','U 238 (ug/L)'];

// ── Token cache (module-level, survives warm starts) ──────────────────────────
let _cachedToken = null;
let _tokenExpiry = 0;
async function getCachedToken() {
  if (_cachedToken && Date.now() < _tokenExpiry) return _cachedToken;
  _cachedToken = await getToken();
  _tokenExpiry = Date.now() + 55 * 60 * 1000;
  return _cachedToken;
}

// ── Graph helpers ─────────────────────────────────────────────────────────────
async function graphGet(path, token) {
  const res = await fetch(`${GRAPH}${path}`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph GET ${path} → ${res.status}: ${await res.text()}`);
  return res.json();
}

async function graphGetBytes(path, token) {
  const res = await fetch(`${GRAPH}${path}`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph download → ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

// ── Date helpers ──────────────────────────────────────────────────────────────
function extractDateStr(labId) {
  return String(labId).match(/^(\d{6})/)?.[1] || null;
}

function dateInfo(mmddyy) {
  const mm  = mmddyy.slice(0, 2), dd = mmddyy.slice(2, 4), yy = mmddyy.slice(4, 6);
  const idx = parseInt(mm) - 1;
  return {
    mm, dd, yy,
    year:        2000 + parseInt(yy),
    monthLong:   MONTH_LONG[idx],
    monthShort:  MONTH_TABS[idx],
    monthFolder: `${MONTH_LONG[idx]} ${2000 + parseInt(yy)}`,
  };
}

function fmtDT(raw) {
  if (!raw) return '';
  const s = String(raw).trim();
  if (!s) return '';
  if (/^\d{2}\/\d{2}\/\d{2}\s+\d{4}$/.test(s)) return s.replace(/(\d{2}\/\d{2}\/\d{2}\s+)(\d{2})(\d{2})$/, '$1$2:$3');
  if (/^\d{2}\/\d{2}\/\d{2}\s+\d{2}:\d{2}$/.test(s)) return s;
  const hasTime = /\d{1,2}:\d{2}/.test(s);
  if (hasTime) {
    const dt = new Date(s);
    if (!isNaN(dt.getTime())) {
      const m  = String(dt.getMonth() + 1).padStart(2, '0');
      const d  = String(dt.getDate()).padStart(2, '0');
      const y  = String(dt.getFullYear()).slice(-2);
      const h  = String(dt.getHours()).padStart(2, '0');
      const mi = String(dt.getMinutes()).padStart(2, '0');
      return `${m}/${d}/${y} ${h}:${mi}`;
    }
  }
  return s;
}

function combineDT(dateStr, timeStr) {
  if (!dateStr) return '';
  const ds = String(dateStr).trim();
  const ts = String(timeStr || '').trim();
  const dm = ds.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!dm) return ds;
  const m = dm[1].padStart(2, '0');
  const d = dm[2].padStart(2, '0');
  const y = dm[3].length === 4 ? dm[3].slice(-2) : dm[3].padStart(2, '0');
  let hhmm = '';
  if (ts) {
    const ampm = ts.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
    const h24  = ts.match(/^(\d{1,2}):(\d{2})/);
    if (ampm) {
      let h = parseInt(ampm[1]);
      const mi = ampm[2];
      if (ampm[3].toUpperCase() === 'PM' && h !== 12) h += 12;
      if (ampm[3].toUpperCase() === 'AM' && h === 12) h = 0;
      hhmm = `${String(h).padStart(2, '0')}:${mi}`;
    } else if (h24) {
      hhmm = `${String(parseInt(h24[1])).padStart(2, '0')}:${h24[2]}`;
    } else {
      hhmm = ts.replace(/[^0-9]/g, '').slice(0, 4);
    }
  }
  return hhmm ? `${m}/${d}/${y} ${hhmm}` : `${m}/${d}/${y}`;
}

// ── SharePoint file helpers ───────────────────────────────────────────────────
function toDrivePath(fullPath) {
  const marker = 'Shared Documents/';
  const idx    = fullPath.indexOf(marker);
  const rel    = idx >= 0 ? fullPath.slice(idx + marker.length) : fullPath.replace(/^\/+/, '');
  return rel.split('/').map(s => encodeURIComponent(s)).join('/');
}

async function listFolder(folderPath, token) {
  const siteId    = process.env.SP_SITE_ID;
  const drivePath = toDrivePath(folderPath);
  try {
    const data = await graphGet(`/sites/${siteId}/drive/root:/${drivePath}:/children?$select=id,name`, token);
    return data.value || [];
  } catch (e) {
    console.warn(`listFolder failed for ${folderPath}:`, e.message);
    return [];
  }
}

async function downloadFile(fileId, token) {
  const siteId = process.env.SP_SITE_ID;
  return graphGetBytes(`/sites/${siteId}/drive/items/${fileId}/content`, token);
}

// Try downloading a file by known path (fast — no folder listing needed)
async function tryDownloadByPath(folderPath, filename, token) {
  const siteId   = process.env.SP_SITE_ID;
  const fullPath = toDrivePath(`${folderPath}/${filename}`);
  try {
    return await graphGetBytes(`/sites/${siteId}/drive/root:/${fullPath}:/content`, token);
  } catch { return null; }
}

// ── Excel helpers ─────────────────────────────────────────────────────────────
function readSheetRows(buffer, sheetName, headerRowNum) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true, raw: false });
  let ws;
  if (sheetName) {
    const sn = wb.SheetNames.find(n => n === sheetName) || wb.SheetNames.find(n => n.toLowerCase() === (sheetName || '').toLowerCase());
    if (!sn) return { rows: [], found: false, sheetNames: wb.SheetNames };
    ws = wb.Sheets[sn];
  } else {
    ws = wb.Sheets[wb.SheetNames[0]];
  }
  const hIdx    = Math.max((headerRowNum || 1) - 1, 0);
  const aoa     = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  if (aoa.length <= hIdx) return { rows: [], found: true, sheetNames: wb.SheetNames };
  const headers = aoa[hIdx].map(h => String(h || '').replace(/[\n\r]+/g, ' ').replace(/\s+/g, ' ').trim());
  const rows    = aoa.slice(hIdx + 1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { if (h) obj[h] = String(row[i] ?? '').trim(); });
    return obj;
  }).filter(r => Object.values(r).some(v => v));
  return { rows, found: true, headers, sheetNames: wb.SheetNames };
}

function readAoA(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true, raw: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
}

// ── pH Parser ─────────────────────────────────────────────────────────────────
function parsePH(buffer, baseId) {
  const { rows } = readSheetRows(buffer, null, 9);
  const byCode   = {};
  for (const row of rows) {
    const code   = (row['Code']      || '').trim();
    const sample = (row['Sample']    || '').trim();
    const ph     = (row['pH']        || '').trim();
    const dt     = (row['Date/Time'] || '').trim();
    const qc     = (row['QC']        || '').trim().toLowerCase();
    if (!code) continue;
    const sampleBase = sample.match(/(\d{6}-\d{3})/)?.[1] || '';
    byCode[code] = { sampleBase, ph, dt, pass: qc === 'pass' };
  }
  for (const [code, dRow] of Object.entries(byCode)) {
    if (!code.endsWith(' D')) continue;
    if (dRow.sampleBase !== baseId) continue;
    if (!dRow.pass) continue;
    const nonDRow = byCode[code.replace(/ D$/, '')];
    if (nonDRow?.ph) return { value: nonDRow.ph, analDT: fmtDT(nonDRow.dt) };
  }
  return null;
}

// ── Control Sheet Parser ──────────────────────────────────────────────────────
function parseControlSheet(buffer, baseId) {
  const aoa = readAoA(buffer);
  if (!aoa.length) return {};

  let hIdx = -1, headers = [], headersLower = [];
  for (let i = 0; i < Math.min(aoa.length, 5); i++) {
    const row = aoa[i].map(c => String(c || '').replace(/[\n\r]+/g, ' ').replace(/\s+/g, ' ').trim());
    if (row.some(c => ['barcode', 'lab id', 'sample id'].includes(c.toLowerCase()))) {
      hIdx = i; headers = row;
      headersLower = headers.map(h => h.replace(/[\n\r]+/g, ' ').replace(/\s+/g, ' ').toLowerCase().trim());
      break;
    }
  }
  if (hIdx < 0) { console.warn('[parseControlSheet] Header row not found'); return {}; }

  const barcodeColIdx = headersLower.findIndex(h => ['barcode', 'lab id', 'sample id'].includes(h));

  const findCol = (key) => {
    const keyLower = key.replace(/[\n\r]+/g, ' ').replace(/\s+/g, ' ').toLowerCase().trim();
    let idx = headersLower.indexOf(keyLower);
    if (idx >= 0) return idx;
    idx = headersLower.findIndex(h => h && (h.startsWith(keyLower.split(' ')[0]) && keyLower.split(' ')[0].length > 3));
    return idx;
  };

  for (let r = hIdx + 1; r < aoa.length; r++) {
    const row        = aoa[r];
    const barcode    = String(row[barcodeColIdx] || '').trim();
    if (!barcode.match(/\d{6}-\d{3}/)) continue;
    const barcodeBase = barcode.match(/(\d{6}-\d{3})/)?.[1];
    if (barcodeBase !== baseId) continue;

    const getH = (header) => {
      const idx = findCol(header);
      return idx >= 0 ? String(row[idx] || '').trim() : '';
    };

    const gallery = {};
    PARAM_CONFIG.filter(p => p.source === 'gallery').forEach(p => {
      const colIdx = findCol(p.colKey);
      if (colIdx < 0) { console.warn(`[parseControlSheet] Column not found: "${p.colKey}"`); return; }
      gallery[p.name] = { value: String(row[colIdx] || '').trim(), analDT: fmtDT(String(row[colIdx + 1] || '').trim()) };
    });

    const pHColIdx   = headersLower.findIndex(h => h === 'ph');
    const pHDTColIdx = pHColIdx >= 0 ? pHColIdx + 1 : -1;

    return {
      gallery,
      coliformMPN:    getH('Coliform MPN'),
      ecoliMPN:       getH('Ecoli MPN'),
      coliformPrepDT: fmtDT(getH('Start Date/Time')),
      coliformAnalDT: fmtDT(getH('End Date/Time')),
      phValue:        pHColIdx >= 0 ? String(row[pHColIdx] || '').trim() : '',
      phDT:           pHDTColIdx >= 0 ? fmtDT(String(row[pHDTColIdx] || '').trim()) : '',
    };
  }
  console.warn(`[parseControlSheet] Lab ID ${baseId} not found`);
  return {};
}

// ── ICP-MS Parser ─────────────────────────────────────────────────────────────
function parseICPMS(buffer, baseId) {
  const { rows, found } = readSheetRows(buffer, 'Concentrations', 1);
  if (!found || !rows.length) return null;

  const entries = [];
  for (const row of rows) {
    const rawId = (row['Sample Id'] || row['Sample ID'] || '').trim();
    if (!rawId) continue;
    const idBase = rawId.match(/(\d{6}-\d{3})/)?.[1];
    if (!idBase || idBase !== baseId) continue;
    const qc = (row['QC Status'] || '').trim().toLowerCase();
    if (qc && qc !== 'passed') continue;
    const dilMatch  = rawId.match(/x(\d+)/i);
    const dilFactor = dilMatch ? parseInt(dilMatch[1]) : 0;
    const acqTime   = fmtDT(row['Acquisition Time'] || '');
    const elements  = {};
    for (const col of ICPMS_COLS) {
      const val = row[col];
      if (val !== undefined && val !== '') {
        const n = parseFloat(val);
        if (!isNaN(n)) elements[col] = n;
      }
    }
    entries.push({ dilFactor, elements, acqTime });
  }
  if (!entries.length) return null;

  entries.sort((a, b) => a.dilFactor - b.dilFactor);
  const merged = {};
  let acqTime  = '';
  for (const entry of entries) {
    if (!acqTime && entry.acqTime) acqTime = entry.acqTime;
    for (const [col, val] of Object.entries(entry.elements)) {
      if (!(col in merged)) merged[col] = val;
    }
  }
  return { elements: merged, acqTime };
}

// ── Acid Sheet Parser ─────────────────────────────────────────────────────────
function parseAcidSheet(buffer, baseId, monthShort) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true, raw: false });
  const candidates = [
    `acidification ${monthShort}`, `acidifcation ${monthShort}`,
    `Acidification ${monthShort}`, `Acidifcation ${monthShort}`,
  ];
  let ws = null;
  for (const name of candidates) {
    const found = wb.SheetNames.find(n => n.toLowerCase() === name.toLowerCase());
    if (found) { ws = wb.Sheets[found]; break; }
  }
  if (!ws) {
    const found = wb.SheetNames.find(n => n.toLowerCase().includes(monthShort.toLowerCase()));
    if (found) ws = wb.Sheets[found];
  }
  if (!ws) { console.warn(`Acid sheet: no tab for month ${monthShort}`); return null; }

  const aoa     = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  let hIdx      = -1;
  for (let i = 0; i < Math.min(aoa.length, 5); i++) {
    const row = aoa[i].map(c => String(c || '').toLowerCase().trim());
    if (row.some(c => c === 'date') && row.some(c => c.includes('sample'))) { hIdx = i; break; }
  }
  const COL_DATE = 0, COL_TIME = 1, COL_SAMPLE = 5;
  const startRow = hIdx >= 0 ? hIdx + 1 : 2;
  for (let r = startRow; r < aoa.length; r++) {
    const row       = aoa[r];
    const sampleRaw = String(row[COL_SAMPLE] || '').trim();
    if (!sampleRaw) continue;
    const cellBase = sampleRaw.match(/(\d{6}-\d{3})/)?.[1];
    if (!cellBase || cellBase !== baseId) continue;
    return combineDT(String(row[COL_DATE] || '').trim(), String(row[COL_TIME] || '').trim());
  }
  return null;
}

// ── SharePoint List: Archived Intake ─────────────────────────────────────────
// Replaces Google Sheets getSampleMeta()
// Columns: LabID (Title), Services, ClientName, DateDrawn, TimeDrawn,
//          ReceivedDate, ReceivedTime, Address, City, State, Zip
async function getSampleMeta(baseId, token) {
  try {
    const siteId = process.env.SP_SITE_ID;
    // Filter by LabID field (Title column stores the full lab ID like 070126-001)
    const filter = encodeURIComponent(`startswith(fields/Title,'${baseId}')`);
    const res    = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Archived Intake/items?$filter=${filter}&$expand=fields&$top=5`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) { console.warn('getSampleMeta:', res.status); return null; }
    const data = await res.json();
    if (!data.value?.length) return null;
    const f = data.value[0].fields;
    return {
      customer:     f.ClientName    || '',
      dateDrawn:    f.DateDrawn     || '',
      timeDrawn:    f.TimeDrawn     || '',
      dateReceived: f.ReceivedDate  || '',
      timeReceived: f.ReceivedTime  || '',
      location:     f.Address       || '',
      city:         f.City          || '',
      state:        f.State         || 'ME',
      zip:          f.Zip           || '',
      services:     f.Services      || '',
    };
  } catch (e) { console.error('getSampleMeta error:', e.message); return null; }
}

// ── SharePoint List: Clients ──────────────────────────────────────────────────
// Replaces Google Sheets getClientInfo()
// Columns: Title (ClientName), ClientCode, Abbrev, Email, Aliases, Phone
async function getClientInfo(customerName, token) {
  const empty = { email: '', phone: '', clientCode: '', abbrev: '' };
  if (!customerName) return empty;
  try {
    const siteId = process.env.SP_SITE_ID;
    // Fetch all clients (list is small, <200 rows)
    const res = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields($select=Title,ClientCode,Abbrev,Email,Aliases,Phone)&$top=500`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return empty;
    const data  = await res.json();
    const name  = customerName.toLowerCase().trim();
    const match = (data.value || []).find(item => {
      const f  = item.fields || {};
      const cn = (f.Title   || '').toLowerCase().trim();
      const al = (f.Aliases || '').toLowerCase();
      return cn === name || al.split(/[,;]/).map(s => s.trim()).some(a => a && (name.includes(a) || a.includes(name)));
    });
    if (!match) return empty;
    const f = match.fields || {};
    return {
      email:      f.Email      || '',
      clientCode: f.ClientCode || '',
      abbrev:     f.Abbrev     || '',
      phone:      f.Phone      || '',
    };
  } catch { return empty; }
}

// ── Result formatting ─────────────────────────────────────────────────────────
function formatResult(rawVal, rl, decimals) {
  if (rawVal === '' || rawVal === null || rawVal === undefined) return '';
  const n = parseFloat(rawVal);
  if (isNaN(n)) return String(rawVal).trim();
  if (rl !== null && rl !== undefined && n < rl) return `<${rl}`;
  if (decimals !== undefined && decimals !== null) return n.toFixed(decimals);
  return parseFloat(n.toFixed(6)).toString();
}

function resultColor(paramName, displayVal, epa) {
  if (!displayVal && displayVal !== 0) return 'none';
  const s = String(displayVal);
  if (s.startsWith('<')) return 'green';
  if (paramName === 'pH Electrometric') {
    const n = parseFloat(s);
    return isNaN(n) ? 'none' : (n >= 6.5 && n <= 8.5) ? 'green' : 'red';
  }
  if (epa === null || epa === undefined || epa === '') return 'blue';
  const n = parseFloat(s);
  if (isNaN(n)) return 'blue';
  return n <= parseFloat(epa) ? 'green' : 'red';
}

// ── Main Azure Function handler ───────────────────────────────────────────────
app.http('generate-report', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {

    try {
      const body = await request.json().catch(() => ({}));
      const { labId, meta: frontendMeta } = body;
      if (!labId) return { status: 400, jsonBody: { error: 'labId required' } };

      const baseId  = String(labId).match(/(\d{6}-\d{3})/)?.[1];
      if (!baseId)  return { status: 400, jsonBody: { error: `Invalid Lab ID: ${labId}` } };

      const dateStr = extractDateStr(baseId);
      if (!dateStr) return { status: 400, jsonBody: { error: `Cannot extract date from Lab ID: ${baseId}` } };

      const di  = dateInfo(dateStr);
      const log = [`Lab ID: ${baseId}`, `Date: ${dateStr}`, `Month folder: ${di.monthFolder}`];

      // ── Get MS Graph token (cached) ──────────────────────────────────────────
      const token = await getCachedToken();
      log.push('auth OK');

      const CTRL_FOLDER = process.env.SP_CONTROL_FOLDER || '/sites/Laboratory/Shared Documents/Documents/Test C';
      const ICPM_FOLDER = process.env.SP_ICPMS_FOLDER   || '/sites/Laboratory/Shared Documents/Documents/Test M';
      const ACID_FOLDER = process.env.SP_ACID_FOLDER    || '/sites/Laboratory/Shared Documents/Documents/metals prep';

      const ctrlMonthFolder  = `${CTRL_FOLDER}/${di.monthFolder}`;
      const icpmsMonthFolder = `${ICPM_FOLDER}/${di.monthFolder}`;

      // ── Fire all data fetches in parallel ────────────────────────────────────
      // Sample meta + client info from SP Lists; instrument files from SP drive
      // Azure Functions → SharePoint is same-network: <1s per call vs 5-10s cross-cloud
      const [
        metaRaw,
        ctrlBuf,
        icpms1Buf,
        icpms2Buf,
        acidFiles,
      ] = await Promise.allSettled([
        frontendMeta
          ? Promise.resolve(null)
          : getSampleMeta(baseId, token),
        tryDownloadByPath(ctrlMonthFolder,  `C_${dateStr}.xlsx`,    token),
        tryDownloadByPath(icpmsMonthFolder, `M_${dateStr}-01.xlsx`, token),
        tryDownloadByPath(icpmsMonthFolder, `M_${dateStr}-02.xlsx`, token),
        listFolder(ACID_FOLDER, token),
      ]);

      // ── Resolve sample meta ───────────────────────────────────────────────────
      const meta = frontendMeta
        ? {
            customer:     frontendMeta.customer     || '',
            dateDrawn:    frontendMeta.dateDrawn     || '',
            timeDrawn:    frontendMeta.timeDrawn     || '',
            dateReceived: frontendMeta.dateReceived  || '',
            timeReceived: frontendMeta.timeReceived  || '',
            location:     frontendMeta.location      || '',
            city:         frontendMeta.city          || '',
            state:        frontendMeta.state         || 'ME',
            zip:          frontendMeta.zip           || '',
            services:     (frontendMeta.tests || []).join(', '),
          }
        : (metaRaw.status === 'fulfilled' ? metaRaw.value : null) || {};

      // ── Resolve client info ───────────────────────────────────────────────────
      const clientInfo = await getClientInfo(meta.customer || '', token);

      // ── Parse control sheet ───────────────────────────────────────────────────
      let ctrlData = {};
      const csBuf  = ctrlBuf.status === 'fulfilled' ? ctrlBuf.value : null;
      if (csBuf) {
        ctrlData = parseControlSheet(csBuf, baseId);
        log.push(`CS: ${Object.keys(ctrlData.gallery || {}).length} gallery, pH=${ctrlData.phValue || '-'}, coliform=${ctrlData.coliformMPN || '-'}`);
      } else {
        log.push(`CS: not found at C_${dateStr}.xlsx`);
      }

      // ── Parse ICP-MS ──────────────────────────────────────────────────────────
      let icpmsResult = null;
      const icpmsBuf  = (icpms1Buf.status === 'fulfilled' && icpms1Buf.value) ? icpms1Buf.value
                      : (icpms2Buf.status === 'fulfilled' && icpms2Buf.value) ? icpms2Buf.value : null;
      if (icpmsBuf) {
        icpmsResult = parseICPMS(icpmsBuf, baseId);
        log.push(icpmsResult ? `ICPMS: ${Object.keys(icpmsResult.elements).length} elements` : 'ICPMS: file found, ID not matched');
      } else {
        log.push(`ICPMS: not found at M_${dateStr}-01/02.xlsx`);
      }

      // ── Parse acid sheet ──────────────────────────────────────────────────────
      let acidPrepDT = null;
      const acidList = acidFiles.status === 'fulfilled' ? (acidFiles.value || []) : [];
      const acidFile = acidList.find(f => /\.(xlsx|xls|xlsm)$/i.test(f.name));
      if (acidFile) {
        try {
          const buf = await downloadFile(acidFile.id, token);
          acidPrepDT = parseAcidSheet(buf, baseId, di.monthShort);
          log.push(`acid: ${acidPrepDT || 'none'}`);
        } catch (e) { log.push(`acid error: ${e.message}`); }
      }

      // ── Resolve pH ────────────────────────────────────────────────────────────
      // pH comes from control sheet (preferred) — pH file fallback not needed when control sheet present
      const phResult = ctrlData.phValue
        ? { value: ctrlData.phValue, analDT: ctrlData.phDT || '' }
        : null;

      // ── Test packages & params ────────────────────────────────────────────────
      const services = meta.services
        ? meta.services.split(/[,;]/).map(s => s.trim()).filter(Boolean)
        : [];
      const isRadon  = services.some(s => s.toLowerCase().includes('radon water'));
      const needsFHA = services.some(s => NEEDS_FHA_TYPES.includes(s));

      const activeParams = PARAM_CONFIG.filter(p => p.packages.some(pkg => services.includes(pkg)));
      const fhaParams    = PARAM_CONFIG.filter(p => FHA_PARAM_NAMES.includes(p.name));
      const icpmsAnalDT  = icpmsResult?.acqTime || '';

      // ── Build param rows ──────────────────────────────────────────────────────
      const buildRow = (p) => {
        let rawVal = '', analDT = '', prepDT = '';
        switch (p.source) {
          case 'gallery': {
            const gd = ctrlData.gallery?.[p.name] || {};
            rawVal = gd.value  || '';
            analDT = gd.analDT || '';
            break;
          }
          case 'icpms': {
            const v = icpmsResult?.elements?.[p.icpmsCol];
            rawVal  = v !== undefined ? String(v) : '';
            analDT  = icpmsAnalDT;
            prepDT  = acidPrepDT || '';
            break;
          }
          case 'ph': {
            rawVal = phResult?.value  || '';
            analDT = phResult?.analDT || '';
            break;
          }
          case 'control': {
            rawVal = p.name === 'Total Coliform' ? ctrlData.coliformMPN || ''
                   : p.name === 'E. Coli'        ? ctrlData.ecoliMPN    || ''
                   : '';
            prepDT = ctrlData.coliformPrepDT || '';
            analDT = ctrlData.coliformAnalDT || '';
            break;
          }
          case 'calc': {
            const ca = parseFloat(icpmsResult?.elements?.['Ca 43 (mg/L)'] ?? '');
            const mg = parseFloat(icpmsResult?.elements?.['Mg 24 (mg/L)'] ?? '');
            if (!isNaN(ca) && !isNaN(mg)) {
              rawVal = (Math.round((ca * 2.497 + mg * 4.118) * 100) / 100).toString();
            }
            analDT = icpmsAnalDT;
            prepDT = acidPrepDT || '';
            break;
          }
        }
        const display = formatResult(rawVal, p.rl, p.decimals);
        return {
          name:     p.name,
          value:    display,
          rl:       p.rl  !== null && p.rl  !== undefined ? String(p.rl)  : '',
          epa:      p.epa !== null && p.epa !== undefined ? String(p.epa) : '',
          unit:     p.unit,
          method:   p.method,
          prepDT,
          analDT,
          time:     analDT,
          color:    resultColor(p.name, display, p.epa),
          source:   p.source,
          icpmsCol: p.icpmsCol || '',
        };
      };

      const paramRows = activeParams.map(buildRow);
      const fhaRows   = needsFHA ? fhaParams.map(buildRow) : [];

      // ── Dates ─────────────────────────────────────────────────────────────────
      const now         = new Date();
      const todayStr    = `${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')}/${String(now.getFullYear()).slice(-2)}`;
      const dtCollected = meta.dateDrawn    ? combineDT(meta.dateDrawn,    meta.timeDrawn)    : '';
      const dtReceived  = meta.dateReceived ? combineDT(meta.dateReceived, meta.timeReceived) : '';

      return {
        status: 200,
        jsonBody: {
          success:    true,
          labId:      baseId,
          isRadon,
          needsFHA,
          reportType: isRadon ? 'RW' : 'COA',
          today:      todayStr,
          log,
          meta: {
            customer:     meta.customer     || '',
            email:        clientInfo.email,
            phone:        clientInfo.phone      || '',
            clientCode:   clientInfo.clientCode || '',
            abbrev:       clientInfo.abbrev     || '',
            location:     meta.location     || '',
            city:         meta.city         || '',
            state:        meta.state        || 'ME',
            zip:          meta.zip          || '',
            dtCollected,
            dtReceived,
            dateDrawn:    meta.dateDrawn    || '',
            timeDrawn:    meta.timeDrawn    || '',
            dateReceived: meta.dateReceived || '',
            timeReceived: meta.timeReceived || '',
            labId:        baseId,
          },
          services,
          paramRows,
          fhaRows,
          radon: { display: '', raw: 0, color: 'green', time: '' },
        },
      };

    } catch (err) {
      context.log(`[generate-report] fatal: ${err.message}`);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
