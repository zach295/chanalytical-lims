// netlify/functions/generate-report.js
// Builds COA report data by pulling from:
//   SharePoint: Control sheet, pH files, ICP-MS files, Acid sheet
//   Google Sheets: Archived Intake (sample meta), Clients (email)
//
// Required env vars:
//   MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, SP_SITE_ID
//   SP_CONTROL_FOLDER  (/sites/Laboratory/Shared Documents/Documents/Test C)
//   SP_ICPMS_FOLDER    (/sites/Laboratory/Shared Documents/Documents/Test M)
//   SP_ACID_FOLDER     (/sites/Laboratory/Shared Documents/Documents/Test P)
//   SPREADSHEET_ID + GOOGLE_* service account creds

const XLSX = require('xlsx');
const { getSheets, SPREADSHEET_ID, SHEETS } = require('./sheets-auth');
const GRAPH = 'https://graph.microsoft.com/v1.0';

// ─────────────────────────────────────────────────────────────────────────────
// PARAMETER CONFIG
// source: gallery=control sheet Gallery cols, icpms=ICP-MS Concentrations tab,
//         ph=pH file, control=control sheet Coliform cols, calc=calculated
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
  { name:'pH Electrometric',             rl:null,    epa:'6.5-8.5', unit:'',     method:'SM4500H+B',        source:'ph',      decimals:2,                                        packages:['Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive','Pro Plus'] },
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

// ICP-MS columns we care about (exact header names in Concentrations tab)
const ICPMS_COLS = ['Na 23 (mg/L)','Mg 24 (mg/L)','Ca 43 (mg/L)','Cr 52 (mg/L)',
                    'Fe 54 (mg/L)','Mn 55 (mg/L)','Co 59 (mg/L)','Cu 63 (mg/L)',
                    'As 75 (ug/L)','Cd 111 (mg/L)','Sb 121 (mg/L)','Pb 208 (ug/L)','U 238 (ug/L)'];

// ── MS Graph ───────────────────────────────────────────────────────────────────
async function getToken() {
  const { MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET } = process.env;
  if (!MS_TENANT_ID) throw new Error('MS_TENANT_ID not set');
  const res = await fetch(
    `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`,
    { method:'POST', headers:{'Content-Type':'application/x-www-form-urlencoded'},
      body: new URLSearchParams({ grant_type:'client_credentials', client_id:MS_CLIENT_ID,
        client_secret:MS_CLIENT_SECRET, scope:'https://graph.microsoft.com/.default' }) }
  );
  const d = await res.json();
  if (!d.access_token) throw new Error(`Auth failed: ${d.error_description}`);
  return d.access_token;
}

async function graphGet(path, token) {
  const res = await fetch(`${GRAPH}${path}`, { headers:{ Authorization:`Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph GET ${path} → ${res.status}: ${await res.text()}`);
  return res.json();
}

async function graphGetBytes(path, token) {
  const res = await fetch(`${GRAPH}${path}`, { headers:{ Authorization:`Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph download → ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

// ── Date helpers ───────────────────────────────────────────────────────────────
function extractDateStr(labId) {
  return String(labId).match(/^(\d{6})/)?.[1] || null;
}

function dateInfo(mmddyy) {
  const mm  = mmddyy.slice(0,2), dd = mmddyy.slice(2,4), yy = mmddyy.slice(4,6);
  const idx = parseInt(mm) - 1;
  return {
    mm, dd, yy,
    year:        2000 + parseInt(yy),
    monthLong:   MONTH_LONG[idx],
    monthShort:  MONTH_TABS[idx],
    monthFolder: `${MONTH_LONG[idx]} ${2000 + parseInt(yy)}`,
  };
}

// Format a raw datetime value (string or Date) → MM/DD/YY HH:MM (24hr)
function fmtDT(raw) {
  if (!raw) return '';
  const s = String(raw).trim();
  if (!s) return '';

  // Already formatted MM/DD/YY HHMM (no colon) → add colon
  if (/^\d{2}\/\d{2}\/\d{2}\s+\d{4}$/.test(s)) {
    return s.replace(/(\d{2}\/\d{2}\/\d{2}\s+)(\d{2})(\d{2})$/, '$1$2:$3');
  }

  // Already formatted MM/DD/YY HH:MM (with colon) → keep as-is
  if (/^\d{2}\/\d{2}\/\d{2}\s+\d{2}:\d{2}$/.test(s)) return s;

  // JS-parseable datetime string (e.g. "3/31/2026 10:51:07 AM" from XLSX)
  const hasTime = /\d{1,2}:\d{2}/.test(s);
  if (hasTime) {
    const dt = new Date(s);
    if (!isNaN(dt.getTime())) {
      const m  = String(dt.getMonth()+1).padStart(2,'0');
      const d  = String(dt.getDate()).padStart(2,'0');
      const y  = String(dt.getFullYear()).slice(-2);
      const h  = String(dt.getHours()).padStart(2,'0');
      const mi = String(dt.getMinutes()).padStart(2,'0');
      return `${m}/${d}/${y} ${h}:${mi}`;
    }
  }
  return s;
}

// Combine separate date + time strings → MM/DD/YY HHMM (24hr)
function combineDT(dateStr, timeStr) {
  if (!dateStr) return '';
  const ds = String(dateStr).trim();
  const ts = String(timeStr || '').trim();

  const dm = ds.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!dm) return ds;
  const m = dm[1].padStart(2,'0');
  const d = dm[2].padStart(2,'0');
  const y = dm[3].length === 4 ? dm[3].slice(-2) : dm[3].padStart(2,'0');

  let hhmm = '';
  if (ts) {
    const ampm = ts.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
    const h24  = ts.match(/^(\d{1,2}):(\d{2})/);
    if (ampm) {
      let h = parseInt(ampm[1]);
      const mi = ampm[2];
      if (ampm[3].toUpperCase() === 'PM' && h !== 12) h += 12;
      if (ampm[3].toUpperCase() === 'AM' && h === 12) h = 0;
      hhmm = `${String(h).padStart(2,'0')}:${mi}`;
    } else if (h24) {
      hhmm = `${String(parseInt(h24[1])).padStart(2,'0')}:${h24[2]}`;
    } else {
      hhmm = ts.replace(/[^0-9]/g,'').slice(0,4);
    }
  }
  return hhmm ? `${m}/${d}/${y} ${hhmm}` : `${m}/${d}/${y}`;
}

// ── SharePoint file helpers ────────────────────────────────────────────────────
// Convert full SharePoint path → Graph API drive-relative path with proper encoding
// e.g. /sites/Laboratory/Shared Documents/Documents/Test C → Documents/Test%20C
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
    const data = await graphGet(
      `/sites/${siteId}/drive/root:/${drivePath}:/children?$select=id,name`,
      token
    );
    return data.value || [];
  } catch(e) {
    console.warn(`listFolder failed for ${folderPath} (drive path: ${drivePath}):`, e.message);
    return [];
  }
}

// Try month subfolder (e.g. "July 2026"), then short form ("Jul 2026"), then base folder
async function resolveMonthFolder(baseFolder, mmddyy, token) {
  const mm   = parseInt(mmddyy.slice(0,2));
  const yy   = parseInt(mmddyy.slice(4,6));
  const year = 2000 + yy;
  const d    = new Date(year, mm-1, 1);
  const longM  = d.toLocaleString('en-US', { month:'long'  }); // July
  const shortM = d.toLocaleString('en-US', { month:'short' }); // Jul
  const candidates = [
    `${baseFolder}/${longM} ${year}`,
    `${baseFolder}/${shortM} ${year}`,
    baseFolder,
  ];
  // Try all candidates IN PARALLEL
  const results = await Promise.allSettled(
    candidates.map(async folder => {
      const files = await listFolder(folder, token);
      if (!files.length) throw new Error('empty');
      return { folder, files };
    })
  );
  const winner = results.find(r => r.status === 'fulfilled');
  if (winner) return winner.value;
  return { folder: baseFolder, files: [] };
}


async function findFiles(folderPath, dateStr, prefixHint, token) {
  const files = await listFolder(folderPath, token);
  const d = dateStr.toLowerCase();
  const p = (prefixHint || '').toLowerCase();
  return files.filter(f => {
    const n = f.name.toLowerCase();
    return n.includes(d) && (!p || n.includes(p)) && (n.endsWith('.xlsx') || n.endsWith('.xls') || n.endsWith('.xlsm'));
  });
}

async function downloadFile(fileId, token) {
  const siteId = process.env.SP_SITE_ID;
  return graphGetBytes(`/sites/${siteId}/drive/items/${fileId}/content`, token);
}

// ── Excel helpers ──────────────────────────────────────────────────────────────
function readSheetRows(buffer, sheetName, headerRowNum) {
  const wb  = XLSX.read(buffer, { type:'buffer', cellDates:true, raw:false });
  let ws, sheetNames = wb.SheetNames;
  if (sheetName) {
    const sn = wb.SheetNames.find(n => n === sheetName)
             || wb.SheetNames.find(n => n.toLowerCase() === (sheetName||'').toLowerCase());
    if (!sn) return { rows:[], found:false, sheetNames };
    ws = wb.Sheets[sn];
  } else {
    ws = wb.Sheets[wb.SheetNames[0]];
  }
  const hIdx = Math.max((headerRowNum || 1) - 1, 0);
  const aoa  = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', raw:false });
  if (aoa.length <= hIdx) return { rows:[], found:true, sheetNames };
  const headers = aoa[hIdx].map(h => String(h||'').replace(/[\n\r]+/g,' ').replace(/\s+/g,' ').trim());
  const rows = aoa.slice(hIdx+1).map(row => {
    const obj = {};
    headers.forEach((h,i) => { if(h) obj[h] = String(row[i]??'').trim(); });
    return obj;
  }).filter(r => Object.values(r).some(v => v));
  return { rows, found:true, headers, sheetNames };
}

function readAoA(buffer) {
  const wb  = XLSX.read(buffer, { type:'buffer', cellDates:true, raw:false });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header:1, defval:'', raw:false });
}

// ── pH Parser ──────────────────────────────────────────────────────────────────
// Headers row 9 (1-indexed). 
// Rule: Find D row (Code ends in ' D') for this Lab ID where QC = Pass.
//       Report the pH and Date/Time from the matching NON-D row (same code number).
function parsePH(buffer, baseId) {
  const { rows } = readSheetRows(buffer, null, 9);

  // Build map: code → row data
  const byCode = {};
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

  // Find a D row for this Lab ID that passed → return the non-D row's pH + datetime
  for (const [code, dRow] of Object.entries(byCode)) {
    if (!code.endsWith(' D')) continue;
    if (dRow.sampleBase !== baseId) continue;
    if (!dRow.pass) continue;
    const nonDCode = code.replace(/ D$/, '');
    const nonDRow  = byCode[nonDCode];
    if (nonDRow?.ph) {
      return { value: nonDRow.ph, analDT: fmtDT(nonDRow.dt) };
    }
  }
  return null;
}

// ── Control Sheet Parser ───────────────────────────────────────────────────────
// Headers row 1. Col A = Barcode.
// Gallery Date/Time = always the column immediately after its result column.
function parseControlSheet(buffer, baseId) {
  const aoa = readAoA(buffer);
  if (!aoa.length) return {};

  // Find header row (contains 'Barcode' or 'Lab ID')
  let hIdx = -1, headers = [], headersLower = [];
  for (let i = 0; i < Math.min(aoa.length, 5); i++) {
    const row = aoa[i].map(c => String(c||'').replace(/[\n\r]+/g,' ').replace(/\s+/g,' ').trim());
    if (row.some(c => ['barcode','lab id','sample id'].includes(c.toLowerCase()))) {
      hIdx = i; headers = row;
      // Normalize headers: replace newlines with spaces, collapse spaces, lowercase
      headersLower = headers.map(h => h.replace(/[\n\r]+/g,' ').replace(/\s+/g,' ').toLowerCase().trim());
      break;
    }
  }
  if (hIdx < 0) {
    console.warn('[parseControlSheet] Header row not found. First row:', aoa[0]?.slice(0,5));
    return {};
  }

  const barcodeColIdx = headersLower.findIndex(h => ['barcode','lab id','sample id'].includes(h));

  // Flexible column finder — matches ignoring case and units suffix
  const findCol = (key) => {
    const keyLower = key.replace(/[\n\r]+/g,' ').replace(/\s+/g,' ').toLowerCase().trim();
    // Exact match first
    let idx = headersLower.indexOf(keyLower);
    if (idx >= 0) return idx;
    // Partial match — key starts with header or header starts with key
    idx = headersLower.findIndex(h => h && (h.startsWith(keyLower.split(' ')[0]) && keyLower.split(' ')[0].length > 3));
    return idx;
  };

  // Find the matching Lab ID row
  for (let r = hIdx+1; r < aoa.length; r++) {
    const row = aoa[r];
    const barcode = String(row[barcodeColIdx] || '').trim();
    if (!barcode.match(/\d{6}-\d{3}/)) continue;
    const barcodeBase = barcode.match(/(\d{6}-\d{3})/)?.[1];
    if (barcodeBase !== baseId) continue;

    // Helper: get column value by header name
    const getH = (header) => {
      const idx = findCol(header);
      return idx >= 0 ? String(row[idx] || '').trim() : '';
    };

    // Gallery results: use flexible matching, Date/Time is next column after result
    const gallery = {};
    PARAM_CONFIG.filter(p => p.source === 'gallery').forEach(p => {
      const colIdx = findCol(p.colKey);
      if (colIdx < 0) {
        console.warn(`[parseControlSheet] Column not found: "${p.colKey}"`);
        return;
      }
      const value  = String(row[colIdx]   || '').trim();
      const dtVal  = String(row[colIdx+1] || '').trim();
      gallery[p.name] = { value, analDT: fmtDT(dtVal) };
    });

    const pHColIdx   = headersLower.findIndex(h => h === 'ph');
    const pHDTColIdx = pHColIdx >= 0 ? pHColIdx + 1 : -1;

    const result = {
      gallery,
      coliformMPN:    getH('Coliform MPN'),
      ecoliMPN:       getH('Ecoli MPN'),
      coliformPrepDT: fmtDT(getH('Start Date/Time')),
      coliformAnalDT: fmtDT(getH('End Date/Time')),
      phValue:        pHColIdx >= 0 ? String(row[pHColIdx] || '').trim() : '',
      phDT:           pHDTColIdx >= 0 ? fmtDT(String(row[pHDTColIdx] || '').trim()) : '',
    };

    console.log(`[parseControlSheet] Found ${baseId}: gallery=${Object.keys(gallery).length}, pH=${result.phValue}, coliform=${result.coliformMPN}`);
    return result;
  }

  console.warn(`[parseControlSheet] Lab ID ${baseId} not found in control sheet`);
  return {};
}


// ── ICP-MS Parser ──────────────────────────────────────────────────────────────
// Sheet: 'Concentrations', headers row 1.
// Only rows where QC Status = 'Passed' and Sample Id matches MMDDYY-### (± dilution suffix).
// Acquisition Time = analysis datetime for all metals.
// For diluted rows (x4, x10): base row takes priority; dilution fills gaps only.
function parseICPMS(buffer, baseId) {
  const { rows, found } = readSheetRows(buffer, 'Concentrations', 1);
  if (!found || !rows.length) return null;

  const entries = [];
  for (const row of rows) {
    const rawId = (row['Sample Id'] || row['Sample ID'] || '').trim();
    if (!rawId) continue;
    const idBase = rawId.match(/(\d{6}-\d{3})/)?.[1];
    if (!idBase || idBase !== baseId) continue;

    // Skip rows with Failed QC
    const qc = (row['QC Status'] || '').trim().toLowerCase();
    if (qc && qc !== 'passed') continue;

    const dilMatch  = rawId.match(/x(\d+)/i);
    const dilFactor = dilMatch ? parseInt(dilMatch[1]) : 0;
    const acqTime   = fmtDT(row['Acquisition Time'] || '');

    const elements = {};
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

  // Sort: base (0) first, then ascending dilution
  entries.sort((a,b) => a.dilFactor - b.dilFactor);

  // Merge: base values take priority; dilutions fill only missing elements
  const merged = {};
  let acqTime = '';
  for (const entry of entries) {
    if (!acqTime && entry.acqTime) acqTime = entry.acqTime;
    for (const [col, val] of Object.entries(entry.elements)) {
      if (!(col in merged)) merged[col] = val;
    }
  }
  return { elements: merged, acqTime };
}

// ── Acid Sheet Parser ──────────────────────────────────────────────────────────
// File: 'metals prep YYYY V1' in Test P folder (single file, updated yearly).
// Tab: 'acidification [Month]' or 'acidifcation [Month]' (some months have typo).
// Columns: A=date, B=time, C=initials, D=lot#, E=amt added, F=sample ID, G=notes
// Match Lab ID in col F (strip any suffix letters after MMDDYY-###).
// Returns combined prep Date/Time string.
function parseAcidSheet(buffer, baseId, monthShort) {
  const wb = XLSX.read(buffer, { type:'buffer', cellDates:true, raw:false });

  // Try multiple tab name variants (typo exists on some months)
  const candidates = [
    `acidification ${monthShort}`,
    `acidifcation ${monthShort}`,
    `Acidification ${monthShort}`,
    `Acidifcation ${monthShort}`,
  ];
  let ws = null;
  for (const name of candidates) {
    const found = wb.SheetNames.find(n => n.toLowerCase() === name.toLowerCase());
    if (found) { ws = wb.Sheets[found]; break; }
  }
  // Last resort: any tab containing the month abbreviation
  if (!ws) {
    const found = wb.SheetNames.find(n => n.toLowerCase().includes(monthShort.toLowerCase()));
    if (found) ws = wb.Sheets[found];
  }
  if (!ws) {
    console.warn(`Acid sheet: no tab found for month ${monthShort}. Available: ${wb.SheetNames.join(', ')}`);
    return null;
  }

  const aoa = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', raw:false });

  // Find header row (has 'date' and 'sample' columns)
  let hIdx = -1;
  for (let i = 0; i < Math.min(aoa.length, 5); i++) {
    const row = aoa[i].map(c => String(c||'').toLowerCase().trim());
    if (row.some(c => c === 'date') && row.some(c => c.includes('sample'))) { hIdx = i; break; }
  }
  // Columns are A=0,B=1,F=5 regardless of header row
  const COL_DATE = 0, COL_TIME = 1, COL_SAMPLE = 5;
  const startRow = hIdx >= 0 ? hIdx + 1 : 2;

  for (let r = startRow; r < aoa.length; r++) {
    const row       = aoa[r];
    const sampleRaw = String(row[COL_SAMPLE] || '').trim();
    if (!sampleRaw) continue;
    const cellBase  = sampleRaw.match(/(\d{6}-\d{3})/)?.[1];
    if (!cellBase || cellBase !== baseId) continue;

    const dateStr = String(row[COL_DATE] || '').trim();
    const timeStr = String(row[COL_TIME] || '').trim();
    return combineDT(dateStr, timeStr);
  }
  return null;
}

// ── Google Sheets lookups ──────────────────────────────────────────────────────
// Archived Intake: A=ts, B=fullId, C=coaTest, D=customer, E=dateDrawn, F=timeDrawn,
//                  G=receivedDate, H=receivedTime, I=location, J=city, K=state,
//                  L=zip, M=reviewedBy, N=notes
async function getSampleMeta(sheets, baseId) {
  try {
    const res  = await sheets.spreadsheets.values.get({ spreadsheetId:SPREADSHEET_ID, range:`'${SHEETS.ARCHIVED_INTAKE}'!A:N` });
    const rows = (res.data.values || []).slice(1);
    const row  = rows.find(r => (r[1]||'').trim().startsWith(baseId));
    if (!row) return null;
    return {
      customer:     row[3]  || '',
      dateDrawn:    row[4]  || '',
      timeDrawn:    row[5]  || '',
      dateReceived: row[6]  || '',
      timeReceived: row[7]  || '',
      location:     row[8]  || '',
      city:         row[9]  || '',
      state:        row[10] || 'ME',
      zip:          row[11] || '',
      services:     row[2]  || '',
    };
  } catch(e) { console.error('getSampleMeta error:', e.message); return null; }
}

// Clients: A=ClientName, B=ClientCode, C=Abbrev, D=Email, E=Aliases, F=Phone
async function getClientInfo(sheets, customerName) {
  if (!customerName) return { email: '', phone: '', clientCode: '', abbrev: '' };
  try {
    const res  = await sheets.spreadsheets.values.get({ spreadsheetId:SPREADSHEET_ID, range:`'${SHEETS.CLIENTS}'!A:G` });
    const rows = (res.data.values || []).slice(1);
    const name = customerName.toLowerCase().trim();
    const match = rows.find(r => {
      const cn = (r[0]||'').toLowerCase().trim();
      const al = (r[4]||'').toLowerCase();
      return cn === name || al.split(/[,;]/).map(s=>s.trim()).some(a => a && (name.includes(a) || a.includes(name)));
    });
    if (!match) return { email: '', phone: '', clientCode: '', abbrev: '' };
    return {
      email:      match[3] || '',
      clientCode: match[1] || '',
      abbrev:     match[2] || '',
      phone:      match[5] || '',
    };
  } catch { return { email: '', phone: '', clientCode: '', abbrev: '' }; }
}

// Keep old function name for backward compatibility
async function getClientEmail(sheets, customerName) {
  return (await getClientInfo(sheets, customerName)).email;
}

// ── Result formatting ──────────────────────────────────────────────────────────
function formatResult(rawVal, rl, decimals) {
  if (rawVal === '' || rawVal === null || rawVal === undefined) return '';
  const n = parseFloat(rawVal);
  if (isNaN(n)) return String(rawVal).trim();
  if (rl !== null && rl !== undefined && n < rl) return `<${rl}`;
  // Use specified decimal places if provided, otherwise trim trailing zeros
  if (decimals !== undefined && decimals !== null) return n.toFixed(decimals);
  return parseFloat(n.toFixed(6)).toString();
}

function resultColor(paramName, displayVal, epa) {
  if (!displayVal && displayVal !== 0) return 'none';
  const s = String(displayVal);
  // <RL values are always below EPA limit
  if (s.startsWith('<')) return 'green';
  if (paramName === 'pH Electrometric') {
    const n = parseFloat(s);
    return isNaN(n) ? 'none' : (n >= 6.5 && n <= 8.5) ? 'green' : 'red';
  }
  // No EPA limit → blue (See Notation)
  if (epa === null || epa === undefined || epa === '') return 'blue';
  const n = parseFloat(s);
  if (isNaN(n)) return 'blue';
  return n <= parseFloat(epa) ? 'green' : 'red';
}

// ── Main Handler ───────────────────────────────────────────────────────────────
exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') return { statusCode:405, body:'Method Not Allowed' };

  try {
    const { labId, meta: frontendMeta } = JSON.parse(event.body || '{}');
    if (!labId) return { statusCode:400, body:JSON.stringify({ error:'labId required' }) };

    const baseId  = String(labId).match(/(\d{6}-\d{3})/)?.[1];
    if (!baseId)  return { statusCode:400, body:JSON.stringify({ error:`Invalid Lab ID: ${labId}` }) };

    const dateStr = extractDateStr(baseId);
    if (!dateStr) return { statusCode:400, body:JSON.stringify({ error:`Cannot extract date from Lab ID: ${baseId}` }) };

    const di  = dateInfo(dateStr);
    const log = [`Lab ID: ${baseId}`, `Date: ${dateStr}`, `Month folder: ${di.monthFolder}`];

    // ── Google Sheets meta (start early, runs in parallel with SP) ────────────
    const sheets      = getSheets();
    // Use pre-loaded metadata from frontend if available (avoids reading entire Archived Intake)
    const metaPromise = frontendMeta
      ? Promise.resolve({
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
        })
      : getSampleMeta(sheets, baseId);

    // Start client info lookup in parallel (uses customer from meta when available)
    const clientPromise = frontendMeta?.customer
      ? getClientInfo(sheets, frontendMeta.customer)
      : metaPromise.then(m => getClientInfo(sheets, m?.customer || ''));

    // ── Results Cache lookup (Netlify Blobs — fast, no timeout) ─────────────
    // Power Automate writes instrument results here after each run.
    // Reads via cache-results.js which uses Netlify Blob storage.
    let phResult = null, ctrlData = {}, icpmsResult = null, acidPrepDT = null;
    try {
      const cacheRes = await fetch(
        `${process.env.URL || 'https://chanalytical.netlify.app'}/.netlify/functions/cache-results?labId=${baseId}`
      );
      if (cacheRes.ok) {
        const cacheData = await cacheRes.json();
        if (cacheData.found && cacheData.data) {
          const d = cacheData.data;
          ctrlData = {
            gallery:        d.gallery        || {},
            coliformMPN:    d.coliformMPN    || '',
            ecoliMPN:       d.ecoliMPN       || '',
            coliformPrepDT: d.coliformPrepDT || '',
            coliformAnalDT: d.coliformAnalDT || '',
            phValue:        d.ph             || '',
            phDT:           d.phDT           || '',
          };
          icpmsResult = d.icpms
            ? { elements: d.icpms, acqTime: d.icpmsAcqTime || '' }
            : null;
          acidPrepDT = d.acidPrepDT || null;
          if (d.ph) phResult = { value: d.ph, analDT: d.phDT || '' };

          const gc = Object.keys(d.gallery || {}).length;
          const mc = Object.keys(d.icpms   || {}).length;
          log.push(`✅ Cache hit: ${gc} gallery, ${mc} metals, pH=${d.ph||'—'}, coliform=${d.coliformMPN||'—'}`);
        } else {
          log.push(`ℹ️ No cache entry for ${baseId} — results blank until Power Automate runs`);
        }
      }
    } catch(e) {
      log.push(`⚠️ Cache read failed: ${e.message}`);
    }

    // ── Google Sheets ─────────────────────────────────────────────────────────
    const [meta, clientInfo] = await Promise.all([metaPromise, clientPromise]);
    const metaResolved = meta || {};
    const clientEmail = clientInfo.email;

    // ── Test packages & params ────────────────────────────────────────────────
    const services = metaResolved.services
      ? metaResolved.services.split(/[,;]/).map(s=>s.trim()).filter(Boolean)
      : [];
    const isRadon  = services.some(s => s.toLowerCase().includes('radon water'));
    const needsFHA = services.some(s => NEEDS_FHA_TYPES.includes(s));

    const activeParams = PARAM_CONFIG.filter(p => p.packages.some(pkg => services.includes(pkg)));
    const fhaParams    = PARAM_CONFIG.filter(p => FHA_PARAM_NAMES.includes(p.name));

    // Shared ICP-MS analysis date/time for all metals
    const icpmsAnalDT = icpmsResult?.acqTime || '';

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
          // Hardness = (Ca × 2.497) + (Mg × 4.118)
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
        rl:       p.rl   !== null && p.rl   !== undefined ? String(p.rl)  : '',
        epa:      p.epa  !== null && p.epa  !== undefined ? String(p.epa) : '',
        unit:     p.unit,
        method:   p.method,
        prepDT,
        analDT,
        time:     analDT,
        color:    resultColor(p.name, display, p.epa),
        source:   p.source,    // needed by dashboard to merge SP data
        icpmsCol: p.icpmsCol || '',
      };
    };

    const paramRows = activeParams.map(buildRow);
    const fhaRows   = needsFHA ? fhaParams.map(buildRow) : [];

    // ── Dates ─────────────────────────────────────────────────────────────────
    const now      = new Date();
    const todayStr = `${String(now.getMonth()+1).padStart(2,'0')}/${String(now.getDate()).padStart(2,'0')}/${String(now.getFullYear()).slice(-2)}`;
    const dtCollected = metaResolved.dateDrawn    ? combineDT(metaResolved.dateDrawn,    metaResolved.timeDrawn)    : '';
    const dtReceived  = metaResolved.dateReceived ? combineDT(metaResolved.dateReceived, metaResolved.timeReceived) : '';

    return {
      statusCode: 200,
      headers: { 'Content-Type':'application/json' },
      body: JSON.stringify({
        success: true,
        labId: baseId,
        isRadon,
        needsFHA,
        reportType: isRadon ? 'RW' : 'COA',
        today: todayStr,
        log,
        meta: {
          customer:     metaResolved.customer     || '',
          email:        clientEmail,
          phone:        clientInfo.phone      || '',
          clientCode:   clientInfo.clientCode || '',
          abbrev:       clientInfo.abbrev     || '',
          location:     metaResolved.location     || '',
          city:         metaResolved.city         || '',
          state:        metaResolved.state        || 'ME',
          zip:          metaResolved.zip          || '',
          dtCollected,
          dtReceived,
          dateDrawn:    metaResolved.dateDrawn    || '',
          timeDrawn:    metaResolved.timeDrawn    || '',
          dateReceived: metaResolved.dateReceived || '',
          timeReceived: metaResolved.timeReceived || '',
          labId: baseId,
        },
        services,
        paramRows,
        fhaRows,
        radon: { display:'', raw:0, color:'green', time:'' },
      }),
    };

  } catch(err) {
    console.error('[generate-report]', err);
    return { statusCode:500, body:JSON.stringify({ error:err.message }) };
  }
};
