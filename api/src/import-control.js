/**
 * import-control.js — Azure version
 * 1. Reads Results Cache to get Lab IDs needing control sheet data
 * 2. Groups IDs by date (MMDDYY from base ID)
 * 3. Finds the matching control sheet file for each date
 * 4. Writes pH, bacteria, Gallery chemistry to Results Cache
 */
const { app }    = require('@azure/functions');
const { getToken, listItems, createItem, updateItem, LISTS } = require('../shared/graph');
const GRAPH = 'https://graph.microsoft.com/v1.0';

let XLSX;
try { XLSX = require('xlsx'); } catch(e) { console.warn('[import-control] xlsx not available:', e.message); }

// Convert any date/time value to military time "MM/DD/YY HH:MM"
function toMilitaryDT(val) {
  if (!val && val !== 0) return '';
  // Excel serial number (number type)
  if (typeof val === 'number') {
    const ms      = (val - 25569) * 86400 * 1000;
    const d       = new Date(Math.round(ms));
    const mm      = String(d.getUTCMonth()+1).padStart(2,'0');
    const dd      = String(d.getUTCDate()).padStart(2,'0');
    const yy      = String(d.getUTCFullYear()).slice(-2);
    const hh      = String(d.getUTCHours()).padStart(2,'0');
    const min     = String(d.getUTCMinutes()).padStart(2,'0');
    return `${mm}/${dd}/${yy} ${hh}:${min}`;
  }
  // JS Date object
  if (val instanceof Date) {
    const mm  = String(val.getMonth()+1).padStart(2,'0');
    const dd  = String(val.getDate()).padStart(2,'0');
    const yy  = String(val.getFullYear()).slice(-2);
    const hh  = String(val.getHours()).padStart(2,'0');
    const min = String(val.getMinutes()).padStart(2,'0');
    return `${mm}/${dd}/${yy} ${hh}:${min}`;
  }
  // String — strip AM/PM and normalize
  const s     = String(val).trim();
  const ampm  = s.match(/^(\d{1,2}\/\d{1,2}\/\d{2,4})\s+(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
  if (ampm) {
    let [, datePart, h, m, ap] = ampm;
    h  = parseInt(h, 10);
    ap = (ap||'').toUpperCase();
    if (ap === 'PM' && h < 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
    datePart = datePart.replace(/(\d{1,2}\/\d{1,2}\/)(\d{4})/, (_, p1, y) => p1 + y.slice(-2));
    return `${datePart} ${String(h).padStart(2,'0')}:${m}`;
  }
  return s;
}

async function graphListFolder(folderRelPath, token) {
  const siteId = process.env.SP_SITE_ID;
  const enc    = folderRelPath.split('/').map(encodeURIComponent).join('/');
  const res    = await fetch(
    `${GRAPH}/sites/${siteId}/drive/root:/${enc}:/children?$select=id,name,file&$top=500`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`listFolder ${res.status}`);
  return (await res.json()).value || [];
}

async function graphDownloadFile(itemId, token) {
  const siteId = process.env.SP_SITE_ID;
  const res    = await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${itemId}/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`downloadFile ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

const COL = {
  BARCODE:0, PH:3, DT_PH:4, COLIFORM:5, ECOLI:6,
  START_DT:8, END_DT:10,
  CHLORIDE:13, DT_CHLORIDE:14, FLUORIDE:15, DT_FLUORIDE:16,
  NITRITE:17, DT_NITRITE:18, NITRATE:19, DT_NITRATE:20,
  ALKALINITY:21, DT_ALKALINITY:22, SULFATE:23, DT_SULFATE:24,
  TANNINS:25, DT_TANNINS:26, TDS:27, DT_TDS:28,
  BROMIDE:29, DT_BROMIDE:30,
};

function cellVal(ws, r, c) {
  const cell = ws[XLSX.utils.encode_cell({ r, c })];
  if (!cell || cell.v === undefined || cell.v === null) return '';
  const v = String(cell.w || cell.v).trim();
  return (v === 'N/A' || v === '#N/A') ? '' : v;
}

function getBaseId(barcode) {
  const m = String(barcode || '').match(/^(\d{6}-\d{3})/);
  return m ? m[1] : '';
}

function getDatePart(baseId) {
  const m = String(baseId || '').match(/^(\d{6})/);
  return m ? m[1] : '';
}

function sn(val) {
  if (!val || val === '') return val;
  const n = parseFloat(val);
  if (isNaN(n)) return val;
  return n < 0 ? '0' : String(Math.round(n * 10000) / 10000);
}

function parseControlFile(buffer, targetIds) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) return [];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows = [];
  for (let r = 1; r <= range.e.r; r++) {
    const barcode = cellVal(ws, r, COL.BARCODE);
    if (!barcode) continue;
    const baseId = getBaseId(barcode);
    if (!baseId) continue;
    if (targetIds && targetIds.size > 0 && !targetIds.has(baseId)) continue;
    rows.push({
      baseId, barcode,
      ph:           cellVal(ws, r, COL.PH),
      dt_ph:        cellVal(ws, r, COL.DT_PH),
      coliform:     cellVal(ws, r, COL.COLIFORM),
      ecoli:        cellVal(ws, r, COL.ECOLI),
      start_dt:     cellVal(ws, r, COL.START_DT),
      end_dt:       cellVal(ws, r, COL.END_DT),
      chloride:     cellVal(ws, r, COL.CHLORIDE),
      dt_chloride:  cellVal(ws, r, COL.DT_CHLORIDE),
      fluoride:     cellVal(ws, r, COL.FLUORIDE),
      dt_fluoride:  cellVal(ws, r, COL.DT_FLUORIDE),
      nitrite:      cellVal(ws, r, COL.NITRITE),
      dt_nitrite:   cellVal(ws, r, COL.DT_NITRITE),
      nitrate:      cellVal(ws, r, COL.NITRATE),
      dt_nitrate:   cellVal(ws, r, COL.DT_NITRATE),
      alkalinity:   cellVal(ws, r, COL.ALKALINITY),
      dt_alkalinity:cellVal(ws, r, COL.DT_ALKALINITY),
      sulfate:      cellVal(ws, r, COL.SULFATE),
      dt_sulfate:   cellVal(ws, r, COL.DT_SULFATE),
      tannins:      cellVal(ws, r, COL.TANNINS),
      dt_tannins:   cellVal(ws, r, COL.DT_TANNINS),
      tds:          cellVal(ws, r, COL.TDS),
      dt_tds:       cellVal(ws, r, COL.DT_TDS),
      bromide:      cellVal(ws, r, COL.BROMIDE),
      dt_bromide:   cellVal(ws, r, COL.DT_BROMIDE),
    });
  }
  return rows;
}

function buildFields(row) {
  const f = {};
  if (row.ph)            f.Title    = sn(row.ph);
  if (row.dt_ph)         f.field_1 = toMilitaryDT(row.dt_ph);
  if (row.coliform)      f.field_2  = row.coliform;
  if (row.ecoli)         f.field_3  = row.ecoli;
  if (row.start_dt)      f.field_4 = toMilitaryDT(row.start_dt);
  if (row.end_dt)        f.field_5 = toMilitaryDT(row.end_dt);
  if (row.chloride)      f.field_6  = sn(row.chloride);
  if (row.dt_chloride)   f.field_7 = toMilitaryDT(row.dt_chloride);
  if (row.fluoride)      f.field_8  = sn(row.fluoride);
  if (row.dt_fluoride)   f.field_9 = toMilitaryDT(row.dt_fluoride);
  if (row.nitrite)       f.field_10 = sn(row.nitrite);
  if (row.dt_nitrite)    f.field_11 = toMilitaryDT(row.dt_nitrite);
  if (row.nitrate)       f.field_12 = sn(row.nitrate);
  if (row.dt_nitrate)    f.field_13 = toMilitaryDT(row.dt_nitrate);
  if (row.alkalinity)    f.field_14 = sn(row.alkalinity);
  if (row.dt_alkalinity) f.field_15 = toMilitaryDT(row.dt_alkalinity);
  if (row.sulfate)       f.field_16 = sn(row.sulfate);
  if (row.dt_sulfate)    f.field_17 = toMilitaryDT(row.dt_sulfate);
  if (row.tannins)       f.field_18 = sn(row.tannins);
  if (row.dt_tannins)    f.field_19 = toMilitaryDT(row.dt_tannins);
  if (row.tds)           f.field_20 = sn(row.tds);
  if (row.dt_tds)        f.field_21 = toMilitaryDT(row.dt_tds);
  if (row.bromide)       f.field_22 = sn(row.bromide);
  if (row.dt_bromide)    f.field_23 = toMilitaryDT(row.dt_bromide);
  return f;
}

app.http('import-control', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (!XLSX) return { status: 500, jsonBody: { error: 'xlsx not installed' } };
      const body = await request.json().catch(() => ({}));
      const { debug, all: importAll } = body;
      const token = await getToken();

      const rawFolder = process.env.SP_CONTROL_FOLDER ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
      const marker = 'Shared Documents/';
      const mi     = rawFolder.indexOf(marker);
      const folder = mi >= 0 ? rawFolder.slice(mi + marker.length) : rawFolder.replace(/^\/+/, '');

      // Step 1: Get Results Cache IDs needing control data
      const cacheItems = await listItems('Results Cache', { top: 500 });
      // Process rows missing pH (Title) OR coliform OR any Gallery chemistry
      const needsControl = cacheItems.filter(r => {
        const hasId = !!(r.LabID || '').trim();
        if (!hasId) return false;
        const hasPH       = !!(r.Title   || '').toString().trim();
        const hasColiform = !!(r.field_2  || '').toString().trim();
        const hasChloride = !!(r.field_6  || '').toString().trim();
        const allFilled   = hasPH && hasColiform && hasChloride;
        return importAll || !allFilled;
      }).slice(0, 30); // process max 30 per run to avoid timeout

      if (!needsControl.length) {
        return { status: 200, jsonBody: { success: true, message: 'All Results Cache entries already have control data', updated: 0 } };
      }

      // Group by date portion
      const byDate = {};
      for (const item of needsControl) {
        const baseId   = String(item.LabID || '').split(' ')[0].trim();
        const datePart = getDatePart(baseId);
        if (!datePart) continue;
        if (!byDate[datePart]) byDate[datePart] = new Set();
        byDate[datePart].add(baseId);
      }

      context.log(`[import-control] Dates: ${Object.keys(byDate).join(', ')}`);

      // Step 2: Parse matching files — look in month subfolder for each date
      const MONTHS_LIST = ['January','February','March','April','May','June',
                           'July','August','September','October','November','December'];
      const allRows   = [];
      const filesUsed = [];

      for (const [datePart, ids] of Object.entries(byDate)) {
        // Build month subfolder path from MMDDYY prefix
        const mm         = datePart.slice(0, 2);
        const yy         = datePart.slice(4, 6);
        const monthName  = MONTHS_LIST[parseInt(mm, 10) - 1] || '';
        const year       = '20' + yy;
        const monthFolder = `${folder}/${monthName} ${year}`;

        // List files in the month subfolder
        let monthFiles = [];
        try { monthFiles = await graphListFolder(monthFolder, token); } catch(e) {
          context.log(`[import-control] Month folder not found: ${monthFolder}`);
          continue;
        }
        const matchingFiles = monthFiles.filter(f => /\.xlsx?$/i.test(f.name) && f.name.includes(datePart));
        if (!matchingFiles.length) {
          context.log(`[import-control] No control file for date ${datePart} in ${monthFolder}`);
          continue;
        }
        for (const file of matchingFiles) {
          filesUsed.push(file.name);
          const buffer = await graphDownloadFile(file.id, token);
          const rows   = parseControlFile(buffer, ids);
          allRows.push(...rows);
          context.log(`[import-control] ${file.name}: ${rows.length} rows`);
        }
      }

      if (debug) {
        return { status: 200, jsonBody: { filesUsed, rowCount: allRows.length, rows: allRows.slice(0, 5) } };
      }

      // Step 4: Merge rows by baseId and write
      const byBase = {};
      for (const row of allRows) {
        if (!byBase[row.baseId]) byBase[row.baseId] = {};
        Object.assign(byBase[row.baseId], buildFields(row));
      }

      const log = []; let updated = 0, created = 0, errors = 0;

      for (const [baseId, fields] of Object.entries(byBase)) {
        if (!Object.keys(fields).length) continue;
        const existing = cacheItems.find(r => {
          const storedBase = String(r.LabID || '').split(' ')[0].trim();
          return storedBase === baseId || r.LabID === baseId;
        });
        if (existing) {
          await updateItem('Results Cache', existing._id, fields)
            .then(() => { updated++; log.push(`Updated: ${baseId}`); })
            .catch(e => { errors++; log.push(`Error ${baseId}: ${e.message}`); });
        } else {
          await createItem('Results Cache', { LabID: baseId, ...fields })
            .then(() => { created++; log.push(`Created: ${baseId}`); })
            .catch(e => { errors++; log.push(`Error ${baseId}: ${e.message}`); });
        }
      }

      return { status: 200, jsonBody: { success: true, filesUsed, rowCount: allRows.length, created, updated, errors, log, samples: Object.keys(byBase) } };

    } catch(e) {
      context.log('[import-control] Error:', e.message, e.stack);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
