/**
 * import-control.js — Azure version
 * 1. Reads Results Cache to get Lab IDs needing control sheet data
 * 2. Groups IDs by date (MMDDYY from base ID)
 * 3. Finds the matching control sheet file for each date
 * 4. Writes pH, bacteria, Gallery chemistry to Results Cache
 */
const { app }    = require('@azure/functions');
const { listFolder, downloadFile, listItems, createItem, updateItem } = require('../shared/graph');

let XLSX;
try { XLSX = require('xlsx'); } catch(e) { console.warn('[import-control] xlsx not available:', e.message); }

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
  if (row.dt_ph)         f.field_1  = row.dt_ph;
  if (row.coliform)      f.field_2  = row.coliform;
  if (row.ecoli)         f.field_3  = row.ecoli;
  if (row.start_dt)      f.field_4  = row.start_dt;
  if (row.end_dt)        f.field_5  = row.end_dt;
  if (row.chloride)      f.field_6  = sn(row.chloride);
  if (row.dt_chloride)   f.field_7  = row.dt_chloride;
  if (row.fluoride)      f.field_8  = sn(row.fluoride);
  if (row.dt_fluoride)   f.field_9  = row.dt_fluoride;
  if (row.nitrite)       f.field_10 = sn(row.nitrite);
  if (row.dt_nitrite)    f.field_11 = row.dt_nitrite;
  if (row.nitrate)       f.field_12 = sn(row.nitrate);
  if (row.dt_nitrate)    f.field_13 = row.dt_nitrate;
  if (row.alkalinity)    f.field_14 = sn(row.alkalinity);
  if (row.dt_alkalinity) f.field_15 = row.dt_alkalinity;
  if (row.sulfate)       f.field_16 = sn(row.sulfate);
  if (row.dt_sulfate)    f.field_17 = row.dt_sulfate;
  if (row.tannins)       f.field_18 = sn(row.tannins);
  if (row.dt_tannins)    f.field_19 = row.dt_tannins;
  if (row.tds)           f.field_20 = sn(row.tds);
  if (row.dt_tds)        f.field_21 = row.dt_tds;
  if (row.bromide)       f.field_22 = sn(row.bromide);
  if (row.dt_bromide)    f.field_23 = row.dt_bromide;
  return f;
}

app.http('import-control', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (!XLSX) return { status: 500, body: JSON.stringify({ error: 'xlsx not installed' }) };
      const body = await request.json().catch(() => ({}));
      const { debug, all: importAll } = body;

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
      });

      if (!needsControl.length) {
        return { status: 200, headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: true, message: 'All Results Cache entries already have control data', updated: 0 }) };
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

      // Step 2: List all control files
      const allFiles   = await listFolder(folder);
      const xlsxFiles  = allFiles.filter(f => /\.xlsx?$/i.test(f.name));

      // Step 3: Parse matching files
      const allRows  = [];
      const filesUsed = [];

      for (const [datePart, ids] of Object.entries(byDate)) {
        const matchingFiles = xlsxFiles.filter(f => f.name.includes(datePart));
        if (!matchingFiles.length) {
          context.log(`[import-control] No file for date ${datePart}`);
          continue;
        }
        for (const file of matchingFiles) {
          filesUsed.push(file.name);
          const buffer = await downloadFile(file.id);
          const rows   = parseControlFile(buffer, ids);
          allRows.push(...rows);
          context.log(`[import-control] ${file.name}: ${rows.length} rows`);
        }
      }

      if (debug) {
        return { status: 200, headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ filesUsed, rowCount: allRows.length, rows: allRows.slice(0, 5) }) };
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

      return { status: 200, headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, filesUsed, rowCount: allRows.length, created, updated, errors, log }) };

    } catch(e) {
      context.log('[import-control] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
