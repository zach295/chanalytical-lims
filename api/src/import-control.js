const { app }    = require('@azure/functions');
const { listFolder, downloadFile, listItems, createItem, updateItem } = require('../shared/graph');

let XLSX;
try { XLSX = require('xlsx'); } catch(e) { console.warn('[import-control] xlsx not available:', e.message); }

const COL = {
  BARCODE:0, CHLORINE:1, DT_CHLORINE:2, PH:3, DT_PH:4,
  COLIFORM:5, ECOLI:6, START_INIT:7, START_DT:8, END_INIT:9, END_DT:10,
  LOT_COLLECT:11, LOT_QUAN:12,
  CHLORIDE:13, DT_CHLORIDE:14, FLUORIDE:15, DT_FLUORIDE:16,
  NITRITE:17, DT_NITRITE:18, NITRATE:19, DT_NITRATE:20,
  ALKALINITY:21, DT_ALKALINITY:22, SULFATE:23, DT_SULFATE:24,
  TANNINS:25, DT_TANNINS:26, TDS:27, DT_TDS:28, BROMIDE:29, DT_BROMIDE:30,
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

// Sanitize numeric values — negative = 0, non-numeric (like "<1") kept as-is
function sn(val) {
  if (!val || val === '') return val;
  const n = parseFloat(val);
  if (isNaN(n)) return val;
  return n < 0 ? '0' : String(Math.round(n * 10000) / 10000);
}

function parseControlSheet(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error('No sheets found');
  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows = [];
  for (let r = 1; r <= range.e.r; r++) {
    const barcode = cellVal(ws, r, COL.BARCODE);
    if (!barcode) continue;
    const baseId = getBaseId(barcode);
    if (!baseId) continue;
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
      const { debug, fileId: specificFileId } = body;

      const rawFolder = process.env.SP_CONTROL_FOLDER ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
      const marker = 'Shared Documents/';
      const mi = rawFolder.indexOf(marker);
      const folder = mi >= 0 ? rawFolder.slice(mi + marker.length) : rawFolder.replace(/^\/+/, '');

      let fileId = specificFileId, fileName = '';
      if (!fileId) {
        const files = await listFolder(folder);
        const xlsx  = files.filter(f => /\.xlsx?$/i.test(f.name));
        if (!xlsx.length) return { status: 404, body: JSON.stringify({ error: 'No Excel files in control folder' }) };
        const latest = xlsx[xlsx.length - 1];
        fileId = latest.id; fileName = latest.name;
        context.log(`[import-control] Using: ${fileName}`);
      }

      const buffer = await downloadFile(fileId);
      const rows   = parseControlSheet(buffer);

      if (debug) return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ fileName, rowCount: rows.length, rows: rows.slice(0, 5) }),
      };

      // Group rows by baseId and merge fields
      const byBase = {};
      for (const row of rows) {
        if (!byBase[row.baseId]) byBase[row.baseId] = {};
        Object.assign(byBase[row.baseId], buildFields(row));
      }

      const cacheItems = await listItems('Results Cache', { top: 500 });
      const log = []; let created = 0, updated = 0, errors = 0;

      for (const [baseId, fields] of Object.entries(byBase)) {
        if (!Object.keys(fields).length) continue;
        // Match on base ID — strip any suffix from stored LabID
        const existing = cacheItems.find(r => {
          const storedId = String(r.LabID || r.Title || '').trim();
          const storedBase = storedId.split(' ')[0].trim();
          return storedBase === baseId || storedId === baseId;
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

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, fileName, rowCount: rows.length, created, updated, errors, log }),
      };
    } catch(e) {
      context.log('[import-control] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
