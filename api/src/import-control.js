/**
 * import-control.js — Azure version
 * Reads the most recent control sheet Excel from SP_CONTROL_FOLDER,
 * parses pH, bacteria, and Gallery chemistry results,
 * then writes to Results Cache.
 *
 * POST {} — imports latest control sheet
 * POST { debug: true } — returns parsed data without writing
 * POST { fileId: "..." } — imports specific file
 */
const { app }    = require('@azure/functions');
const { listFolder, downloadFile, listItems, createItem, updateItem } = require('../shared/graph');

// Lazy-load xlsx
let XLSX;
try { XLSX = require('xlsx'); } catch(e) { console.warn('[import-control] xlsx not available:', e.message); }

// Column index mapping (0-based, matches control sheet layout)
const COL = {
  BARCODE:     0,  // A — Lab ID (e.g. "072827-001 ALK, PH")
  CHLORINE:    1,  // B — Chlorine Abs/Pres
  DT_CHLORINE: 2,  // C — Date/Time
  PH:          3,  // D — pH value
  DT_PH:       4,  // E — Date/Time
  COLIFORM:    5,  // F — Coliform MPN
  ECOLI:       6,  // G — E.coli MPN
  START_INIT:  7,  // H — Starting Initials
  START_DT:    8,  // I — Start Date/Time (bacteria)
  END_INIT:    9,  // J — Ending Initials
  END_DT:      10, // K — End Date/Time (bacteria)
  LOT_COLLECT: 11, // L — Collect LOT #
  LOT_QUAN:    12, // M — Quan-Tray LOT #
  CHLORIDE:    13, // N — Chloride mg/L
  DT_CHLORIDE: 14, // O — Date/Time
  FLUORIDE:    15, // P — Fluoride mg/L
  DT_FLUORIDE: 16, // Q — Date/Time
  NITRITE:     17, // R — Nitrite mg/L
  DT_NITRITE:  18, // S — Date/Time
  NITRATE:     19, // T — Nitrate mg/L
  DT_NITRATE:  20, // U — Date/Time
  ALKALINITY:  21, // V — Alkalinity mg/L
  DT_ALKALINITY: 22, // W — Date/Time
  SULFATE:     23, // X — Sulfate mg/L
  DT_SULFATE:  24, // Y — Date/Time
  TANNINS:     25, // Z — Tannins mg/L
  DT_TANNINS:  26, // AA — Date/Time
  TDS:         27, // AB — Total Dissolved Solids
  DT_TDS:      28, // AC — Date/Time
  BROMIDE:     29, // AD — Bromide
  DT_BROMIDE:  30, // AE — Date/Time
};

function cellVal(ws, r, c) {
  const cell = ws[XLSX.utils.encode_cell({ r, c })];
  if (!cell || cell.v === undefined || cell.v === null) return '';
  // Use formatted string if available, else raw value
  const v = cell.w || String(cell.v);
  return v.trim() === 'N/A' || v.trim() === '#N/A' ? '' : v.trim();
}

function getBaseId(barcode) {
  const m = String(barcode || '').match(/^(\d{6}-\d{3})/);
  return m ? m[1] : '';
}

function parseControlSheet(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error(`No sheets found in control sheet file`);

  const range = XLSX.utils.decode_range(ws['!ref']);
  const rows = [];

  // Start from row 1 (skip header row 0)
  for (let r = 1; r <= range.e.r; r++) {
    const barcode = cellVal(ws, r, COL.BARCODE);
    if (!barcode) continue;

    const baseId = getBaseId(barcode);
    if (!baseId) continue; // Skip non-sample rows

    rows.push({
      baseId,
      barcode,
      ph:          cellVal(ws, r, COL.PH),
      dt_ph:       cellVal(ws, r, COL.DT_PH),
      coliform:    cellVal(ws, r, COL.COLIFORM),
      ecoli:       cellVal(ws, r, COL.ECOLI),
      start_dt:    cellVal(ws, r, COL.START_DT),
      end_dt:      cellVal(ws, r, COL.END_DT),
      chloride:    cellVal(ws, r, COL.CHLORIDE),
      dt_chloride: cellVal(ws, r, COL.DT_CHLORIDE),
      fluoride:    cellVal(ws, r, COL.FLUORIDE),
      dt_fluoride: cellVal(ws, r, COL.DT_FLUORIDE),
      nitrite:     cellVal(ws, r, COL.NITRITE),
      dt_nitrite:  cellVal(ws, r, COL.DT_NITRITE),
      nitrate:     cellVal(ws, r, COL.NITRATE),
      dt_nitrate:  cellVal(ws, r, COL.DT_NITRATE),
      alkalinity:  cellVal(ws, r, COL.ALKALINITY),
      dt_alkalinity: cellVal(ws, r, COL.DT_ALKALINITY),
      sulfate:     cellVal(ws, r, COL.SULFATE),
      dt_sulfate:  cellVal(ws, r, COL.DT_SULFATE),
      tannins:     cellVal(ws, r, COL.TANNINS),
      dt_tannins:  cellVal(ws, r, COL.DT_TANNINS),
      tds:         cellVal(ws, r, COL.TDS),
      dt_tds:      cellVal(ws, r, COL.DT_TDS),
      bromide:     cellVal(ws, r, COL.BROMIDE),
      dt_bromide:  cellVal(ws, r, COL.DT_BROMIDE),
    });
  }

  return rows;
}

// Build Results Cache fields from a parsed row
function buildCacheFields(row) {
  const fields = {};

  // PH → Title field (renamed to PH in Results Cache)
  if (row.ph) fields.Title = row.ph;

  // field_1 = Date/Time (PH date)
  if (row.dt_ph)        fields.field_1  = row.dt_ph;
  if (row.coliform)     fields.field_2  = row.coliform;
  if (row.ecoli)        fields.field_3  = row.ecoli;
  if (row.start_dt)     fields.field_4  = row.start_dt;
  if (row.end_dt)       fields.field_5  = row.end_dt;
  if (row.chloride)     fields.field_6  = row.chloride;
  if (row.dt_chloride)  fields.field_7  = row.dt_chloride;
  if (row.fluoride)     fields.field_8  = row.fluoride;
  if (row.dt_fluoride)  fields.field_9  = row.dt_fluoride;
  if (row.nitrite)      fields.field_10 = row.nitrite;
  if (row.dt_nitrite)   fields.field_11 = row.dt_nitrite;
  if (row.nitrate)      fields.field_12 = row.nitrate;
  if (row.dt_nitrate)   fields.field_13 = row.dt_nitrate;
  if (row.alkalinity)   fields.field_14 = row.alkalinity;
  if (row.dt_alkalinity)fields.field_15 = row.dt_alkalinity;
  if (row.sulfate)      fields.field_16 = row.sulfate;
  if (row.dt_sulfate)   fields.field_17 = row.dt_sulfate;
  if (row.tannins)      fields.field_18 = row.tannins;
  if (row.dt_tannins)   fields.field_19 = row.dt_tannins;
  if (row.tds)          fields.field_20 = row.tds;
  if (row.dt_tds)       fields.field_21 = row.dt_tds;
  if (row.bromide)      fields.field_22 = row.bromide;
  if (row.dt_bromide)   fields.field_23 = row.dt_bromide;

  return fields;
}

app.http('import-control', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (!XLSX) return { status: 500, body: JSON.stringify({ error: 'xlsx package not installed' }) };

      const body = await request.json().catch(() => ({}));
      const { debug, fileId: specificFileId } = body;

      const rawFolder = process.env.SP_CONTROL_FOLDER ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
      const marker    = 'Shared Documents/';
      const markerIdx = rawFolder.indexOf(marker);
      const folder    = markerIdx >= 0
        ? rawFolder.slice(markerIdx + marker.length)
        : rawFolder.replace(/^\/+/, '');

      // Find file
      let fileId = specificFileId;
      let fileName = '';
      if (!fileId) {
        const files     = await listFolder(folder);
        const xlsxFiles = files.filter(f => /\.xlsx?$/i.test(f.name));
        if (!xlsxFiles.length) return { status: 404, body: JSON.stringify({ error: 'No Excel files in control sheet folder' }) };
        const latest = xlsxFiles[xlsxFiles.length - 1];
        fileId   = latest.id;
        fileName = latest.name;
        context.log(`[import-control] Using: ${fileName}`);
      }

      // Download and parse
      const buffer = await downloadFile(fileId);
      const rows   = parseControlSheet(buffer);

      if (debug) {
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ fileName, rowCount: rows.length, rows: rows.slice(0, 5) }),
        };
      }

      // Write to Results Cache
      const cacheItems = await listItems('Results Cache', { top: 500 });
      const log = [];
      let created = 0, updated = 0, errors = 0;

      // Group rows by baseId — if multiple rows for same base ID, merge them
      const byBase = {};
      for (const row of rows) {
        if (!byBase[row.baseId]) byBase[row.baseId] = {};
        const fields = buildCacheFields(row);
        Object.assign(byBase[row.baseId], fields);
      }

      for (const [baseId, fields] of Object.entries(byBase)) {
        if (!Object.keys(fields).length) continue;

        const existing = cacheItems.find(r => (r.LabID || '').trim() === baseId);

        if (existing) {
          await updateItem('Results Cache', existing._id, fields)
            .then(() => { updated++; log.push(`Updated: ${baseId}`); })
            .catch(e => { errors++; log.push(`Error updating ${baseId}: ${e.message}`); });
        } else {
          await createItem('Results Cache', { LabID: baseId, ...fields })
            .then(() => { created++; log.push(`Created: ${baseId}`); })
            .catch(e => { errors++; log.push(`Error creating ${baseId}: ${e.message}`); });
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
