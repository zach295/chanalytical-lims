/**
 * import-icpms.js — Azure version
 * Reads the most recent ICP-MS Excel file from SP_ICPMS_FOLDER,
 * parses the Concentrations sheet, handles red (rejected) cells
 * and dilution fallback logic, then writes to Results Cache.
 *
 * POST {} — imports latest file
 * POST { debug: true } — returns parsed data without writing
 * POST { fileId: "..." } — imports specific file
 */
const { app }      = require('@azure/functions');
const { getToken, listFolder, downloadFile, listItems, createItem, updateItem } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

// Lazy-load xlsx to avoid startup crash if not installed
let XLSX;
try { XLSX = require('xlsx'); } catch(e) { console.warn('[import-icpms] xlsx not available:', e.message); }

// Maps ICP-MS column header → Results Cache internal field name
const ELEMENT_MAP = {
  'Na 23':  'Sodium_x0028_Na23_x0029_',
  'Mg 24':  'Magnesium_x0028_Mg24_x0029_',
  'Ca 43':  'Calcium_x0028_Ca43_x0029_',
  'Cr 52':  'Chromium_x0028_Cr52_x0029_',
  'Fe 54':  'Iron_x0028_Fe54_x0029_',
  'Mn 55':  'Manganese_x0028_Mn55_x0029_',
  'Co 59':  'Cobalt_x0028_Co59_x0029_',
  'Cu 63':  'Copper_x0028_Cu63_x0029_',
  'As 75':  'Arsenic_x0028_As75_x0029_',
  'Cd 111': 'Cadmium_x0028_Cd111_x0029_',
  'Sb 121': 'Antimony_x0028_Sb121_x0029_',
  'Pb 208': 'Lead_x0028_Pb208_x0029_',
  'U 238':  'Uranium_x0028_U238_x0029_',
};

// Skip internal standards and QC rows
const SKIP_ELEMENTS = new Set(['Sc 45', 'Ge 74', 'In 115', 'Bi 209', 'Sr 88']);
const QC_PREFIXES   = ['cal ', 'ccb', 'ccs', 'cqc', 'smsd', 'sms_', 'cvm', 'qcs', 'calibration', 'blank', 'ccv'];

function isQCRow(id) {
  const low = String(id || '').toLowerCase().trim();
  return !low || QC_PREFIXES.some(p => low.startsWith(p));
}

function isSampleRow(id) {
  // Must match MMDDYY-NNN or MMDDYY-NNN x5 etc.
  return /^\d{6}-\d{3}(\s+[xX]\d+)?$/.test(String(id || '').trim());
}

function getDilution(id) {
  const m = String(id || '').match(/[xX](\d+)\s*$/);
  return m ? parseInt(m[1]) : 1;
}

function getBaseId(id) {
  // Extract MMDDYY-NNN regardless of what comes after
  const m = String(id || '').match(/^(\d{6}-\d{3})/);
  return m ? m[1] : '';
}

// Check if a cell has a red/salmon background (rejected)
// SheetJS captures direct fill colors; conditional formatting colors may not appear
function isCellRed(cell) {
  if (!cell || !cell.s) return false;
  const fg = String(cell.s.fgColor?.rgb || '').toUpperCase();
  const bg = String(cell.s.bgColor?.rgb || '').toUpperCase();
  const pt = String(cell.s.patternType || '').toLowerCase();
  // Red/salmon/pink fill patterns used in ICP-MS software
  const redPatterns = [
    'FF0000','C0504D','FF5050','FF9999','FFC7CE',
    'FF4444','CC0000','FF3333','EA9999','FF8080',
    'FFB6B6','FFBFBF','FF6666','FF0066','E06666',
    'CC4125','FF7575','FFAAAA','FF4500','DC143C',
  ];
  const hasFill = pt && pt !== 'none' && pt !== '';
  const isRed = redPatterns.some(p => fg.includes(p) || bg.includes(p));
  return isRed || (hasFill && redPatterns.some(p => fg.startsWith(p.slice(0,4))));
}

function parseIcpms(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellStyles: true });

  // Find Concentrations sheet
  const sheetName = wb.SheetNames.find(n => /concentrat/i.test(n)) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Concentrations sheet not found. Available: ${wb.SheetNames.join(', ')}`);

  const range = XLSX.utils.decode_range(ws['!ref']);

  // Read headers (row 0)
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
    headers[c] = cell ? String(cell.v || '').trim().replace(/\s*\(.*?\)\s*/g, '').trim() : '';
  }

  // Map element keys to column indices
  const elementCols = {};
  for (const elemKey of Object.keys(ELEMENT_MAP)) {
    const idx = headers.findIndex(h => {
      const hn = h.replace(/\s/g, '');
      const ek = elemKey.replace(/\s/g, '');
      return hn === ek || h === elemKey;
    });
    if (idx >= 0) elementCols[elemKey] = idx;
  }

  // Find key columns
  const sampleIdCol = headers.findIndex(h => /sample.?id/i.test(h) || h.toLowerCase() === 'sample id');
  const acqTimeCol  = headers.findIndex(h => /acquisition/i.test(h));

  if (sampleIdCol < 0) throw new Error(`Sample ID column not found. Headers: ${headers.slice(0,10).join(', ')}`);

  const rows = [];
  for (let r = 1; r <= range.e.r; r++) {
    const idCell   = ws[XLSX.utils.encode_cell({ r, c: sampleIdCol })];
    const sampleId = idCell ? String(idCell.v || '').trim() : '';
    if (!sampleId || isQCRow(sampleId) || !isSampleRow(sampleId)) continue;

    const acqCell = ws[XLSX.utils.encode_cell({ r, c: acqTimeCol })];
    const acqTime = acqCell ? String(acqCell.w || acqCell.v || '') : '';

    const elements = {};
    for (const [elemKey, colIdx] of Object.entries(elementCols)) {
      const cell     = ws[XLSX.utils.encode_cell({ r, c: colIdx })];
      const value    = cell && cell.v !== undefined && cell.v !== null ? cell.v : null;
      const rejected = isCellRed(cell);
      elements[elemKey] = { value, rejected };
    }

    rows.push({
      sampleId,
      baseId:   getBaseId(sampleId),
      dilution: getDilution(sampleId),
      acqTime,
      elements,
    });
  }

  return { rows, sheetName };
}

// Merge dilutions — prefer non-diluted, fall back for rejected cells
function mergeResults(rows) {
  const byBase = {};
  for (const row of rows) {
    if (!byBase[row.baseId]) byBase[row.baseId] = [];
    byBase[row.baseId].push(row);
  }

  const merged = {};
  for (const [baseId, baseRows] of Object.entries(byBase)) {
    baseRows.sort((a, b) => a.dilution - b.dilution); // 1, 2, 5, 10...

    const result = { baseId, acqTime: baseRows[0]?.acqTime || '', elements: {} };

    for (const elemKey of Object.keys(ELEMENT_MAP)) {
      for (const row of baseRows) {
        const el = row.elements[elemKey];
        if (!el) continue;
        if (!el.rejected && el.value !== null && el.value !== undefined) {
          result.elements[elemKey] = { value: el.value, dilution: row.dilution };
          break; // Got a good value — stop trying dilutions
        }
      }
    }

    merged[baseId] = result;
  }

  return merged;
}

app.http('import-icpms', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (!XLSX) return { status: 500, body: JSON.stringify({ error: 'xlsx package not installed on server. Check package.json dependencies.' }) };

      const body = await request.json().catch(() => ({}));
      const { debug, fileId: specificFileId } = body;

      const rawFolder = process.env.SP_ICPMS_FOLDER ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test ICPMS';
      // listFolder needs path relative to drive root (strip up to and including "Shared Documents/")
      const marker = 'Shared Documents/';
      const markerIdx = rawFolder.indexOf(marker);
      const icpmsFolder = markerIdx >= 0
        ? rawFolder.slice(markerIdx + marker.length)
        : rawFolder.replace(/^\/+/, '');

      // Find file
      let fileId = specificFileId;
      let fileName = '';
      if (!fileId) {
        const files = await listFolder(icpmsFolder);
        const xlsxFiles = files.filter(f => /\.xlsx?$/i.test(f.name));
        if (!xlsxFiles.length) return { status: 404, body: JSON.stringify({ error: 'No Excel files in ICPMS folder' }) };
        const latest = xlsxFiles[xlsxFiles.length - 1];
        fileId   = latest.id;
        fileName = latest.name;
        context.log(`[import-icpms] Using: ${fileName}`);
      }

      // Download and parse
      const buffer = await downloadFile(fileId);
      const { rows, sheetName } = parseIcpms(buffer);
      const merged = mergeResults(rows);

      if (debug) {
        // Also include raw cell style info for first sample row to diagnose red detection
        const wb2 = XLSX.read(buffer, { type: 'buffer', cellStyles: true });
        const ws2 = wb2.Sheets[sheetName];
        const range2 = XLSX.utils.decode_range(ws2['!ref']);
        const cellStyles = {};
        // Sample a few cells from row 19 (index 18) and row 24 (index 23) — the "rejected" rows
        for (let r = 1; r <= Math.min(30, range2.e.r); r++) {
          const idCell = ws2[XLSX.utils.encode_cell({ r, c: 0 })];
          const id = idCell ? String(idCell.v || '') : '';
          if (!id || !isSampleRow(id)) continue;
          const sampleCells = {};
          for (let c = 0; c <= Math.min(10, range2.e.c); c++) {
            const cell = ws2[XLSX.utils.encode_cell({ r, c })];
            if (cell && cell.s) {
              sampleCells[`col${c}`] = {
                v: cell.v,
                fgColor: cell.s.fgColor,
                bgColor: cell.s.bgColor,
                patternType: cell.s.patternType,
              };
            }
          }
          cellStyles[id] = sampleCells;
        }
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ fileName, sheetName, rowCount: rows.length, sampleCount: Object.keys(merged).length, merged, cellStyles }),
        };
      }

      // Write to Results Cache
      const cacheItems = await listItems('Results Cache', { top: 500 });
      const log = [];
      let created = 0, updated = 0, errors = 0;

      for (const [baseId, result] of Object.entries(merged)) {
        const fields = { AcquisitionTime: result.acqTime || '' };

        for (const [elemKey, elemResult] of Object.entries(result.elements)) {
          const fieldName = ELEMENT_MAP[elemKey];
          if (fieldName && elemResult) {
            // Format value: round to 4 decimal places for display
            const val = typeof elemResult.value === 'number'
              ? String(Math.round(elemResult.value * 10000) / 10000)
              : String(elemResult.value || '');
            fields[fieldName] = val;
          }
        }

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
        body: JSON.stringify({ success: true, fileName, sheetName, sampleCount: Object.keys(merged).length, created, updated, errors, log }),
      };

    } catch(e) {
      context.log('[import-icpms] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
