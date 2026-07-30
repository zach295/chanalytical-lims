/**
 * import-icpms.js — Azure version
 * 1. Reads Results Cache to get Lab IDs needing ICP-MS data
 * 2. Groups IDs by date (MMDDYY from base ID)
 * 3. Finds ALL matching ICP-MS files for each date (handles multiple runs)
 * 4. Merges results across files, writes to Results Cache
 */
const { app }    = require('@azure/functions');
const { listFolder, downloadFile, listItems, createItem, updateItem } = require('../shared/graph');

let XLSX;
try { XLSX = require('xlsx'); } catch(e) { console.warn('[import-icpms] xlsx not available:', e.message); }

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

const QC_PREFIXES = ['cal ', 'ccb', 'ccs', 'cqc', 'smsd', 'sms_', 'cvm', 'qcs', 'calibration', 'blank', 'ccv'];

function isQCRow(id) {
  const low = String(id || '').toLowerCase().trim();
  return !low || QC_PREFIXES.some(p => low.startsWith(p));
}

function isSampleRow(id) {
  return /^\d{6}-\d{3}/.test(String(id || '').trim());
}

function getDilution(id) {
  const m = String(id || '').match(/[xX](\d+)/);
  return m ? parseInt(m[1]) : 1;
}

function getBaseId(id) {
  const m = String(id || '').match(/^(\d{6}-\d{3})/);
  return m ? m[1] : '';
}

function getDatePart(baseId) {
  const m = String(baseId || '').match(/^(\d{6})/);
  return m ? m[1] : '';
}

function isCellRed(cell) {
  if (!cell || !cell.s) return false;
  const fg = String(cell.s.fgColor?.rgb || '').toUpperCase();
  const bg = String(cell.s.bgColor?.rgb || '').toUpperCase();
  const redPatterns = [
    'FA8072','FF0000','C0504D','FF5050','FF9999','FFC7CE',
    'FF4444','CC0000','FF3333','EA9999','FF8080',
    'FFB6B6','FFBFBF','FF6666','FF0066','E06666',
    'CC4125','FF7575','FFAAAA','FF4500','DC143C',
  ];
  return redPatterns.some(p => fg.includes(p) || bg.includes(p));
}

function parseIcpmsFile(buffer, targetIds) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellStyles: true });
  const sheetName = wb.SheetNames.find(n => /concentrat/i.test(n)) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];

  const range = XLSX.utils.decode_range(ws['!ref']);
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
    headers[c] = cell ? String(cell.v || '').trim().replace(/\s*\(.*?\)\s*/g, '').trim() : '';
  }

  const elementCols = {};
  for (const elemKey of Object.keys(ELEMENT_MAP)) {
    const idx = headers.findIndex(h => {
      const hn = h.replace(/\s/g, '');
      const ek = elemKey.replace(/\s/g, '');
      return hn === ek || h === elemKey;
    });
    if (idx >= 0) elementCols[elemKey] = idx;
  }

  const sampleIdCol = headers.findIndex(h => /sample.?id/i.test(h) || h.toLowerCase() === 'sample id');
  const acqTimeCol  = headers.findIndex(h => /acquisition/i.test(h));
  if (sampleIdCol < 0) return [];

  const rows = [];
  for (let r = 1; r <= range.e.r; r++) {
    const idCell   = ws[XLSX.utils.encode_cell({ r, c: sampleIdCol })];
    const sampleId = idCell ? String(idCell.v || '').trim() : '';
    if (!sampleId || isQCRow(sampleId) || !isSampleRow(sampleId)) continue;
    const baseId = getBaseId(sampleId);
    if (!baseId) continue;
    // Only process IDs we need
    if (targetIds && targetIds.size > 0 && !targetIds.has(baseId)) continue;

    const acqCell = ws[XLSX.utils.encode_cell({ r, c: acqTimeCol })];
    const acqTime = acqCell ? String(acqCell.w || acqCell.v || '') : '';

    const elements = {};
    for (const [elemKey, colIdx] of Object.entries(elementCols)) {
      const cell     = ws[XLSX.utils.encode_cell({ r, c: colIdx })];
      const value    = cell && cell.v !== undefined && cell.v !== null ? cell.v : null;
      const rejected = isCellRed(cell);
      elements[elemKey] = { value, rejected };
    }

    rows.push({ sampleId, baseId, dilution: getDilution(sampleId), acqTime, elements });
  }

  return rows;
}

function mergeResults(rows) {
  const byBase = {};
  for (const row of rows) {
    if (!byBase[row.baseId]) byBase[row.baseId] = [];
    byBase[row.baseId].push(row);
  }

  const merged = {};
  for (const [baseId, baseRows] of Object.entries(byBase)) {
    baseRows.sort((a, b) => a.dilution - b.dilution);
    const result = { baseId, acqTime: baseRows[0]?.acqTime || '', elements: {} };
    for (const elemKey of Object.keys(ELEMENT_MAP)) {
      for (const row of baseRows) {
        const el = row.elements[elemKey];
        if (!el) continue;
        if (!el.rejected && el.value !== null && el.value !== undefined) {
          result.elements[elemKey] = { value: el.value, dilution: row.dilution };
          break;
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
      if (!XLSX) return { status: 500, body: JSON.stringify({ error: 'xlsx not installed' }) };
      const body = await request.json().catch(() => ({}));
      const { debug, all: importAll } = body;

      const rawFolder = process.env.SP_ICPMS_FOLDER ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test ICPMS';
      const marker = 'Shared Documents/';
      const mi     = rawFolder.indexOf(marker);
      const folder = mi >= 0 ? rawFolder.slice(mi + marker.length) : rawFolder.replace(/^\/+/, '');

      // Step 1: Get Results Cache — find IDs needing ICP-MS data
      const cacheItems = await listItems('Results Cache', { top: 500 });
      const needsIcpms = cacheItems.filter(r => {
        const hasId = !!(r.LabID || '').trim();
        const hasData = !!(r.AcquisitionTime || r.Sodium_x0028_Na23_x0029_ || '').trim();
        return hasId && (importAll || !hasData);
      });

      if (!needsIcpms.length) {
        return { status: 200, headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: true, message: 'All Results Cache entries already have ICP-MS data', updated: 0 }) };
      }

      // Group IDs by date portion (MMDDYY)
      const byDate = {};
      for (const item of needsIcpms) {
        const baseId   = String(item.LabID || '').split(' ')[0].trim();
        const datePart = getDatePart(baseId);
        if (!datePart) continue;
        if (!byDate[datePart]) byDate[datePart] = new Set();
        byDate[datePart].add(baseId);
      }

      context.log(`[import-icpms] Dates to process: ${Object.keys(byDate).join(', ')}`);

      // Step 2: List all files in ICPMS folder
      const allFiles = await listFolder(folder);
      const xlsxFiles = allFiles.filter(f => /\.xlsx?$/i.test(f.name));

      // Step 3: For each date group, find matching files and parse
      const allRows = [];
      const filesUsed = [];

      for (const [datePart, ids] of Object.entries(byDate)) {
        // Match files containing the date portion (e.g., M_072826-01.xlsx, M_072826-02.xlsx)
        const matchingFiles = xlsxFiles.filter(f => f.name.includes(datePart));
        if (!matchingFiles.length) {
          context.log(`[import-icpms] No files found for date ${datePart}`);
          continue;
        }
        context.log(`[import-icpms] Found ${matchingFiles.length} file(s) for ${datePart}: ${matchingFiles.map(f=>f.name).join(', ')}`);

        for (const file of matchingFiles) {
          filesUsed.push(file.name);
          const buffer = await downloadFile(file.id);
          const rows   = parseIcpmsFile(buffer, ids);
          allRows.push(...rows);
          context.log(`[import-icpms] ${file.name}: ${rows.length} rows`);
        }
      }

      const merged = mergeResults(allRows);

      if (debug) {
        return { status: 200, headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ filesUsed, sampleCount: Object.keys(merged).length, merged }) };
      }

      // Step 4: Write to Results Cache
      const log = []; let updated = 0, created = 0, errors = 0;

      for (const [baseId, result] of Object.entries(merged)) {
        const fields = { AcquisitionTime: result.acqTime || '' };
        for (const [elemKey, elemResult] of Object.entries(result.elements)) {
          const fieldName = ELEMENT_MAP[elemKey];
          if (fieldName && elemResult) {
            const num = typeof elemResult.value === 'number' ? elemResult.value : parseFloat(elemResult.value);
            fields[fieldName] = isNaN(num) ? '' : num < 0 ? '0' : String(Math.round(num * 10000) / 10000);
          }
        }

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
        body: JSON.stringify({ success: true, filesUsed, sampleCount: Object.keys(merged).length, created, updated, errors, log }) };

    } catch(e) {
      context.log('[import-icpms] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
