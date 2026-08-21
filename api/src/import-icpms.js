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

// Detect Arsenic Speciation ICP-MS suffixes
// TAs = Total Arsenic row, As3 = Arsenic III row
// Handles both "072326-014 TAs" and "072326-014TAs" formats
function getSpecSuffix(id) {
  const s = String(id || '').trim();
  if (/[\s-]?TAs$/i.test(s)) return 'TAs';
  if (/[\s-]?As3$/i.test(s)) return 'As3';
  return null;
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

    const specSuffix = getSpecSuffix(sampleId);
    if (specSuffix) console.log(`[icpms] Speciation row detected: "${sampleId}" → ${specSuffix}`);
    rows.push({ sampleId, baseId, specSuffix, dilution: getDilution(sampleId), acqTime, elements });
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
    // Separate As3 (Arsenic III speciation) rows from regular/TAs rows
    const as3Rows     = baseRows.filter(r => r.specSuffix === 'As3');
    const regularRows = baseRows.filter(r => r.specSuffix !== 'As3');
    regularRows.sort((a, b) => a.dilution - b.dilution);

    const result = { baseId, acqTime: regularRows[0]?.acqTime || baseRows[0]?.acqTime || '', elements: {}, arsenicIII: null };

    // Fill elements from regular/TAs rows only
    for (const elemKey of Object.keys(ELEMENT_MAP)) {
      for (const row of regularRows) {
        const el = row.elements[elemKey];
        if (!el) continue;
        if (!el.rejected && el.value !== null && el.value !== undefined) {
          result.elements[elemKey] = { value: el.value, dilution: row.dilution };
          break;
        }
      }
    }

    // Log what was found for this baseId
    if (as3Rows.length > 0 || regularRows.some(r => r.specSuffix === 'TAs')) {
      console.log(`[icpms] Speciation rows for ${baseId}: as3=${as3Rows.length} TAs=${regularRows.filter(r=>r.specSuffix==='TAs').length}`);
    }

    // Capture Arsenic III + its acquisition time from As3 row (As 75 column)
    if (as3Rows.length > 0) {
      for (const row of as3Rows.sort((a,b) => a.dilution - b.dilution)) {
        const as75 = row.elements['As 75'];
        if (as75 && !as75.rejected && as75.value !== null) {
          const num = typeof as75.value === 'number' ? as75.value : parseFloat(as75.value);
          if (!isNaN(num)) {
            result.arsenicIII        = num < 0 ? 0 : Math.round(num * 10000) / 10000;
            result.arsenicIIIAcqTime = row.acqTime || '';
            break;
          }
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
      // Always process all rows — find ones missing ANY element data
      // Results Cache "Lab ID" column may have internal name LabID OR Lab_x0020_ID
      const getLabId = r => String(r.LabID || r['Lab_x0020_ID'] || r['Lab ID'] || '').trim();
      const needsIcpms = cacheItems.filter(r => {
        const hasId = !!getLabId(r);
        if (!hasId) return false;
        // Skip only if ALL element fields are already populated
        const elementsFilled = Object.values({
          a: r.Sodium_x0028_Na23_x0029_,
          b: r.Arsenic_x0028_As75_x0029_,
          c: r.Uranium_x0028_U238_x0029_,
          d: r.Iron_x0028_Fe54_x0029_,
        }).every(v => !!(v || '').toString().trim());
        return importAll || !elementsFilled;
      });

      if (!needsIcpms.length) {
        return { status: 200, headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: true, message: 'All Results Cache entries already have ICP-MS data', updated: 0 }) };
      }

      // Group IDs by date portion (MMDDYY)
      const byDate = {};
      for (const item of needsIcpms) {
        const baseId   = getLabId(item).split(' ')[0].trim();
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
        const fields = { AcquisitionTime: toMilitaryDT(result.acqTime) };
        for (const [elemKey, elemResult] of Object.entries(result.elements)) {
          const fieldName = ELEMENT_MAP[elemKey];
          if (fieldName && elemResult) {
            const num = typeof elemResult.value === 'number' ? elemResult.value : parseFloat(elemResult.value);
            fields[fieldName] = isNaN(num) ? '' : num < 0 ? '0' : String(Math.round(num * 10000) / 10000);
          }
        }
        // ArsenicIII written separately to avoid blocking main update if field name is wrong

        const existing = cacheItems.find(r => {
          const storedBase = getLabId(r).split(' ')[0].trim();
          return storedBase === baseId || getLabId(r) === baseId;
        });

        if (existing) {
          await updateItem('Results Cache', existing._id, fields)
            .then(() => { updated++; log.push(`Updated: ${baseId}`); })
            .catch(e => { errors++; log.push(`Error ${baseId}: ${e.message}`); });
          // Separately write ArsenicIII — try multiple possible internal field names
          // so this doesn't block the main update if the name is wrong
          if (result.arsenicIII !== null && result.arsenicIII !== undefined) {
            const asIIIVal = String(result.arsenicIII);
            const asIIITime = result.arsenicIIIAcqTime || '';
            const fieldVariants = [
              { Arsenic3: asIIIVal, Arsenic3AcquisitionTime: asIIITime }, // SP internal names
              { ArsenicIII: asIIIVal }, // fallback attempts
              { ArsenicIII0: asIIIVal },
            ];
            for (const variant of fieldVariants) {
              try {
                await updateItem('Results Cache', existing._id, variant);
                log.push(`ArsenicIII written for ${baseId}: ${asIIIVal}`);
                break;
              } catch(e) {
                if (!e.message?.includes('not recognized') && !e.message?.includes('400')) throw e;
                // Try next variant
              }
            }
          }
        } else {
          await createItem('Results Cache', { LabID: baseId, ...fields })
            .then(() => { created++; log.push(`Created: ${baseId}`); })
            .catch(e => { errors++; log.push(`Error ${baseId}: ${e.message}`); });
        }
      }

      // ── Read Metals Prep (Acid Sheet) for MetalsStartDate_x002f_Time ──────────
      const acidDebug = { step: 'start', folder: '', fileFound: false, sheetFound: false, rows: 0, updated: 0 };
      try {
        const acidFolderRaw = process.env.SP_ACID_FOLDER || '';
        const acidMarker    = 'Shared Documents/';
        const acidMi        = acidFolderRaw.indexOf(acidMarker);
        const acidFolder    = acidMi >= 0 ? acidFolderRaw.slice(acidMi + acidMarker.length) : acidFolderRaw.replace(/^\/+/, '');
        acidDebug.folder    = acidFolder;
        acidDebug.step      = 'folder resolved';

        if (acidFolder) {
          const acidFiles = await listFolder(acidFolder);
          const acidFile  = acidFiles.find(f => /metals.?prep/i.test(f.name) && /\.xlsx?$/i.test(f.name));

          if (acidFile) {
            acidDebug.fileFound = true;
            acidDebug.fileName  = acidFile.name;
            const acidBuf   = await downloadFile(acidFile.id);
            const acidWb    = XLSX.read(acidBuf, { type: 'buffer', cellDates: true });
            acidDebug.sheets = acidWb.SheetNames;
            const monthAbbr = new Date().toLocaleString('en-US', { month: 'short', timeZone: 'America/New_York' });
            acidDebug.monthAbbr = monthAbbr;
            const sheetName = acidWb.SheetNames.find(s => s.toLowerCase().includes(monthAbbr.toLowerCase()));

            if (sheetName) {
              acidDebug.sheetFound = true;
              acidDebug.sheetName  = sheetName;
              const acidWs    = acidWb.Sheets[sheetName];
              const acidRange = XLSX.utils.decode_range(acidWs['!ref'] || 'A1:F1');
              const acidMap   = {};

              for (let r = 1; r <= acidRange.e.r; r++) {
                const datCell = acidWs[XLSX.utils.encode_cell({ r, c: 0 })]; // col A: date
                const timCell = acidWs[XLSX.utils.encode_cell({ r, c: 1 })]; // col B: time
                const idCell  = acidWs[XLSX.utils.encode_cell({ r, c: 5 })]; // col F: sample ID
                if (!idCell) continue;

                const baseId  = String(idCell.v || '').trim().match(/(\d{6}-\d{3})/)?.[1] || '';
                if (!baseId) continue;

                // Format date from col A
                let dateStr = '';
                if (datCell) {
                  if (datCell.t === 'd' && datCell.v instanceof Date) {
                    const d = datCell.v;
                    dateStr = `${d.getMonth()+1}/${d.getDate()}/${String(d.getFullYear()).slice(-2)}`;
                  } else { dateStr = String(datCell.w || datCell.v || '').trim(); }
                }

                // Format time from col B — convert to military if needed
                let timeStr = '';
                if (timCell) {
                  if (timCell.t === 'd' && timCell.v instanceof Date) {
                    const t = timCell.v;
                    timeStr = `${String(t.getHours()).padStart(2,'0')}:${String(t.getMinutes()).padStart(2,'0')}`;
                  } else {
                    const tw   = String(timCell.w || timCell.v || '').trim();
                    const ampm = tw.match(/(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?/i);
                    if (ampm) {
                      let h = parseInt(ampm[1], 10); const m = ampm[2];
                      if ((ampm[3]||'').toUpperCase() === 'PM' && h < 12) h += 12;
                      if ((ampm[3]||'').toUpperCase() === 'AM' && h === 12) h = 0;
                      timeStr = `${String(h).padStart(2,'0')}:${m}`;
                    } else { timeStr = tw; }
                  }
                }

                if (dateStr) acidMap[baseId] = timeStr ? `${dateStr} ${timeStr}` : dateStr;
              }

              // Write to Results Cache
              let acidUpdated = 0;
              for (const [baseId, startDT] of Object.entries(acidMap)) {
                const existing = cacheItems.find(r => String(r.LabID||'').split(' ')[0].trim() === baseId);
                if (existing) {
                  await updateItem('Results Cache', existing._id, { MetalsStartDate_x002f_Time: startDT })
                    .then(() => acidUpdated++)
                    .catch(() => {});
                }
              }
              context.log(`[import-icpms] Acid sheet: ${Object.keys(acidMap).length} rows, ${acidUpdated} MetalsStart updated`);
              acidDebug.rows    = Object.keys(acidMap).length;
              acidDebug.updated = acidUpdated;
              acidDebug.sampleIds = Object.keys(acidMap).slice(0, 5);
            } else {
              context.log(`[import-icpms] Acid sheet: no sheet matching "${monthAbbr}"`);
            }
          } else {
            context.log('[import-icpms] Acid sheet: metals prep file not found');
          }
        }
      } catch(acidErr) {
        context.log('[import-icpms] Acid sheet error:', acidErr.message);
        acidDebug.error = acidErr.message;
      }

      return { status: 200, headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, filesUsed, sampleCount: Object.keys(merged).length, created, updated, errors, log, acidDebug }) };

    } catch(e) {
      context.log('[import-icpms] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
