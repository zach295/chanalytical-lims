/**
 * import-icpms.js — Azure version
 * 1. Reads Results Cache to get Lab IDs needing ICP-MS data
 * 2. Groups IDs by date (MMDDYY from base ID)
 * 3. Finds ALL matching ICP-MS files for each date (handles multiple runs)
 * 4. Merges results across files, writes to Results Cache
 */
const { app }    = require('@azure/functions');
const { listFolder, downloadFile, listItems, createItem, updateItem, getToken } = require('../shared/graph');

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

// Maps element key → SharePoint per-element acquisition time column
const ELEMENT_TIME_MAP = {
  'Na 23':  'AcqTime_Na',
  'Mg 24':  'AcqTime_Mg',
  'Ca 43':  'AcqTime_Ca',
  'Cr 52':  'AcqTime_Cr',
  'Fe 54':  'AcqTime_Fe',
  'Mn 55':  'AcqTime_Mn',
  'Co 59':  'AcqTime_Co',
  'Cu 63':  'AcqTime_Cu',
  'As 75':  'AcqTime_As',
  'Cd 111': 'AcqTime_Cd',
  'Sb 121': 'AcqTime_Sb',
  'Pb 208': 'AcqTime_Pb',
  'U 238':  'AcqTime_U',
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
  if (!cell) return false;
  const s = cell.s;
  if (!s) return false;

  // SheetJS stores fill colors at cell.s.fill.fgColor (not cell.s.fgColor directly)
  const fg = String(
    s.fill?.fgColor?.rgb || s.fill?.fgColor?.theme ||
    s.fgColor?.rgb || s.fgColor?.theme || ''
  ).toUpperCase();
  const bg = String(
    s.fill?.bgColor?.rgb || s.fill?.bgColor?.theme ||
    s.bgColor?.rgb || s.bgColor?.theme || ''
  ).toUpperCase();

  const combined = fg + '|' + bg;
  if (!combined.trim()) return false;

  // Red shades used by ICP-MS software — checked as substring so ARGB (8-char) works too
  const redPatterns = [
    'FF0000','C0504D','FA8072','FF5050','FF4444','FF9999',
    'FFC7CE','CC0000','FF3333','EA9999','FF8080','E06666',
    'FFB6B6','FFBFBF','FF6666','FF0066','CC4125','FF7575',
    'FFAAAA','FF4500','DC143C','FF3300','FF2222','FF1111',
  ];
  return redPatterns.some(p => combined.includes(p));
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

  const sampleIdCol = headers.findIndex(h =>
    /sample.?id/i.test(h) || /sample.?name/i.test(h) || h.toLowerCase() === 'sample id' ||
    h.toLowerCase() === 'sample name' || h.toLowerCase() === 'name' || h.toLowerCase() === 'id'
  );
  const acqTimeCol  = headers.findIndex(h => /acquisition/i.test(h) || /acq.*time/i.test(h) || /date.*time/i.test(h));
  if (sampleIdCol < 0) {
    console.log(`[icpms] No sample ID column found. Headers: ${headers.slice(0,15).join(', ')}`);
    return [];
  }
  console.log(`[icpms] Using column ${sampleIdCol} ("${headers[sampleIdCol]}") as sample ID, column ${acqTimeCol} as acqTime`);

  const rows = [];
  const rawIdsFound = []; // for diagnostics
  for (let r = 1; r <= range.e.r; r++) {
    const idCell   = ws[XLSX.utils.encode_cell({ r, c: sampleIdCol })];
    const sampleId = idCell ? String(idCell.v || '').trim() : '';
    if (!sampleId) continue;
    if (rawIdsFound.length < 10) rawIdsFound.push({ raw: sampleId, base: getBaseId(sampleId) });
    if (isQCRow(sampleId) || !isSampleRow(sampleId)) continue;
    const baseId = getBaseId(sampleId);
    if (!baseId) continue;
    // Only process IDs we need
        if (targetIds && targetIds.size > 0 && !targetIds.has(baseId)) continue;
        // Log first few IDs found so we can see if they're matching
        if (rows.length < 3) console.log(`[icpms] Found sample row: "${sampleId}" → baseId="${baseId}"`);

    const acqCell = ws[XLSX.utils.encode_cell({ r, c: acqTimeCol })];
    const acqTime = acqCell ? String(acqCell.w || acqCell.v || '') : '';

    const elements = {};
    for (const [elemKey, colIdx] of Object.entries(elementCols)) {
      const cell     = ws[XLSX.utils.encode_cell({ r, c: colIdx })];
      const value    = cell && cell.v !== undefined && cell.v !== null ? cell.v : null;
      const rejected = isCellRed(cell);
      // Log rejected cells so we can verify detection is working
      if (rejected) console.log(`[icpms] REJECTED cell (${elemKey}) row ${r+1}: value=${value} fg=${JSON.stringify(cell?.s?.fill?.fgColor||cell?.s?.fgColor)}`);
      elements[elemKey] = { value, rejected };
    }

    const specSuffix = getSpecSuffix(sampleId);
    if (specSuffix) console.log(`[icpms] Speciation row detected: "${sampleId}" → ${specSuffix}`);
    rows.push({ sampleId, baseId, specSuffix, dilution: getDilution(sampleId), acqTime, elements });
  }

  rows._rawIds = rawIdsFound;
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

    // Fill elements from regular/TAs rows only — track per-element acqTime
    for (const elemKey of Object.keys(ELEMENT_MAP)) {
      for (const row of regularRows) {
        const el = row.elements[elemKey];
        if (!el) continue;
        if (!el.rejected && el.value !== null && el.value !== undefined) {
          result.elements[elemKey] = { value: el.value, dilution: row.dilution, acqTime: row.acqTime || '' };
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
      // Always process all IDs — overwrite existing data so corrections take effect
      const needsIcpms = cacheItems.filter(r => {
        if (/\bREJ\b/i.test(r.LabID || '')) return false;
        return !!getLabId(r);
      });

      if (!needsIcpms.length) {
        return { status: 200, headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: true, message: 'No Results Cache entries found', updated: 0 }) };
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

      // Step 2: For each date group, look in month subfolder then flat folder
      const MONTHS_LIST = ['January','February','March','April','May','June',
                           'July','August','September','October','November','December'];
      const allRows = [];
      const filesUsed = [];
      const diagInfo = {
        datesSearched: Object.keys(byDate),
        idsNeeded: Object.fromEntries(Object.entries(byDate).map(([d, s]) => [d, [...s]])),
        rowsFoundInFiles: 0,
        mergedSampleCount: 0,
        rawIdsFromFile: {},
      };

      for (const [datePart, ids] of Object.entries(byDate)) {
        // Build month subfolder: MMDDYY → "August 2026"
        const mm        = datePart.slice(0, 2);
        const yy        = datePart.slice(4, 6);
        const year      = `20${yy}`;
        const monthName = MONTHS_LIST[parseInt(mm, 10) - 1] || '';
        // Also try 3-letter abbreviation (Aug, Sep, etc.)
        const monthAbbr = monthName.slice(0, 3);

        // Try month subfolder first, then flat root folder
        let matchingFiles = [];
        for (const tryFolder of [
          `${folder}/${monthName} ${year}`,
          `${folder}/${monthAbbr} ${year}`,
          folder,
        ]) {
          try {
            const folderFiles = await listFolder(tryFolder);
            const found = folderFiles.filter(f =>
              /\.xlsx?$/i.test(f.name) && f.name.replace(/[_\-\s]/g,'').toLowerCase().includes(datePart.replace(/[_\-\s]/g,'').toLowerCase())
            );
            if (found.length) {
              matchingFiles = found;
              context.log(`[import-icpms] Found ${found.length} file(s) in ${tryFolder}`);
              break;
            }
          } catch(e) {
            context.log(`[import-icpms] Folder ${tryFolder} not accessible: ${e.message}`);
          }
        }

        if (!matchingFiles.length) {
          context.log(`[import-icpms] No files found for date ${datePart} in any subfolder`);
          continue;
        }

        for (const file of matchingFiles) {
          filesUsed.push(file.name);
          const buffer = await downloadFile(file.id);
          const rows   = parseIcpmsFile(buffer, ids);
          allRows.push(...rows);
          context.log(`[import-icpms] ${file.name}: ${rows.length} rows, raw IDs sample: ${(rows._rawIds||[]).join(', ')}`);
          diagInfo.rawIdsFromFile = diagInfo.rawIdsFromFile || {};
          diagInfo.rawIdsFromFile[file.name] = rows._rawIds || [];
        }
      }

      const merged = mergeResults(allRows);

      // Update diag with final counts
      diagInfo.filesUsed = filesUsed;
      diagInfo.rowsFoundInFiles = allRows.length;
      diagInfo.mergedSampleCount = Object.keys(merged).length;

      if (debug) {
        return { status: 200, headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ filesUsed, sampleCount: Object.keys(merged).length, merged, diag: diagInfo }) };
      }

      // Discover actual internal field names for AcqTime columns using schema endpoint
      let resolvedTimeMap = { ...ELEMENT_TIME_MAP };
      try {
        const siteId2 = process.env.SP_SITE_ID;
        const schemaToken = await getToken();
        const fieldsRes = await fetch(
          `${GRAPH}/sites/${siteId2}/lists/Results Cache/fields?$select=name,displayName,staticName&$top=200`,
          { headers: { Authorization: `Bearer ${schemaToken}` } }
        );
        if (fieldsRes.ok) {
          const fieldsData = await fieldsRes.json();
          const spFields = fieldsData.value || [];
          // Map displayName → internal name for AcqTime columns
          const displayToInternal = {};
          for (const f of spFields) {
            if (f.displayName) displayToInternal[f.displayName] = f.name;
          }
          diagInfo.timeFieldNames = spFields
            .filter(f => f.displayName?.toLowerCase().includes('acqtime'))
            .map(f => `${f.displayName}→${f.name}`);
          // Resolve each element's time field to its actual internal name
          for (const [elemKey, displayName] of Object.entries(ELEMENT_TIME_MAP)) {
            if (displayToInternal[displayName]) {
              resolvedTimeMap[elemKey] = displayToInternal[displayName];
            }
          }
          context.log('[import-icpms] Resolved time fields:', JSON.stringify(resolvedTimeMap));
        }
      } catch(e) { context.log('[import-icpms] Field discovery error:', e.message); }

      // Step 4: Write to Results Cache
      const log = []; let updated = 0, created = 0, errors = 0;

      for (const [baseId, result] of Object.entries(merged)) {
        // Element values only (these SP field names are confirmed working)
        const fields = {};
        const timeFields = {};
        for (const [elemKey, elemResult] of Object.entries(result.elements)) {
          const fieldName     = ELEMENT_MAP[elemKey];
          const timeFieldName = resolvedTimeMap[elemKey];
          if (fieldName && elemResult) {
            const num = typeof elemResult.value === 'number' ? elemResult.value : parseFloat(elemResult.value);
            fields[fieldName] = isNaN(num) ? '' : num < 0 ? '0' : String(Math.round(num * 10000) / 10000);
          }
          if (timeFieldName && elemResult) {
            timeFields[timeFieldName] = toMilitaryDT(elemResult.acqTime || result.acqTime);
          }
        }
        // ArsenicIII written separately to avoid blocking main update if field name is wrong

        const existing = cacheItems.find(r => {
          const storedBase = getLabId(r).split(' ')[0].trim();
          return storedBase === baseId || getLabId(r) === baseId;
        });

        if (existing) {
          // Write element values
          await updateItem('Results Cache', existing._id, fields)
            .then(() => { updated++; log.push(`Updated: ${baseId}`); })
            .catch(e => { errors++; log.push(`Error ${baseId}: ${e.message}`); });
          // Write per-element times separately — won't block the main update
          if (Object.keys(timeFields).length) {
            updateItem('Results Cache', existing._id, timeFields)
              .catch(e => log.push(`AcqTime write skipped for ${baseId}: ${e.message.slice(0,80)}`));
          }
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

            // ── Activity Log ─────────────────────────────────────────────────────────
      try {
        const _now = new Date();
        const _ld  = _now.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
        const _lt  = _now.toLocaleTimeString('en-US', { timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
        await createItem('Activity Log', {
          Title: `${_ld} ICP-MS Import`, Client: 'Import',
          ActivityType: 'ICP-MS Import', Notes: `Updated: ${updated}, Created: ${created}, Errors: ${errors} | Files: ${filesUsed.join(", ")}`,
          By: 'System', LogDate: _ld, LogTime: _lt, Quantity: 0,
        }).catch(()=>{});
      } catch(e) {}

      return { status: 200, headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, filesUsed, sampleCount: Object.keys(merged).length, created, updated, errors, log, diag: { ...diagInfo, errors } }) };

    } catch(e) {
      context.log('[import-icpms] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
