/**
 * import-acid.js — Azure version
 * Reads the acid prep sheet (one Excel per year, monthly tabs),
 * finds the last business day's entries, and writes
 * MetalsStartDate_x002f_Time to Results Cache for each sample.
 *
 * POST {} — imports from last business day
 * POST { debug: true } — returns parsed data without writing
 * POST { date: "2026-07-28" } — import specific date
 */
const { app }    = require('@azure/functions');
const { listFolder, downloadFile, listItems, updateItem, createItem } = require('../shared/graph');

let XLSX;
try { XLSX = require('xlsx'); } catch(e) { console.warn('[import-acid] xlsx not available:', e.message); }

// Tab name patterns for each month (handles typos in sheet names)
const MONTH_TAB_PATTERNS = [
  /jan/i, /feb/i, /mar/i, /apr/i, /may/i, /jun/i,
  /jul/i, /aug/i, /sep/i, /oct/i, /nov/i, /dec/i,
];

function getLastBusinessDay(fromDate) {
  const d = fromDate ? new Date(fromDate) : new Date();
  const dow = d.getDay(); // 0=Sun, 1=Mon ... 6=Sat
  let daysBack = 1;
  if (dow === 1) daysBack = 3; // Monday → Friday
  if (dow === 0) daysBack = 2; // Sunday → Friday
  const result = new Date(d);
  result.setDate(result.getDate() - daysBack);
  return result;
}

function findMonthTab(wb, month0) {
  // month0 is 0-based (0=Jan, 6=Jul, etc.)
  const pattern = MONTH_TAB_PATTERNS[month0];
  return wb.SheetNames.find(n => pattern.test(n)) || null;
}

function cellStr(ws, r, c) {
  const cell = ws[XLSX.utils.encode_cell({ r, c })];
  if (!cell || cell.v === undefined) return '';
  return String(cell.w || cell.v).trim();
}

function getBaseId(sampleId) {
  const m = String(sampleId || '').match(/^(\d{6}-\d{3})/);
  return m ? m[1] : '';
}

// Format date+time for Results Cache: "7/28/2026 4:01 PM" → "07/28/26 16:01"
function formatDateTime(dateStr, timeStr) {
  if (!dateStr) return '';
  try {
    // Parse date like "7/28/2026" or Excel serial number
    let datePart = '';
    const dNum = parseFloat(dateStr);
    if (!isNaN(dNum) && dNum > 40000) {
      // Excel serial date
      const d = new Date(Date.UTC(1899, 11, 30) + dNum * 86400000);
      datePart = `${String(d.getUTCMonth()+1).padStart(2,'0')}/${String(d.getUTCDate()).padStart(2,'0')}/${String(d.getUTCFullYear()).slice(-2)}`;
    } else {
      // String like "7/28/2026"
      const parts = String(dateStr).split('/');
      if (parts.length >= 3) {
        const yr = parts[2].length === 4 ? parts[2].slice(-2) : parts[2];
        datePart = `${String(parts[0]).padStart(2,'0')}/${String(parts[1]).padStart(2,'0')}/${yr}`;
      } else {
        datePart = dateStr;
      }
    }

    // Parse time like "4:01 PM"
    let timePart = '';
    if (timeStr) {
      const tNum = parseFloat(timeStr);
      if (!isNaN(tNum) && tNum < 1) {
        // Excel time fraction
        const totalMin = Math.round(tNum * 1440);
        const h = Math.floor(totalMin / 60);
        const m = totalMin % 60;
        timePart = `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
      } else {
        // String like "4:01 PM"
        const am = String(timeStr).match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
        if (am) {
          let h = parseInt(am[1]), m = parseInt(am[2]);
          const isPM = am[3] && am[3].toUpperCase() === 'PM';
          if (isPM && h < 12) h += 12;
          if (!isPM && h === 12) h = 0;
          timePart = `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
        }
      }
    }

    return timePart ? `${datePart} ${timePart}` : datePart;
  } catch(e) {
    return `${dateStr} ${timeStr}`.trim();
  }
}

// Check if a date cell matches a target date
function dateMatches(cellDateStr, targetDate) {
  if (!cellDateStr) return false;
  const target = `${targetDate.getMonth()+1}/${targetDate.getDate()}/${targetDate.getFullYear()}`;

  // Handle Excel serial date
  const dNum = parseFloat(cellDateStr);
  if (!isNaN(dNum) && dNum > 40000) {
    const d = new Date(Date.UTC(1899, 11, 30) + dNum * 86400000);
    const cellTarget = `${d.getUTCMonth()+1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
    return cellTarget === target;
  }

  // Handle string date
  const parts = String(cellDateStr).split('/');
  if (parts.length >= 3) {
    const yr = parts[2].length === 2 ? `20${parts[2]}` : parts[2];
    const cellTarget = `${parseInt(parts[0])}/${parseInt(parts[1])}/${yr}`;
    return cellTarget === target;
  }
  return false;
}

function parseAcidSheet(buffer, targetDate) {
  const wb = XLSX.read(buffer, { type: 'buffer' });

  // Find the right month tab
  const month0 = targetDate.getMonth();
  const tabName = findMonthTab(wb, month0);
  if (!tabName) throw new Error(`No tab found for month ${month0+1}. Available tabs: ${wb.SheetNames.join(', ')}`);

  const ws = wb.Sheets[tabName];
  if (!ws) throw new Error(`Tab "${tabName}" not found`);

  const range = XLSX.utils.decode_range(ws['!ref']);
  const results = {}; // baseId → { dateStr, timeStr, formatted }

  // Find header row — look for "sample ID" or "sample id" in col F
  let dataStartRow = 1; // default skip row 0 (header)
  for (let r = 0; r <= Math.min(5, range.e.r); r++) {
    const cellF = cellStr(ws, r, 5); // col F (index 5)
    if (/sample.?id/i.test(cellF)) { dataStartRow = r + 1; break; }
  }

  for (let r = dataStartRow; r <= range.e.r; r++) {
    const dateStr   = cellStr(ws, r, 0); // A
    const timeStr   = cellStr(ws, r, 1); // B
    const sampleId  = cellStr(ws, r, 5); // F

    if (!sampleId || !dateStr) continue;
    if (!dateMatches(dateStr, targetDate)) continue;

    const baseId = getBaseId(sampleId);
    if (!baseId) continue;

    const formatted = formatDateTime(dateStr, timeStr);
    if (!results[baseId]) {
      results[baseId] = { dateStr, timeStr, formatted, tabName };
    }
  }

  return { results, tabName, targetDate: targetDate.toDateString() };
}

app.http('import-acid', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (!XLSX) return { status: 500, body: JSON.stringify({ error: 'xlsx not installed' }) };

      const body = await request.json().catch(() => ({}));
      const { debug, fileId: specificFileId, date: specificDate } = body;

      // Determine target date
      const targetDate = specificDate
        ? new Date(specificDate + 'T12:00:00Z')
        : getLastBusinessDay();
      context.log(`[import-acid] Target date: ${targetDate.toDateString()}`);

      // Find acid sheet file
      const rawFolder = process.env.SP_ACID_FOLDER ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test M';
      const marker = 'Shared Documents/';
      const mi = rawFolder.indexOf(marker);
      const folder = mi >= 0 ? rawFolder.slice(mi + marker.length) : rawFolder.replace(/^\/+/, '');

      let fileId = specificFileId, fileName = '';
      if (!fileId) {
        const files     = await listFolder(folder);
        const xlsxFiles = files.filter(f => /\.xlsx?$/i.test(f.name));
        if (!xlsxFiles.length) return { status: 404, body: JSON.stringify({ error: 'No Excel files in acid folder' }) };
        // Use the most recent file (or the one matching current year)
        const year = String(targetDate.getFullYear());
        const yearFile = xlsxFiles.find(f => f.name.includes(year));
        const latest   = yearFile || xlsxFiles[xlsxFiles.length - 1];
        fileId   = latest.id;
        fileName = latest.name;
        context.log(`[import-acid] Using: ${fileName}`);
      }

      // Download and parse
      const buffer = await downloadFile(fileId);
      const { results, tabName, targetDate: targetDateStr } = parseAcidSheet(buffer, targetDate);

      if (debug) {
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ fileName, tabName, targetDate: targetDateStr, sampleCount: Object.keys(results).length, results }),
        };
      }

      // Write MetalsStartDate_x002f_Time to Results Cache
      const cacheItems = await listItems('Results Cache', { top: 500 });
      const log = []; let updated = 0, created = 0, errors = 0;

      for (const [baseId, entry] of Object.entries(results)) {
        const fields = { MetalsStartDate_x002f_Time: entry.formatted };

        const existing = cacheItems.find(r => {
          const storedId   = String(r.LabID || '').trim();
          const storedBase = storedId.split(' ')[0].trim();
          return storedBase === baseId || storedId === baseId;
        });

        if (existing) {
          await updateItem('Results Cache', existing._id, fields)
            .then(() => { updated++; log.push(`Updated: ${baseId} → ${entry.formatted}`); })
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
        body: JSON.stringify({ success: true, fileName, tabName, targetDate: targetDateStr, sampleCount: Object.keys(results).length, updated, created, errors, log }),
      };

    } catch(e) {
      context.log('[import-acid] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
