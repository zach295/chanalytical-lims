/**
 * import-acid.js — Azure version
 * 1. Reads Results Cache to get all Lab IDs missing MetalsStartDate
 * 2. Opens the acid sheet Excel file
 * 3. For each Lab ID, finds it in the correct month tab (derived from MMDDYY)
 * 4. Writes the date+time back to Results Cache
 */
const { app }    = require('@azure/functions');
const { listFolder, downloadFile, listItems, updateItem, createItem } = require('../shared/graph');
let XLSX;
try { XLSX = require('xlsx'); } catch(e) { console.warn('[import-acid] xlsx not available:', e.message); }

const MONTH_TABS = [
  /jan/i, /feb/i, /mar/i, /apr/i, /may/i, /jun/i,
  /jul/i, /aug/i, /sep/i, /oct/i, /nov/i, /dec/i,
];

function findMonthTab(wb, month0) {
  return wb.SheetNames.find(n => MONTH_TABS[month0].test(n) && /acid/i.test(n)) || null;
}

function monthFromBaseId(baseId) {
  const m = String(baseId || '').match(/^(\d{2})(\d{2})(\d{2})-/);
  if (!m) return null;
  return parseInt(m[1]) - 1;
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

function formatDateTime(dateStr, timeStr) {
  if (!dateStr) return '';
  try {
    let datePart = '';
    const dNum = parseFloat(dateStr);
    if (!isNaN(dNum) && dNum > 40000) {
      const d = new Date(Date.UTC(1899, 11, 30) + dNum * 86400000);
      datePart = `${String(d.getUTCMonth()+1).padStart(2,'0')}/${String(d.getUTCDate()).padStart(2,'0')}/${String(d.getUTCFullYear()).slice(-2)}`;
    } else {
      const parts = String(dateStr).split('/');
      if (parts.length >= 3) {
        const yr = parts[2].length === 4 ? parts[2].slice(-2) : parts[2];
        datePart = `${String(parts[0]).padStart(2,'0')}/${String(parts[1]).padStart(2,'0')}/${yr}`;
      } else datePart = dateStr;
    }
    let timePart = '';
    if (timeStr) {
      const tNum = parseFloat(timeStr);
      if (!isNaN(tNum) && tNum < 1) {
        const totalMin = Math.round(tNum * 1440);
        const h = Math.floor(totalMin / 60), mn = totalMin % 60;
        timePart = `${String(h).padStart(2,'0')}:${String(mn).padStart(2,'0')}`;
      } else {
        const am = String(timeStr).match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
        if (am) {
          let h = parseInt(am[1]); const mn = parseInt(am[2]);
          if (am[3]?.toUpperCase() === 'PM' && h < 12) h += 12;
          if (am[3]?.toUpperCase() === 'AM' && h === 12) h = 0;
          timePart = `${String(h).padStart(2,'0')}:${String(mn).padStart(2,'0')}`;
        }
      }
    }
    return timePart ? `${datePart} ${timePart}` : datePart;
  } catch { return `${dateStr} ${timeStr}`.trim(); }
}

function buildAcidLookup(wb, targetIds) {
  const lookup = {};
  const byMonth = {};
  for (const id of targetIds) {
    const m = monthFromBaseId(id);
    if (m === null) continue;
    if (!byMonth[m]) byMonth[m] = new Set();
    byMonth[m].add(id);
  }
  for (const [month0Str, ids] of Object.entries(byMonth)) {
    const month0  = parseInt(month0Str);
    const tabName = findMonthTab(wb, month0);
    if (!tabName) continue;
    const ws    = wb.Sheets[tabName];
    if (!ws) continue;
    const range = XLSX.utils.decode_range(ws['!ref']);
    let dataStart = 1;
    for (let r = 0; r <= Math.min(5, range.e.r); r++) {
      const f = cellStr(ws, r, 5);
      if (/sample.?id/i.test(f)) { dataStart = r + 1; break; }
    }
    for (let r = dataStart; r <= range.e.r; r++) {
      const sampleId = cellStr(ws, r, 5);
      if (!sampleId) continue;
      const baseId = getBaseId(sampleId);
      if (!baseId || !ids.has(baseId)) continue;
      if (lookup[baseId]) continue;
      lookup[baseId] = {
        formatted: formatDateTime(cellStr(ws, r, 0), cellStr(ws, r, 1)),
        tabName,
      };
    }
  }
  return lookup;
}

app.http('import-acid', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (!XLSX) return { status: 500, jsonBody: { error: 'xlsx not installed' } };
      const body       = await request.json().catch(() => ({}));
      const dateFilter = body.datePrefix ? String(body.datePrefix).trim() : '';
      const { debug, all: importAll } = body;

      if (!dateFilter) return { status: 400, jsonBody: { error: 'Select a date first' } };
      const allCache  = await listItems('Results Cache', { top: 2000 });
      const needsAcid = allCache.filter(r => {
        const labId = String(r.LabID || r['Lab_x0020_ID'] || r['Lab ID'] || '').trim();
        return labId && labId.startsWith(dateFilter) && !/\bREJ\b/i.test(labId);
      });

      if (!needsAcid.length) {
        return { status: 200, jsonBody: { success: true, message: 'All entries already have acid dates', updated: 0 } };
      }

      const targetIds = needsAcid.map(r => String(r.LabID || '').split(' ')[0].trim()).filter(Boolean);
      context.log(`[import-acid] ${targetIds.length} IDs need acid dates`);

      const rawFolder = process.env.SP_ACID_FOLDER || '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test M';
      const marker    = 'Shared Documents/';
      const mi        = rawFolder.indexOf(marker);
      const folder    = mi >= 0 ? rawFolder.slice(mi + marker.length) : rawFolder.replace(/^\/+/, '');

      const files      = await listFolder(folder);
      const xlsxFiles  = files.filter(f => /\.(xlsx?|xlsm|xlsb)$/i.test(f.name));
      if (!xlsxFiles.length) return { status: 404, jsonBody: { error: 'No Excel files in acid folder', folder, allFilesFound: files.map(f => f.name), hint: 'Check SP_ACID_FOLDER env var' } };

      const year     = String(new Date().getFullYear());
      const latest   = xlsxFiles.find(f => f.name.includes(year)) || xlsxFiles[xlsxFiles.length - 1];
      context.log(`[import-acid] Using: ${latest.name}`);

      const buffer = await downloadFile(latest.id);
      const wb     = XLSX.read(buffer, { type: 'buffer' });
      const lookup = buildAcidLookup(wb, targetIds);

      context.log(`[import-acid] Found dates for ${Object.keys(lookup).length}/${targetIds.length} IDs`);

      if (debug) {
        return { status: 200, jsonBody: { fileName: latest.name, targetCount: targetIds.length, foundCount: Object.keys(lookup).length, missing: targetIds.filter(id => !lookup[id]), lookup } };
      }

      const log = []; let updated = 0, errors = 0, notFound = 0;
      for (const cacheItem of needsAcid) {
        const baseId = String(cacheItem.LabID || '').split(' ')[0].trim();
        const entry  = lookup[baseId];
        if (!entry || !entry.formatted) { notFound++; continue; }
        await updateItem('Results Cache', cacheItem._id, { MetalsStartDate_x002f_Time: entry.formatted })
          .then(() => { updated++; log.push(`${baseId} → ${entry.formatted}`); })
          .catch(e => { errors++; log.push(`Error ${baseId}: ${e.message}`); });
      }

            // ── Activity Log ─────────────────────────────────────────────────────────
      try {
        const _now = new Date();
        const _ld  = _now.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
        const _lt  = _now.toLocaleTimeString('en-US', { timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
        await createItem('Activity Log', {
          Title: `${_ld} Acid Sheet Import`, Client: 'Import',
          ActivityType: 'Acid Sheet Import', Notes: `Updated: ${updated}, Not found: ${notFound}, Errors: ${errors}`,
          By: 'System', LogDate: _ld, LogTime: _lt, Quantity: 0,
        }).catch(()=>{});
      } catch(e) {}

      return { status: 200, jsonBody: { success: true, fileName: latest.name, updated, notFound, errors, log: log.slice(0, 30) } };

    } catch(e) {
      context.log('[import-acid] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
