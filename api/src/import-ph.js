/**
 * import-ph.js
 * POST /api/import-ph
 * Reads pH_MMDDYY-NN.xlsx files from Control Sheets month folders.
 * Headers at row 10. Columns:
 *   Col 0: Analyst
 *   Col 1: Code (S1, S1 D, S2, S2 D, ...)
 *   Col 2: Sample (Lab ID)
 *   Col 3: pH result
 *   Col 4: Temp C
 *   Col 5: Date/Time
 *   Col 6: QC (Pass/Fail)
 *   Col 7: Comments
 *
 * Logic: for each sample, take pH + Date/Time from the S# row,
 * but only use it if the S# D (duplicate) row QC = "pass".
 * Results Cache: Title = pH, field_1 = Date/Time
 */
const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');

const GRAPH  = 'https://graph.microsoft.com/v1.0';
const MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];

function getMonthFolder(datePrefix) {
  const mm = datePrefix.slice(0, 2);
  const yy = datePrefix.slice(4, 6);
  return `${MONTHS[parseInt(mm, 10) - 1]} 20${yy}`;
}

function toMilitaryDT(val) {
  if (!val && val !== 0) return '';
  if (typeof val === 'number') {
    const ms = (val - 25569) * 86400 * 1000;
    const d  = new Date(Math.round(ms));
    return `${String(d.getUTCMonth()+1).padStart(2,'0')}/${String(d.getUTCDate()).padStart(2,'0')}/${String(d.getUTCFullYear()).slice(-2)} ${String(d.getUTCHours()).padStart(2,'0')}:${String(d.getUTCMinutes()).padStart(2,'0')}`;
  }
  const s    = String(val).trim();
  const ampm = s.match(/^(\d{1,2}\/\d{1,2}\/\d{2,4})\s+(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
  if (ampm) {
    let [, datePart, h, m, ap] = ampm;
    h = parseInt(h, 10); ap = (ap||'').toUpperCase();
    if (ap === 'PM' && h < 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
    datePart = datePart.replace(/(\d{1,2}\/\d{1,2}\/)(\d{4})/, (_, p1, y) => p1 + y.slice(-2));
    return `${datePart} ${String(h).padStart(2,'0')}:${m}`;
  }
  return s;
}

async function readSheetValues(siteId, fileId, token) {
  const wbBase = `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook`;
  const sesRes = await fetch(`${wbBase}/createSession`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ persistChanges: false }),
  });
  if (!sesRes.ok) throw new Error('Session failed: ' + sesRes.status);
  const sid = (await sesRes.json()).id;
  const hdr = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid };
  const wsRes = await fetch(`${wbBase}/worksheets`, { headers: hdr });
  if (!wsRes.ok) throw new Error('Worksheets failed');
  const sheets = (await wsRes.json()).value || [];
  if (!sheets.length) throw new Error('No sheets');
  const wsId  = sheets[0].id;
  const urRes = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook/worksheets/${wsId}/usedRange?$select=values`, { headers: hdr });
  const rows  = urRes.ok ? ((await urRes.json()).values || []) : [];
  await fetch(`${wbBase}/closeSession`, { method: 'POST', headers: hdr }).catch(() => {});
  return rows;
}

function parsePHFile(rows) {
  // Headers at row 10 (index 9), data starts at row 11 (index 10)
  const DATA_START = 10;
  const COL_ANALYST = 0;
  const COL_CODE    = 1; // S1, S1 D, S2, S2 D...
  const COL_SAMPLE  = 2; // Lab ID
  const COL_PH      = 3; // pH result
  const COL_TEMP    = 4; // Temp C (not imported)
  const COL_DT      = 5; // Date/Time
  const COL_QC      = 6; // QC (Pass/Fail)
  const COL_COMMENT = 7; // Comments (not imported)

  const sampleMap = {};

  for (let r = DATA_START; r < rows.length; r++) {
    const row    = rows[r] || [];
    const code   = String(row[COL_CODE] || '').trim();
    const sample = String(row[COL_SAMPLE] || '').trim();
    if (!code || !sample) continue;

    const baseMatch = sample.match(/^(\d{6}-\d{3})/);
    const labId     = baseMatch ? baseMatch[1] : sample;

    if (!sampleMap[labId]) sampleMap[labId] = { measurements: {}, duplicates: {} };

    const isDuplicate = / D$/i.test(code);
    const baseCode    = isDuplicate ? code.replace(/ D$/i, '').trim() : code;

    if (isDuplicate) {
      const qcRaw = row[COL_QC];
      const qcStr = String(qcRaw ?? '').toLowerCase().trim();
      const passed = qcStr === 'pass' || qcStr === 'true' || qcStr === '1' || qcRaw === true || qcRaw === 1;
      sampleMap[labId].duplicates[baseCode] = passed;
    } else {
      sampleMap[labId].measurements[baseCode] = {
        ph: row[COL_PH],
        dt: row[COL_DT],
      };
    }
  }

  const results = {};
  for (const [labId, data] of Object.entries(sampleMap)) {
    for (const [code, measurement] of Object.entries(data.measurements)) {
      const duplicatePassed = data.duplicates[code];
      if (duplicatePassed === true && measurement.ph != null) {
        if (!results[labId]) {
          results[labId] = { ph: measurement.ph, dt: measurement.dt };
        }
      }
    }
  }

  return results;
}

app.http('import-ph', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {
    try {
      const body = await request.json().catch(() => ({}));
      const dateFilter = String(body.datePrefix || body.dateFilter || '').trim();
      const siteId  = process.env.SP_SITE_ID;
      const token   = await getToken();
      const authHdr = { Authorization: `Bearer ${token}` };

      if (!dateFilter) return { status: 400, jsonBody: { error: 'Select a date first' } };
      let rcItems = [], rcNext = `${GRAPH}/sites/${siteId}/lists/Results Cache/items?$expand=fields($select=id,LabID)&$top=999`;
      while (rcNext) {
        const r    = await fetch(rcNext, { headers: authHdr });
        const data = await r.json();
        rcItems.push(...(data.value || []));
        rcNext = data['@odata.nextLink'] || null;
      }
      const cacheItems = rcItems.filter(i => {
        const id = String(i.fields?.LabID || '').trim();
        return id.startsWith(dateFilter) && !/\bREJ\b/i.test(id);
      });
      if (!cacheItems.length) return { status: 200, jsonBody: { success: true, updated: 0, log: [`No Results Cache items found for ${dateFilter}`] } };
      const byPrefix = { [dateFilter]: cacheItems };

      const controlFolder = process.env.SP_CONTROL_FOLDER ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
      const marker  = 'Shared Documents/';
      const mIdx    = controlFolder.indexOf(marker);
      const relPath = mIdx >= 0 ? controlFolder.slice(mIdx + marker.length) : controlFolder.replace(/^\/+/, '');

      let totalUpdated = 0;
      const log = [];

      for (const [prefix, items] of Object.entries(byPrefix)) {
        const monthFolder = getMonthFolder(prefix);
        const folderPath  = `${relPath}/${monthFolder}`;
        const encFolder   = folderPath.split('/').map(encodeURIComponent).join('/');

        const folderRes = await fetch(
          `${GRAPH}/sites/${siteId}/drive/root:/${encFolder}:/children?$select=id,name&$top=200`,
          { headers: authHdr }
        ).catch(() => null);

        if (!folderRes || !folderRes.ok) {
          log.push(`⚠️ Month folder not found: ${monthFolder}`);
          continue;
        }

        const folderData = await folderRes.json();
        const phFiles    = (folderData.value || []).filter(f =>
          new RegExp(`^pH_${prefix}(-\\d+)?\\.(xlsx?|xlsm)$`, 'i').test(f.name)
        );

        if (!phFiles.length) {
          log.push(`⚠️ No pH files found for ${prefix} in ${monthFolder}`);
          continue;
        }

        const mergedResults = {};
        for (const file of phFiles) {
          const rows = await readSheetValues(siteId, file.id, token);
          for (let i = 10; i < Math.min(15, rows.length); i++) {
            context.log(`[ph-debug] row ${i}:`, JSON.stringify(rows[i]?.slice(0, 8)));
          }
          const results = parsePHFile(rows);
          context.log(`[ph] ${file.name}: ${Object.keys(results).length} valid results`);
          Object.assign(mergedResults, results);
        }

        for (const item of items) {
          const labId  = String(item.fields.LabID || '').trim();
          const baseId = labId.replace(/\s+RW\s*$/i, '').trim();
          const result = mergedResults[baseId] || mergedResults[labId];
          if (!result) { log.push(`ℹ️ ${labId}: no passing pH result`); continue; }

          const patch = await fetch(
            `${GRAPH}/sites/${siteId}/lists/Results Cache/items/${item.id}/fields`,
            { method: 'PATCH', headers: { ...authHdr, 'Content-Type': 'application/json' },
              body: JSON.stringify({
                Title:   String(result.ph),
                field_1: toMilitaryDT(result.dt),
              }) }
          );
          if (patch.ok) {
            totalUpdated++;
            log.push(`✅ ${labId}: pH=${result.ph}`);
          } else {
            log.push(`⚠️ ${labId}: update failed (${patch.status})`);
          }
        }
      }

      return { status: 200, jsonBody: { success: true, updated: totalUpdated, log } };

    } catch(e) {
      context.log('[import-ph] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  },
});
