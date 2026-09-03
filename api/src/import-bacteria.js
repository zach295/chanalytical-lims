/**
 * import-bacteria.js
 * POST /api/import-bacteria
 * Reads bac_MMDDYY.xlsm files from Control Sheets month folders.
 * Headers at row 5. Columns used:
 *   Col A (0): Sample ID  → match to Results Cache LabID
 *   Col E (4): Time In    → field_4 (Start Date/Time)
 *   Col G (6): Time Out   → field_5 (End Date/Time)
 *   Col K (10): Coliform MPN → field_2
 *   Col L (11): E. Coli MPN  → field_3
 */
const { app }     = require('@azure/functions');
const { getToken } = require('../shared/graph');
const { writeActivityLog } = require('../shared/audit');

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

async function readSheetValues(siteId, fileId, token, context) {
  const wbBase = `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook`;
  const sesRes = await fetch(`${wbBase}/createSession`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ persistChanges: false }),
  });
  if (!sesRes.ok) throw new Error('Session failed: ' + sesRes.status);
  const sid = (await sesRes.json()).id;
  const hdr = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid };

  // Get first sheet
  const wsRes = await fetch(`${wbBase}/worksheets`, { headers: hdr });
  if (!wsRes.ok) throw new Error('Worksheets failed');
  const sheets = (await wsRes.json()).value || [];
  if (!sheets.length) throw new Error('No sheets found');

  const wsId  = sheets[0].id;
  const urRes = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook/worksheets/${wsId}/usedRange?$select=values`, { headers: hdr });
  const rows  = urRes.ok ? ((await urRes.json()).values || []) : [];

  await fetch(`${wbBase}/closeSession`, { method: 'POST', headers: hdr }).catch(() => {});
  return rows;
}

app.http('import-bacteria', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {
    try {
      const body      = await request.json().catch(() => ({}));
      const dateFilter = body.datePrefix ? String(body.datePrefix).trim() : '';
      const debug   = body.debug === true;
      const siteId  = process.env.SP_SITE_ID;
      const token   = await getToken();
      const authHdr = { Authorization: `Bearer ${token}` };

      // ── 1. Load Results Cache items for selected date only ──────────────────
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
      if (!cacheItems.length) {
        return { status: 200, jsonBody: { success: true, updated: 0, log: [`No Results Cache items found for ${dateFilter}`] } };
      }
      const byPrefix = { [dateFilter]: cacheItems };

      // ── 2. Resolve control sheets folder ────────────────────────────────────
      const controlFolder = process.env.SP_CONTROL_FOLDER ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
      const marker  = 'Shared Documents/';
      const mIdx    = controlFolder.indexOf(marker);
      const relPath = mIdx >= 0
        ? controlFolder.slice(mIdx + marker.length)
        : controlFolder.replace(/^\/+/, '');

      let totalUpdated = 0;
      const log = [];
      const importedSamples = [];

      // ── 3. Process each date prefix ─────────────────────────────────────────
      for (const [prefix, items] of Object.entries(byPrefix)) {
        const monthFolder = getMonthFolder(prefix);
        const bacName     = `bac_${prefix}.xlsm`;
        const bacPath     = `${relPath}/${monthFolder}/${bacName}`;
        const encPath     = bacPath.split('/').map(encodeURIComponent).join('/');

        const fileRes = await fetch(
          `${GRAPH}/sites/${siteId}/drive/root:/${encPath}?$select=id`,
          { headers: authHdr }
        ).catch(() => null);

        if (!fileRes || !fileRes.ok) {
          log.push(`⚠️ ${bacName} not found in ${monthFolder}`);
          continue;
        }

        const { id: fileId } = await fileRes.json();
        const rows = await readSheetValues(siteId, fileId, token, context);

        // Headers are at row 5 (index 4) — data starts at row 6 (index 5)
        const HEADER_ROW  = 4;
        const DATA_START  = 5;
        const COL_SAMPLEID   = 0;
        const COL_TIME_IN    = 4;
        const COL_TIME_OUT   = 6;
        const COL_COLIFORM   = 10;
        const COL_ECOLI      = 11;

        // Build lookup from sample ID → result row
        const bacData = {};
        for (let r = DATA_START; r < rows.length; r++) {
          const row      = rows[r] || [];
          const sampleId = String(row[COL_SAMPLEID] || '').trim();
          if (!sampleId) continue;
          // Extract base lab ID (first MMDDYY-NNN)
          const baseMatch = sampleId.match(/^(\d{6}-\d{3})/);
          const baseId    = baseMatch ? baseMatch[1] : sampleId;
          const coliform  = row[COL_COLIFORM];
          const ecoli     = row[COL_ECOLI];
          const timeIn    = row[COL_TIME_IN];
          const timeOut   = row[COL_TIME_OUT];
          if (coliform == null && ecoli == null) continue;
          bacData[baseId] = { coliform, ecoli, timeIn, timeOut };
        }

        context.log(`[bac] ${bacName}: ${Object.keys(bacData).length} results`);

        // ── 4. Update Results Cache ────────────────────────────────────────────
        for (const item of items) {
          const labId  = String(item.fields.LabID || '').trim();
          const baseId = labId.replace(/\s+RW\s*$/i, '').trim();
          const result = bacData[baseId] || bacData[labId];
          if (!result) { log.push(`ℹ️ ${labId}: no bacteria result in ${bacName}`); continue; }

          const fields = {};
          if (result.coliform != null) fields.field_2 = String(result.coliform);
          if (result.ecoli    != null) fields.field_3 = String(result.ecoli);
          if (result.timeIn   != null) fields.field_4 = toMilitaryDT(result.timeIn);
          if (result.timeOut  != null) fields.field_5 = toMilitaryDT(result.timeOut);

          const patch = await fetch(
            `${GRAPH}/sites/${siteId}/lists/Results Cache/items/${item.id}/fields`,
            { method: 'PATCH', headers: { ...authHdr, 'Content-Type': 'application/json' },
              body: JSON.stringify(fields) }
          );
          if (patch.ok) {
            totalUpdated++;
            log.push(`✅ ${labId}: coliform=${result.coliform} ecoli=${result.ecoli}`);
            importedSamples.push({ labId, notes: `Total Coliform=${result.coliform} | E. Coli=${result.ecoli} | Start=${toMilitaryDT(result.timeIn)} | End=${toMilitaryDT(result.timeOut)} | Source: ${bacName}` });
          } else {
            log.push(`⚠️ ${labId}: update failed (${patch.status})`);
          }
        }
      }

      const actor = body.importedBy || body.updatedBy || 'Lab Staff';
      const auditWarnings = [];
      for (const sample of importedSamples) {
        const audit = await writeActivityLog({ labId: sample.labId, type: 'Results Imported - Bacteria', notes: sample.notes, by: actor, context });
        if (!audit.success) auditWarnings.push(`${sample.labId}: ${audit.error}`);
      }
      return { status: 200, jsonBody: { success: true, updated: totalUpdated, log, auditWarnings } };

    } catch(e) {
      context.log('[import-bacteria] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  },
});
