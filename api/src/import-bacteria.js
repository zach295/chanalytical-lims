/**
 * import-bacteria.js
 * POST /api/import-bacteria
 * Reads bac_MMDDYY.xlsm files from Control Sheets month folders.
 * Identifies Sample ID, Time In, Time Out, Coliform MPN, and E. coli MPN
 * by worksheet headers so inserted columns cannot silently shift results.
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

        // Locate the header row and map columns by their labels rather than
        // fixed positions. The updated sheet adds analyst and incubation fields.
        const normalizeHeader = value => String(value ?? '').toLowerCase()
          .replace(/[^a-z0-9]+/g, ' ').trim();
        const headers = rows.slice(0, 15).map((row, index) => ({
          index, cells: (row || []).map(normalizeHeader),
        }));
        const header = headers.find(h =>
          h.cells.some(c => c === 'sample id') &&
          h.cells.some(c => c === 'time in') &&
          h.cells.some(c => c === 'time out') &&
          h.cells.some(c => c.includes('coliform') && c.includes('mpn')) &&
          h.cells.some(c => c.includes('e coli') && c.includes('mpn'))
        );
        if (!header) throw new Error(`Bacteria import: required column headers not found in ${bacName}; no results were imported from this file.`);
        const column = predicate => header.cells.findIndex(predicate);
        const DATA_START = header.index + 1;
        const COL_SAMPLEID = column(c => c === 'sample id');
        const COL_TIME_IN = column(c => c === 'time in');
        const COL_TIME_OUT = column(c => c === 'time out');
        const COL_COLIFORM = column(c => c.includes('coliform') && c.includes('mpn') && !c.includes('e coli'));
        const COL_ECOLI = column(c => c.includes('e coli') && c.includes('mpn'));
        if ([COL_SAMPLEID,COL_TIME_IN,COL_TIME_OUT,COL_COLIFORM,COL_ECOLI].some(c => c < 0))
          throw new Error(`Bacteria import: ambiguous or missing MPN columns in ${bacName}; import stopped.`);
        context.log('[bac] mapped columns', { sampleId:COL_SAMPLEID, timeIn:COL_TIME_IN, timeOut:COL_TIME_OUT, coliform:COL_COLIFORM, ecoli:COL_ECOLI });

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
          if ((coliform == null || String(coliform).trim() === '') && (ecoli == null || String(ecoli).trim() === '')) continue;
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
