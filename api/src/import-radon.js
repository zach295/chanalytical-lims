/**
 * import-radon.js
 * POST /api/import-radon
 * Reads RCS_MMDDYY.xlsx files and imports radon results into:
 *   - Results Cache (Radon, RadonDate, RadonTime columns)
 *   - Reports to be Billed list (RW Results column)
 */
const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');

const GRAPH  = 'https://graph.microsoft.com/v1.0';
const MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];

function getMonthFolder(datePrefix) {
  const mm   = datePrefix.slice(0, 2);
  const yy   = datePrefix.slice(4, 6);
  return `${MONTHS[parseInt(mm, 10) - 1]} Radon 20${yy}`;
}

async function graphGet(url, token) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`GET ${url.slice(-60)} → ${res.status}`);
  return res.json();
}

async function graphPatch(url, body, token) {
  const res = await fetch(url, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  return res;
}

async function getListId(siteId, displayName, token) {
  const data = await graphGet(
    `${GRAPH}/sites/${siteId}/lists?$select=id,displayName&$top=50`, token);
  return (data.value || []).find(l => l.displayName === displayName)?.id || null;
}

async function readWorksheetValues(siteId, fileId, token) {
  const wbBase = `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook`;
  // Open session
  const sesRes = await fetch(`${wbBase}/createSession`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ persistChanges: false }),
  });
  const { id: sid } = await sesRes.json();
  const wbHdr = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid };
  try {
    const sheetsRes = await fetch(`${wbBase}/worksheets`, { headers: wbHdr });
    const wsId = ((await sheetsRes.json()).value || [])[0]?.id;
    if (!wsId) return [];
    const rangeRes = await fetch(
      `${wbBase}/worksheets/${wsId}/usedRange?$select=values`,
      { headers: wbHdr }
    );
    return (await rangeRes.json()).values || [];
  } finally {
    await fetch(`${wbBase}/closeSession`, {
      method: 'POST', headers: { ...wbHdr, 'Content-Type': 'application/json' }
    }).catch(() => {});
  }
}

app.http('import-radon', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const token   = await getToken();
      const siteId  = process.env.SP_SITE_ID;
      const authHdr = { Authorization: `Bearer ${token}` };

      // ── 1. Load Results Cache — find rows with empty Radon field ──────────────
      const rcListId = await getListId(siteId, 'Results Cache', token);
      if (!rcListId) throw new Error('Results Cache list not found');

      const rcRes  = await fetch(
        `${GRAPH}/sites/${siteId}/lists/${rcListId}/items?$expand=fields&$top=2000`,
        { headers: authHdr }
      );
      const rcItems = ((await rcRes.json()).value || [])
        .filter(i => i.fields?.LabID && !i.fields?.Radon);

      if (!rcItems.length) {
        return { status: 200, jsonBody: { success: true, message: 'No pending radon samples found', updated: 0 } };
      }

      // ── 2. Group by MMDDYY prefix ─────────────────────────────────────────────
      const byPrefix = {};
      for (const item of rcItems) {
        const labId  = String(item.fields.LabID || '').trim();
        const prefix = labId.slice(0, 6); // MMDDYY
        if (!prefix.match(/^\d{6}$/)) continue;
        if (!byPrefix[prefix]) byPrefix[prefix] = [];
        byPrefix[prefix].push(item);
      }

      // ── 3. Get RTB list ID for RW Results update ──────────────────────────────
      const rtbListId = await getListId(siteId, 'Reports to be Billed', token);

      // ── 4. Load control folder path ───────────────────────────────────────────
      const controlFolder = process.env.SP_CONTROL_FOLDER ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
      const marker   = 'Shared Documents/';
      const mIdx     = controlFolder.indexOf(marker);
      const relPath  = mIdx >= 0
        ? controlFolder.slice(mIdx + marker.length)
        : controlFolder.replace(/^\/+/, '');

      let totalUpdated = 0;
      const log = [];

      // ── 5. Process each MMDDYY prefix ────────────────────────────────────────
      for (const [prefix, cacheItems] of Object.entries(byPrefix)) {
        const monthFolder = getMonthFolder(prefix);
        const rcsName     = `RCS_${prefix}.xlsx`;
        const rcsPath     = `${relPath}/${monthFolder}/${rcsName}`;
        const encPath     = rcsPath.split('/').map(encodeURIComponent).join('/');

        // Find the RCS file
        const fileRes = await fetch(
          `${GRAPH}/sites/${siteId}/drive/root:/${encPath}?$select=id`,
          { headers: authHdr }
        ).catch(() => null);

        if (!fileRes || !fileRes.ok) {
          log.push(`⚠️ ${rcsName} not found in ${monthFolder}`);
          continue;
        }
        const { id: fileId } = await fileRes.json();

        // Read all rows from the RCS
        const rows = await readWorksheetValues(siteId, fileId, token);
        if (!rows.length) { log.push(`⚠️ ${rcsName} is empty`); continue; }

        // Find column indices from header row
        const headers = rows[0].map(h => String(h || '').toLowerCase().trim());
        const colA    = 0; // Lab Barcode # always col A
        const colResult = headers.findIndex(h => h.includes('result') && h.includes('pci'));
        const colDateTested = headers.findIndex(h => h.includes('date') && h.includes('test'));
        const colTimeTested = headers.findIndex(h => h.includes('time') && h.includes('test'));

        if (colResult < 0) {
          log.push(`⚠️ ${rcsName}: could not find Results pCi/L column. Headers: ${headers.join(', ')}`);
          continue;
        }

        // Build lookup: base lab ID → { result, dateTested, timeTested }
        const rcsData = {};
        for (let i = 1; i < rows.length; i++) {
          const row     = rows[i];
          const barcode = String(row[colA] || '').trim();
          if (!barcode) continue;
          const baseId  = barcode.split(' ')[0].trim();
          const result  = row[colResult] != null ? String(row[colResult]).trim() : '';
          const dateTested = colDateTested >= 0 ? String(row[colDateTested] || '').trim() : '';
          const timeTested = colTimeTested >= 0 ? String(row[colTimeTested] || '').trim() : '';
          if (result) rcsData[baseId] = { result, dateTested, timeTested, fullBarcode: barcode };
        }

        context.log(`[Radon] ${rcsName}: found ${Object.keys(rcsData).length} results`);

        // ── 6. Update Results Cache and RTB for each matching cache item ──────
        for (const cacheItem of cacheItems) {
          const labId  = String(cacheItem.fields.LabID || '').trim();
          const rcsRow = rcsData[labId];
          if (!rcsRow || !rcsRow.result) {
            log.push(`ℹ️ ${labId}: no result in ${rcsName}`);
            continue;
          }

          // Update Results Cache
          const rcPatch = await fetch(
            `${GRAPH}/sites/${siteId}/lists/${rcListId}/items/${cacheItem.id}/fields`,
            { method: 'PATCH', headers: { ...authHdr, 'Content-Type': 'application/json' },
              body: JSON.stringify({
                Radon:     rcsRow.result,
                RadonDate: rcsRow.dateTested,
                RadonTime: rcsRow.timeTested,
              }) }
          );
          if (!rcPatch.ok) {
            log.push(`⚠️ ${labId}: Results Cache update failed (${rcPatch.status})`);
            continue;
          }

          // Update Reports to be Billed — find matching row by Title (Lab #)
          if (rtbListId) {
            const rtbSearch = await fetch(
              `${GRAPH}/sites/${siteId}/lists/${rtbListId}/items?$expand=fields($select=id,Title)&$top=2000`,
              { headers: authHdr }
            );
            if (rtbSearch.ok) {
              const rtbItems = (await rtbSearch.json()).value || [];
              const rtbMatch = rtbItems.find(i => String(i.fields?.Title || '').trim() === labId);
              if (rtbMatch) {
                // Get internal name for RW Results using colMap
                const colsRes  = await fetch(
                  `${GRAPH}/sites/${siteId}/lists/${rtbListId}/columns?$select=name,displayName`,
                  { headers: authHdr }
                );
                let rwField = 'RW_x0020_Results';
                if (colsRes.ok) {
                  const cols = (await colsRes.json()).value || [];
                  const col  = cols.find(c => c.displayName === 'RW Results');
                  if (col) rwField = col.name;
                }
                await fetch(
                  `${GRAPH}/sites/${siteId}/lists/${rtbListId}/items/${rtbMatch.id}/fields`,
                  { method: 'PATCH', headers: { ...authHdr, 'Content-Type': 'application/json' },
                    body: JSON.stringify({ [rwField]: rcsRow.result }) }
                );
              }
            }
          }

          totalUpdated++;
          log.push(`✅ ${labId}: ${rcsRow.result} pCi/L`);
        }
      }

      return {
        status: 200,
        jsonBody: { success: true, updated: totalUpdated, log },
      };

    } catch(e) {
      context.log('[import-radon] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  },
});
