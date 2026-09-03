/**
 * import-radon.js
 * POST /api/import-radon
 * Reads RCS_MMDDYY.xlsx files and imports radon results into:
 *   - Results Cache (Radon, RadonDate, RadonTime columns)
 *   - Reports to be Billed list (RW Results column)
 */
const { app }      = require('@azure/functions');
const { getToken, createItem, listItems } = require('../shared/graph');

const GRAPH  = 'https://graph.microsoft.com/v1.0';

// Convert any date/time value to military time "MM/DD/YY HH:MM"
function toMilitaryDT(val) {
  if (!val && val !== 0) return '';
  if (typeof val === 'number') {
    const ms  = (val - 25569) * 86400 * 1000;
    const d   = new Date(Math.round(ms));
    return `${String(d.getUTCMonth()+1).padStart(2,'0')}/${String(d.getUTCDate()).padStart(2,'0')}/${String(d.getUTCFullYear()).slice(-2)} ${String(d.getUTCHours()).padStart(2,'0')}:${String(d.getUTCMinutes()).padStart(2,'0')}`;
  }
  if (val instanceof Date) {
    return `${String(val.getMonth()+1).padStart(2,'0')}/${String(val.getDate()).padStart(2,'0')}/${String(val.getFullYear()).slice(-2)} ${String(val.getHours()).padStart(2,'0')}:${String(val.getMinutes()).padStart(2,'0')}`;
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

async function readWorksheetValues(siteId, fileId, token, context) {
  const wbBase = `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook`;

  // Open session
  const sesRes = await fetch(`${wbBase}/createSession`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ persistChanges: false }),
  });
  if (!sesRes.ok) {
    if (context) context.log('[RCS read] Session open failed:', sesRes.status, await sesRes.text().catch(()=>''));
    return [];
  }
  const sesData = await sesRes.json();
  const sid = sesData.id;
  if (!sid) {
    if (context) context.log('[RCS read] No session ID returned:', JSON.stringify(sesData));
    return [];
  }
  const wbHdr = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid, 'Content-Type': 'application/json' };

  try {
    const sheetsRes = await fetch(`${wbBase}/worksheets`, { headers: wbHdr });
    if (!sheetsRes.ok) {
      if (context) context.log('[RCS read] Worksheets fetch failed:', sheetsRes.status);
      return [];
    }
    const sheets = (await sheetsRes.json()).value || [];
    if (!sheets.length) { if (context) context.log('[RCS read] No worksheets'); return []; }
    const wsId = sheets[0].id;
    if (context) context.log(`[RCS read] worksheet id=${wsId}`);

    // Try usedRange first, fall back to explicit large range
    let rows = [];
    const usedRes = await fetch(`${wbBase}/worksheets/${wsId}/usedRange?$select=values`, { headers: wbHdr });
    if (usedRes.ok) {
      const usedData = await usedRes.json();
      rows = usedData.values || [];
      if (context) context.log(`[RCS read] usedRange returned ${rows.length} rows`);
    }
    // Fallback: read explicit range if usedRange returned nothing
    if (!rows.length) {
      const rangeRes = await fetch(`${wbBase}/worksheets/${wsId}/range(address='A1:L100')?$select=values`, { headers: wbHdr });
      if (rangeRes.ok) {
        const all = (await rangeRes.json()).values || [];
        rows = all.filter(r => String(r[0] || '').trim()); // only rows with data in col A
        if (context) context.log(`[RCS read] fallback range: ${all.length} total, ${rows.length} with data`);
      }
    }
    return rows;
  } finally {
    await fetch(`${wbBase}/closeSession`, { method: 'POST', headers: wbHdr }).catch(() => {});
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

      // ── 1. Read requested date and load matching Results Cache rows ───────
      let reqBody = {};
      try { reqBody = await request.json(); } catch (_) {}
      const rawDateFilter = String(reqBody.dateFilter || '').trim();
      const dateFilter = /^\d{6}$/.test(rawDateFilter) ? rawDateFilter : '';
      if (!dateFilter) {
        return { status: 400, jsonBody: { error: 'Select a valid date before importing radon results' } };
      }

      const rcListId = await getListId(siteId, 'Results Cache', token);
      if (!rcListId) throw new Error('Results Cache list not found');

      let rcItems = [];
      let rcNextUrl = `${GRAPH}/sites/${siteId}/lists/${rcListId}/items?$expand=fields&$top=999`;
      while (rcNextUrl) {
        const rr = await fetch(rcNextUrl, { headers: authHdr });
        if (!rr.ok) throw new Error(`Results Cache read failed (${rr.status})`);
        const dd = await rr.json();
        rcItems.push(...(dd.value || []));
        rcNextUrl = dd['@odata.nextLink'] || null;
      }

      const dateRcItems = rcItems.filter(i => {
        const id = String(i.fields?.LabID || '').trim();
        return id.startsWith(dateFilter) && !/\bREJ\b/i.test(id);
      });
      const byPrefix = { [dateFilter]: dateRcItems };

      const rtbListId = await getListId(siteId, 'Reports to be Billed', token);

      const controlFolder = process.env.SP_CONTROL_FOLDER ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
      const marker = 'Shared Documents/';
      const mIdx = controlFolder.indexOf(marker);
      const relPath = mIdx >= 0
        ? controlFolder.slice(mIdx + marker.length)
        : controlFolder.replace(/^\/+/, '');

      let totalUpdated = 0;
      const log = [];

      // ── 2. Process the selected day's RCS file ─────────────────────────────
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
        const rows = await readWorksheetValues(siteId, fileId, token, context);
        if (!rows.length) { log.push(`⚠️ ${rcsName} is empty`); continue; }

        // RCS column layout (fixed positions, no header row):
        // A=Lab Barcode, B=Bubble Y/N, C=Bubble Size, D=blank,
        // E=Date Drawn, F=Time Drawn, G=Date Tested, H=Time Tested,
        // I=Time Lap, J=Multiplier, K=Mean pCi/L, L=Results pCi/L
        const colA          = 0;  // Lab Barcode #
        const colDateTested = 6;  // G - Date Tested
        const colTimeTested = 7;  // H - Time Tested
        const colResult     = 11; // L - Results pCi/L
        // Detect if first row is a header row
        const firstRowIsHeader = rows.length > 0 &&
          String(rows[0][0] || '').toLowerCase().includes('barcode') ||
          String(rows[0][0] || '').toLowerCase().includes('lab');
        const startRow = firstRowIsHeader ? 1 : 0;

        // Build lookup: base lab ID → { result, dateTested, timeTested }
        const rcsData = {};
        for (let i = startRow; i < rows.length; i++) {
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
                RadonDate: toMilitaryDT(rcsRow.dateTested),
                RadonTime: toMilitaryDT(rcsRow.timeTested),
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
              // Strip any suffix (e.g. " RW") from labId to match RTB Title
              const baseLabId = labId.replace(/\s+RW\s*$/i, '').trim();
              const rtbMatch  = rtbItems.find(i =>
                String(i.fields?.Title || '').trim() === labId ||
                String(i.fields?.Title || '').trim() === baseLabId
              );
              if (!rtbMatch) {
                const rtbSample = rtbItems.slice(0,3).map(i => i.fields?.Title);
                log.push(`⚠️ ${labId}: no RTB match (searching as "${labId}" / "${baseLabId}"; RTB sample: ${JSON.stringify(rtbSample)})`);
              }
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
