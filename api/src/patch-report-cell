/**
 * patch-report-cell.js
 * Writes a single edited cell value to the kept Excel temp file.
 * Called when user edits a field in the report preview.
 * Returns the updated cell value so the dashboard can confirm the write.
 */
const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');
const GRAPH        = 'https://graph.microsoft.com/v1.0';

function normalizeCell(v) {
  return String(v || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').toLowerCase().trim();
}
function colLetter(n) {
  let s = '';
  while (n >= 0) { s = String.fromCharCode((n % 26) + 65) + s; n = Math.floor(n / 26) - 1; }
  return s;
}

app.http('patch-report-cell', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body = await request.json().catch(() => ({}));
      const { tempId, paramName, field, value } = body;
      // field: 'value' | 'qualifier' | 'prepDT' | 'analDT' | 'meta'
      // metaField: for meta fields (customer, location, etc.)

      if (!tempId)    return { status: 400, jsonBody: { error: 'tempId required' } };
      if (!paramName) return { status: 400, jsonBody: { error: 'paramName required' } };

      const siteId = process.env.SP_SITE_ID;
      const token  = await getToken();

      // Open workbook session
      const sr = await fetch(
        `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/createSession`,
        { method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({ persistChanges: true }) }
      );
      if (!sr.ok) return { status: 500, jsonBody: { error: 'Cannot open workbook session' } };
      const sid = (await sr.json()).id;
      const wbHdr = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid, 'Content-Type': 'application/json' };

      // Get all sheets to find the right one
      const sheetsRes = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`, { headers: wbHdr });
      const sheets    = sheetsRes.ok ? (await sheetsRes.json()).value || [] : [];
      // Use the first visible sheet (Lab Report) for standard params
      const sheet     = sheets.find(s => /^lab report$/i.test(s.name) || /^radon/i.test(s.name)) || sheets[0];
      if (!sheet) return { status: 404, jsonBody: { error: 'No sheet found in workbook' } };

      const base = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${sheet.id}`;

      // Read used range to find parameter row
      const rr = await fetch(`${base}/usedRange?$select=values,columnCount`, { headers: wbHdr });
      if (!rr.ok) return { status: 500, jsonBody: { error: 'Cannot read sheet' } };
      const { values: rows, columnCount: nc } = await rr.json();

      // Find header row and column positions
      let hdrRow = -1, colResult = -1, colPrepDT = -1, colAnalDT = -1, colQualifier = 5;
      for (let r = 0; r < rows.length; r++) {
        const rl = (rows[r] || []).map(c => normalizeCell(c));
        if (rl.some(c => c.includes('your result') || c === 'result')) {
          hdrRow = r;
          rl.forEach((c, i) => {
            if (c.includes('your result') || c === 'result')          colResult    = i;
            else if (c.includes('preparation') || c.includes('prep')) colPrepDT    = i;
            else if (c.includes('analysis') || c.includes('anal'))    colAnalDT    = i;
            else if (c.includes('qualifier') || c.includes('qual'))   colQualifier = i;
          });
          break;
        }
      }

      // Find the row for this parameter
      const nameLow = normalizeCell(paramName);
      let targetRow = -1;
      for (let r = hdrRow + 1; r < rows.length; r++) {
        const rowNameA = normalizeCell((rows[r] || [])[0]);
        const rowNameB = normalizeCell((rows[r] || [])[1]);
        const rowName  = rowNameA || rowNameB;
        if (!rowName) continue;
        if (rowName === nameLow || rowName.startsWith(nameLow) || nameLow.startsWith(rowName)) {
          targetRow = r;
          break;
        }
      }

      if (targetRow < 0) {
        await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/closeSession`,
          { method: 'POST', headers: wbHdr }).catch(() => {});
        return { status: 404, jsonBody: { error: `Parameter not found in sheet: ${paramName}` } };
      }

      // Determine which column to update
      let targetCol = -1;
      if (field === 'value'     && colResult >= 0)    targetCol = colResult;
      if (field === 'qualifier' && colQualifier >= 0)  targetCol = colQualifier;
      if (field === 'prepDT'    && colPrepDT >= 0)     targetCol = colPrepDT;
      if (field === 'analDT'    && colAnalDT >= 0)     targetCol = colAnalDT;

      if (targetCol >= 0) {
        const addr = `${colLetter(targetCol)}${targetRow + 1}`;
        await fetch(`${base}/range(address='${addr}')`, {
          method: 'PATCH', headers: wbHdr,
          body: JSON.stringify({ values: [[String(value || '')]] }),
        });
        context.log(`[patch-cell] Updated ${paramName}.${field} → "${value}" at ${addr}`);
      }

      // Close session
      await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/closeSession`,
        { method: 'POST', headers: wbHdr }).catch(() => {});

      return { status: 200, jsonBody: { success: true, paramName, field, value } };
    } catch(e) {
      context.log('[patch-report-cell] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  },
});
