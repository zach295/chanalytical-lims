/**
 * patch-report-cell.js — Phase 4
 * Writes a single edited field to the kept Excel file (created by prepare-report.js).
 * Handles result values, qualifiers, prep dates, analysis dates, and meta fields.
 * When Ca or Mg are updated, also recalculates and writes Hardness.
 * Returns the new color hex for the updated parameter so the dashboard can update the dot.
 */

const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');
const GRAPH        = 'https://graph.microsoft.com/v1.0';

// ── Color rules — exact CF formulas from Excel template ───────────────────────
function calcColor(paramName, displayVal) {
  if (!displayVal && displayVal !== 0) return null;
  const s   = String(displayVal).trim();
  if (!s) return null;
  const n   = parseFloat(s);
  const num = !isNaN(n);
  const rl  = s.startsWith('<');

  switch (paramName) {
    case 'Chloride, Total':
      if (rl || (num && n < 250))  return '#00CC44';
      if (num && n >= 250)          return '#0070C0';
      return null;
    case 'Fluoride, Total':
      if (rl || (num && n < 1.9))            return '#00CC44';
      if (num && n >= 1.9 && n <= 3.9)       return '#0070C0';
      if (num && n >= 4)                      return '#FF0000';
      return null;
    case 'Nitrite-Nitrogen, Total':
      if (rl || (num && n < 1))   return '#00CC44';
      if (num && n >= 1)           return '#FF0000';
      return null;
    case 'Nitrate-Nitrogen, Total':
      if (rl || (num && n < 10))  return '#00CC44';
      if (num && n >= 10)          return '#FF0000';
      return null;
    case 'Arsenic, Total':
      if (rl || (num && n < 10))  return '#00CC44';
      if (num && n >= 10)          return '#FF0000';
      return null;
    case 'Lead, Total':
      if (rl || (num && n < 15))  return '#00CC44';
      if (num && n >= 15)          return '#FF0000';
      return null;
    case 'Uranium, Total':
      if (rl || (num && n < 30))  return '#00CC44';
      if (num && n >= 30)          return '#FF0000';
      return null;
    case 'Copper, Total':
      if (rl || (num && n < 0.9))            return '#00CC44';
      if (num && n >= 0.9 && n <= 1.29)      return '#0070C0';
      if (num && n >= 1.3)                    return '#FF0000';
      return null;
    case 'Iron, Total':
      if (rl || (num && n < 0.3))  return '#00CC44';
      if (num && n >= 0.3)          return '#0070C0';
      return null;
    case 'Manganese, Total':
      if (rl || (num && n < 0.05))  return '#00CC44';
      if (num && n >= 0.05)          return '#0070C0';
      return null;
    case 'Sodium, Total':
      if (num && n >= 20)  return '#0070C0';
      return null;
    case 'Antimony, Total':
      if (rl || (num && n < 0.006))  return '#00CC44';
      if (num && n >= 0.006)          return '#FF0000';
      return null;
    case 'Cadmium, Total':
      if (rl || (num && n < 0.005))  return '#00CC44';
      if (num && n >= 0.005)          return '#FF0000';
      return null;
    case 'Chromium, Total':
      if (rl || (num && n < 0.1))  return '#00CC44';
      if (num && n >= 0.1)          return '#FF0000';
      return null;
    case 'pH Electrometric':
      if (num && n >= 6.5 && n <= 8.5)  return '#00CC44';
      if (num && (n < 6.5 || n > 8.5))  return '#0070C0';
      return null;
    case 'Sulfate':
      if (rl || (num && n < 250))  return '#00CC44';
      if (num && n >= 250)          return '#FF0000';
      return null;
    case 'Total Coliform':
      if (rl || (num && n < 1))  return '#00CC44';
      if (num && n >= 1)          return '#0070C0';
      return null;
    case 'E. Coli':
      if (rl || (num && n < 1))  return '#00CC44';
      if (num && n >= 1)          return '#FF0000';
      return null;
    case 'Radon Water':
      if (rl || (num && n < 4000))   return '#00CC44';
      if (num && n >= 4000)           return '#0070C0';
      return null;
    default:
      return null;
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function colLetter(n) {
  let s = '';
  while (n >= 0) { s = String.fromCharCode((n % 26) + 65) + s; n = Math.floor(n / 26) - 1; }
  return s;
}
function normalizeCell(v) {
  return String(v || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').toLowerCase().trim();
}

async function gReq(method, path, token, body, sid) {
  const h = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
  if (sid) h['workbook-session-id'] = sid;
  const o = { method, headers: h };
  if (body !== undefined) o.body = JSON.stringify(body);
  return fetch(`${GRAPH}${path}`, o);
}

// ── Main handler ──────────────────────────────────────────────────────────────
app.http('patch-report-cell', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {
    try {
      const body = await request.json().catch(() => ({}));
      const { tempId, paramName, field, value, sheetType = 'lab' } = body;
      // field: 'value' | 'qualifier' | 'prepDT' | 'analDT'
      // sheetType: 'lab' | 'fha' | 'radon'

      if (!tempId)    return { status: 400, jsonBody: { error: 'tempId required' } };
      if (!paramName) return { status: 400, jsonBody: { error: 'paramName required' } };
      if (!field)     return { status: 400, jsonBody: { error: 'field required' } };

      const siteId = process.env.SP_SITE_ID;
      const token  = await getToken();

      // ── Open workbook session ─────────────────────────────────────────────
      const sr = await fetch(
        `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/createSession`,
        { method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({ persistChanges: true }) }
      );
      if (!sr.ok) return { status: 500, jsonBody: { error: 'Cannot open workbook session' } };
      const sid    = (await sr.json()).id;
      const wbBase = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook`;

      // ── Find the right sheet ──────────────────────────────────────────────
      const sheetsR = await gReq('GET', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`, token, undefined, sid);
      const sheets  = sheetsR.ok ? (await sheetsR.json()).value || [] : [];
      let sheet;
      if (sheetType === 'fha')   sheet = sheets.find(s => /^fha/i.test(s.name));
      if (sheetType === 'radon') sheet = sheets.find(s => /^radon/i.test(s.name));
      if (!sheet)                sheet = sheets.find(s => /^lab report/i.test(s.name)) || sheets[0];
      if (!sheet) {
        await gReq('POST', `${wbBase}/closeSession`, token, {}, sid).catch(() => {});
        return { status: 404, jsonBody: { error: 'Sheet not found' } };
      }

      const wsBase = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${sheet.id}`;

      // ── Read sheet to find header row and column positions ─────────────────
      const rr   = await gReq('GET', `${wsBase}/usedRange?$select=values,columnCount`, token, undefined, sid);
      if (!rr.ok) {
        await gReq('POST', `${wbBase}/closeSession`, token, {}, sid).catch(() => {});
        return { status: 500, jsonBody: { error: 'Cannot read sheet' } };
      }
      const { values: rows, columnCount: nc } = await rr.json();

      let hdrRow = -1, colResult = -1, colPrepDT = -1, colAnalDT = -1, colQualifier = -1;
      for (let r = 0; r < rows.length; r++) {
        const rl = (rows[r] || []).map(c => normalizeCell(c));
        if (rl.some(c => c.includes('your result') || c === 'result')) {
          hdrRow = r;
          rl.forEach((c, i) => {
            if (c.includes('your result') || c === 'result')           colResult    = i;
            else if (c.includes('preparation') || c.includes('prep'))  colPrepDT    = i;
            else if (c.includes('analysis')    || c.includes('anal'))  colAnalDT    = i;
            else if (c.includes('qualifier')   || c.includes('qual'))  colQualifier = i;
          });
          break;
        }
      }

      // ── Find the parameter row ────────────────────────────────────────────
      const nameLow = normalizeCell(paramName);
      let targetRow = -1;
      for (let r = hdrRow + 1; r < rows.length; r++) {
        const rowName = normalizeCell((rows[r] || [])[0]) || normalizeCell((rows[r] || [])[1]);
        if (!rowName) continue;
        if (rowName === nameLow || rowName.startsWith(nameLow) || nameLow.startsWith(rowName)) {
          targetRow = r;
          break;
        }
      }

      if (targetRow < 0) {
        await gReq('POST', `${wbBase}/closeSession`, token, {}, sid).catch(() => {});
        return { status: 404, jsonBody: { error: `Parameter not found: ${paramName}` } };
      }

      // ── Write value to the right column ───────────────────────────────────
      const writes = [];
      const addr   = (col) => `${colLetter(col)}${targetRow + 1}`;

      if (field === 'value'     && colResult    >= 0) writes.push([addr(colResult),    value]);
      if (field === 'qualifier' && colQualifier >= 0) writes.push([addr(colQualifier), value]);
      if (field === 'prepDT'    && colPrepDT    >= 0) writes.push([addr(colPrepDT),    value]);
      if (field === 'analDT'    && colAnalDT    >= 0) writes.push([addr(colAnalDT),    value]);

      // ── Update color indicator cell when value changes ─────────────────────
      let newHex = null;
      if (field === 'value' && colResult > 0) {
        newHex = calcColor(paramName, value);
        if (newHex) writes.push([`${colLetter(colResult - 1)}${targetRow + 1}`, null, newHex]);
      }

      // ── Hardness recalculation when Ca or Mg are updated ──────────────────
      const hardnessWrites = [];
      if (field === 'value' && (paramName === 'Calcium, Total' || paramName === 'Magnesium, Total')) {
        // Find Ca and Mg values — use new value for the edited param, read the other from sheet
        let caVal, mgVal;
        const otherParam = paramName === 'Calcium, Total' ? 'Magnesium, Total' : 'Calcium, Total';
        const otherLow   = normalizeCell(otherParam);

        if (paramName === 'Calcium, Total')   { caVal = parseFloat(value); }
        else                                  { mgVal = parseFloat(value); }

        // Find the other param's current value in the sheet
        for (let r = hdrRow + 1; r < rows.length; r++) {
          const rowName = normalizeCell((rows[r] || [])[0]) || normalizeCell((rows[r] || [])[1]);
          if (rowName === otherLow || rowName.startsWith(otherLow) || otherLow.startsWith(rowName)) {
            const otherVal = colResult >= 0 ? parseFloat(String((rows[r] || [])[colResult] || '')) : NaN;
            if (paramName === 'Calcium, Total')  mgVal = otherVal;
            else                                  caVal = otherVal;
            break;
          }
        }

        if (!isNaN(caVal) && !isNaN(mgVal)) {
          const hardness     = Math.round((caVal * 2.497 + mgVal * 4.118) * 100) / 100;
          const hardnessStr  = String(hardness);
          const hardnessLow  = 'hardness by calculation';

          for (let r = hdrRow + 1; r < rows.length; r++) {
            const rowName = normalizeCell((rows[r] || [])[0]) || normalizeCell((rows[r] || [])[1]);
            if (rowName === hardnessLow || rowName.startsWith('hardness')) {
              const hAddr = colResult >= 0 ? `${colLetter(colResult)}${r + 1}` : null;
              if (hAddr) hardnessWrites.push([hAddr, hardnessStr, null]);
              context.log(`[patch] Hardness recalculated: Ca=${caVal} Mg=${mgVal} → ${hardnessStr}`);
              break;
            }
          }
        }
      }

      // ── Batch all writes ──────────────────────────────────────────────────
      const allWrites = [...writes, ...hardnessWrites];
      const batchReqs = allWrites.map(([cellAddr, cellVal, colorHex]) => {
        if (colorHex) {
          return {
            url:  `${wsBase}/range(address='${cellAddr}')/format/fill`,
            body: { color: colorHex },
          };
        }
        return {
          url:  `${wsBase}/range(address='${cellAddr}')`,
          body: { values: [[String(cellVal ?? '')]] },
        };
      });

      // Send as Graph batch
      if (batchReqs.length) {
        const batchRes = await fetch(`${GRAPH}/$batch`, {
          method: 'POST',
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({
            requests: batchReqs.map((r, i) => ({
              id:      String(i + 1),
              method:  r.url.includes('/format/fill') ? 'PATCH' : 'PATCH',
              url:     r.url,
              headers: { 'Content-Type': 'application/json', 'workbook-session-id': sid },
              body:    r.body,
            })),
          }),
        });
        if (batchRes.ok) {
          const batchData = await batchRes.json();
          const errors    = (batchData.responses || []).filter(r => parseInt(r.status) >= 400);
          if (errors.length) context.log('[patch] Batch errors:', JSON.stringify(errors.slice(0, 3)));
          else context.log(`[patch] ${batchReqs.length} writes successful`);
        }
      }

      // ── Close session ─────────────────────────────────────────────────────
      await gReq('POST', `${wbBase}/closeSession`, token, {}, sid).catch(() => {});

      // Build hardness response for dashboard
      const hardnessUpdate = hardnessWrites.length > 0 ? {
        paramName: 'Hardness by calculation',
        value:     hardnessWrites[0][1],
        newHex:    null,
      } : null;

      return {
        status:   200,
        jsonBody: { success: true, paramName, field, value, newHex, hardnessUpdate },
      };

    } catch(e) {
      context.log('[patch-report-cell] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  },
});
