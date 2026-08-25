/**
 * patch-report-cell.js — Phase 4
 */
const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');
const GRAPH        = 'https://graph.microsoft.com/v1.0';

function calcColor(paramName, displayVal) {
  if (!displayVal && displayVal !== 0) return null;
  const s = String(displayVal).trim();
  if (!s) return null;
  const n = parseFloat(s);
  const num = !isNaN(n);
  const rl  = s.startsWith('<');
  if (paramName === 'Chloride, Total')          { if (rl||(num&&n<250)) return '#00CC44'; if (num&&n>=250) return '#0070C0'; return null; }
  if (paramName === 'Fluoride, Total')           { if (rl||(num&&n<1.9)) return '#00CC44'; if (num&&n>=1.9&&n<=3.9) return '#0070C0'; if (num&&n>=4) return '#FF0000'; return null; }
  if (paramName === 'Nitrite-Nitrogen, Total')   { if (rl||(num&&n<1))   return '#00CC44'; if (num&&n>=1)  return '#FF0000'; return null; }
  if (paramName === 'Nitrate-Nitrogen, Total')   { if (rl||(num&&n<10))  return '#00CC44'; if (num&&n>=10) return '#FF0000'; return null; }
  if (paramName === 'Arsenic, Total')            { if (rl||(num&&n<10))  return '#00CC44'; if (num&&n>=10) return '#FF0000'; return null; }
  if (paramName === 'Lead, Total')               { if (rl||(num&&n<15))  return '#00CC44'; if (num&&n>=15) return '#FF0000'; return null; }
  if (paramName === 'Uranium, Total')            { if (rl||(num&&n<30))  return '#00CC44'; if (num&&n>=30) return '#FF0000'; return null; }
  if (paramName === 'Copper, Total')             { if (rl||(num&&n<0.9)) return '#00CC44'; if (num&&n>=0.9&&n<=1.29) return '#0070C0'; if (num&&n>=1.3) return '#FF0000'; return null; }
  if (paramName === 'Iron, Total')               { if (rl||(num&&n<0.3)) return '#00CC44'; if (num&&n>=0.3)  return '#0070C0'; return null; }
  if (paramName === 'Manganese, Total')          { if (rl||(num&&n<0.05))return '#00CC44'; if (num&&n>=0.05) return '#0070C0'; return null; }
  if (paramName === 'Sodium, Total')             { if (num&&n>=20) return '#0070C0'; return null; }
  if (paramName === 'Antimony, Total')           { if (rl||(num&&n<0.006)) return '#00CC44'; if (num&&n>=0.006) return '#FF0000'; return null; }
  if (paramName === 'Cadmium, Total')            { if (rl||(num&&n<0.005)) return '#00CC44'; if (num&&n>=0.005) return '#FF0000'; return null; }
  if (paramName === 'Chromium, Total')           { if (rl||(num&&n<0.1))   return '#00CC44'; if (num&&n>=0.1)   return '#FF0000'; return null; }
  if (paramName === 'pH Electrometric')          { if (num&&n>=6.5&&n<=8.5) return '#00CC44'; if (num&&(n<6.5||n>8.5)) return '#0070C0'; return null; }
  if (paramName === 'Sulfate')                   { if (rl||(num&&n<250)) return '#00CC44'; if (num&&n>=250) return '#FF0000'; return null; }
  if (paramName === 'Total Coliform')            { if (rl||(num&&n<1))   return '#00CC44'; if (num&&n>=1)  return '#0070C0'; return null; }
  if (paramName === 'E. Coli')                   { if (rl||(num&&n<1))   return '#00CC44'; if (num&&n>=1)  return '#FF0000'; return null; }
  if (paramName === 'Radon Water')               { if (rl||(num&&n<4000))return '#00CC44'; if (num&&n>=4000) return '#0070C0'; return null; }
  return null;
}

function colLetter(n) {
  let s = '';
  while (n >= 0) { s = String.fromCharCode((n % 26) + 65) + s; n = Math.floor(n / 26) - 1; }
  return s;
}

function normalizeCell(v) {
  return String(v || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').toLowerCase().trim();
}

async function gReq(method, path, token, body, sid) {
  const h = { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' };
  if (sid) h['workbook-session-id'] = sid;
  const o = { method, headers: h };
  if (body !== undefined) o.body = JSON.stringify(body);
  return fetch(GRAPH + path, o);
}

app.http('patch-report-cell', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {
    try {
      const body = await request.json().catch(() => ({}));
      const { tempId, paramName, field, value, sheetType } = body;

      if (!tempId)    return { status: 400, jsonBody: { error: 'tempId required' } };
      if (!paramName) return { status: 400, jsonBody: { error: 'paramName required' } };
      if (!field)     return { status: 400, jsonBody: { error: 'field required' } };

      const siteId = process.env.SP_SITE_ID;
      const token  = await getToken();
      const wbBase = GRAPH + '/sites/' + siteId + '/drive/items/' + tempId + '/workbook';

      // Open session
      const sr = await fetch(wbBase + '/createSession', {
        method: 'POST',
        headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify({ persistChanges: true }),
      });
      if (!sr.ok) return { status: 500, jsonBody: { error: 'Session failed' } };
      const sid = (await sr.json()).id;

      // Find sheet
      const sheetsR = await gReq('GET', '/sites/' + siteId + '/drive/items/' + tempId + '/workbook/worksheets', token, undefined, sid);
      const sheets  = sheetsR.ok ? (await sheetsR.json()).value || [] : [];
      let sheet = null;
      if (sheetType === 'fha')   sheet = sheets.find(function(s) { return /^fha/i.test(s.name); });
      if (sheetType === 'radon') sheet = sheets.find(function(s) { return /^radon/i.test(s.name); });
      if (!sheet)                sheet = sheets.find(function(s) { return /^lab report/i.test(s.name); }) || sheets[0];
      if (!sheet) {
        await gReq('POST', '/sites/' + siteId + '/drive/items/' + tempId + '/workbook/closeSession', token, {}, sid).catch(function(){});
        return { status: 404, jsonBody: { error: 'Sheet not found' } };
      }

      const wsPath = '/sites/' + siteId + '/drive/items/' + tempId + '/workbook/worksheets/' + sheet.id;

      // Read sheet
      const rr = await gReq('GET', wsPath + '/usedRange?$select=values,columnCount', token, undefined, sid);
      if (!rr.ok) {
        await gReq('POST', '/sites/' + siteId + '/drive/items/' + tempId + '/workbook/closeSession', token, {}, sid).catch(function(){});
        return { status: 500, jsonBody: { error: 'Cannot read sheet' } };
      }
      const sheetData = await rr.json();
      const rows = sheetData.values || [];

      // Find header row
      var hdrRow = -1, colResult = -1, colPrepDT = -1, colAnalDT = -1, colQualifier = -1;
      for (var r = 0; r < rows.length; r++) {
        var rl = (rows[r] || []).map(function(c) { return normalizeCell(c); });
        if (rl.some(function(c) { return c.includes('your result') || c === 'result'; })) {
          hdrRow = r;
          rl.forEach(function(c, i) {
            if (c.includes('your result') || c === 'result')           colResult    = i;
            else if (c.includes('preparation') || c.includes('prep'))  colPrepDT    = i;
            else if (c.includes('analysis')    || c.includes('anal'))  colAnalDT    = i;
            else if (c.includes('qualifier')   || c.includes('qual'))  colQualifier = i;
          });
          break;
        }
      }

      // Find parameter row
      var nameLow = normalizeCell(paramName);
      var targetRow = -1;
      for (var r2 = hdrRow + 1; r2 < rows.length; r2++) {
        var rowName = normalizeCell((rows[r2] || [])[0]) || normalizeCell((rows[r2] || [])[1]);
        if (!rowName) continue;
        if (rowName === nameLow || rowName.startsWith(nameLow) || nameLow.startsWith(rowName)) {
          targetRow = r2;
          break;
        }
      }

      if (targetRow < 0) {
        await gReq('POST', '/sites/' + siteId + '/drive/items/' + tempId + '/workbook/closeSession', token, {}, sid).catch(function(){});
        return { status: 404, jsonBody: { error: 'Parameter not found: ' + paramName } };
      }

      // Build writes
      var patchReqs = [];
      var addr = function(col) { return colLetter(col) + (targetRow + 1); };

      if (field === 'value'     && colResult    >= 0) patchReqs.push({ url: wsPath + '/range(address=\'' + addr(colResult)    + '\')', body: { values: [[String(value || '')]] } });
      if (field === 'qualifier' && colQualifier >= 0) patchReqs.push({ url: wsPath + '/range(address=\'' + addr(colQualifier) + '\')', body: { values: [[String(value || '')]] } });
      if (field === 'prepDT'    && colPrepDT    >= 0) patchReqs.push({ url: wsPath + '/range(address=\'' + addr(colPrepDT)    + '\')', body: { values: [[String(value || '')]] } });
      if (field === 'analDT'    && colAnalDT    >= 0) patchReqs.push({ url: wsPath + '/range(address=\'' + addr(colAnalDT)    + '\')', body: { values: [[String(value || '')]] } });

      // Color indicator
      var newHex = null;
      if (field === 'value' && colResult > 0) {
        newHex = calcColor(paramName, value);
        if (newHex) patchReqs.push({ url: wsPath + '/range(address=\'' + colLetter(colResult - 1) + (targetRow + 1) + '\')/format/fill', body: { color: newHex } });
      }

      // Hardness recalculation
      var hardnessUpdate = null;
      if (field === 'value' && (paramName === 'Calcium, Total' || paramName === 'Magnesium, Total')) {
        var caVal = NaN, mgVal = NaN;
        var otherParam = paramName === 'Calcium, Total' ? 'Magnesium, Total' : 'Calcium, Total';
        var otherLow   = normalizeCell(otherParam);
        if (paramName === 'Calcium, Total')  caVal = parseFloat(value);
        else                                  mgVal = parseFloat(value);
        for (var r3 = hdrRow + 1; r3 < rows.length; r3++) {
          var rn = normalizeCell((rows[r3] || [])[0]) || normalizeCell((rows[r3] || [])[1]);
          if (rn === otherLow || rn.startsWith(otherLow) || otherLow.startsWith(rn)) {
            var ov = colResult >= 0 ? parseFloat(String((rows[r3] || [])[colResult] || '')) : NaN;
            if (paramName === 'Calcium, Total') mgVal = ov;
            else                                caVal = ov;
            break;
          }
        }
        if (!isNaN(caVal) && !isNaN(mgVal)) {
          var hardness = Math.round((caVal * 2.497 + mgVal * 4.118) * 100) / 100;
          var hardnessStr = String(hardness);
          var hardnessLow = 'hardness by calculation';
          for (var r4 = hdrRow + 1; r4 < rows.length; r4++) {
            var hn = normalizeCell((rows[r4] || [])[0]) || normalizeCell((rows[r4] || [])[1]);
            if (hn === hardnessLow || hn.startsWith('hardness')) {
              if (colResult >= 0) patchReqs.push({ url: wsPath + '/range(address=\'' + colLetter(colResult) + (r4 + 1) + '\')', body: { values: [[hardnessStr]] } });
              hardnessUpdate = { paramName: 'Hardness by calculation', value: hardnessStr };
              break;
            }
          }
        }
      }

      // Send batch
      if (patchReqs.length) {
        await fetch(GRAPH + '/$batch', {
          method:  'POST',
          headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
          body:    JSON.stringify({
            requests: patchReqs.map(function(r, i) {
              return { id: String(i + 1), method: 'PATCH', url: r.url, headers: { 'Content-Type': 'application/json', 'workbook-session-id': sid }, body: r.body };
            }),
          }),
        });
      }

      await gReq('POST', '/sites/' + siteId + '/drive/items/' + tempId + '/workbook/closeSession', token, {}, sid).catch(function(){});
      context.log('[patch-report-cell] OK — ' + paramName + ' ' + field + ' = ' + value);

      return { status: 200, jsonBody: { success: true, paramName, field, value, newHex, hardnessUpdate } };

    } catch(e) {
      context.log('[patch-report-cell] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  },
});
