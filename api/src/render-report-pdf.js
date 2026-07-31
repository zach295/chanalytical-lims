/**
 * render-report-pdf.js — Azure version
 * 1. Copies the Report Templates.xlsx from SharePoint
 * 2. Uses Excel Workbook API to fill values (preserves all formatting/logo)
 * 3. Deletes rows for parameters not in this test
 * 4. Downloads as PDF via Graph (?format=pdf)
 * 5. Cleans up the temp copy
 * 6. Returns base64 PDF
 *
 * POST { reportData, authorizedBy, reviewDate }
 */
const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

// ── Template location ─────────────────────────────────────────────────────────
// SP_REPORT_TEMPLATE env var — full SharePoint path to Report Templates.xlsx
// e.g. /sites/Laboratory/Shared Documents/Documents/Lab Scans/Report Templates.xlsx
function toDrivePath(fullPath) {
  const marker = 'Shared Documents/';
  const idx    = fullPath.indexOf(marker);
  const rel    = idx >= 0 ? fullPath.slice(idx + marker.length) : fullPath.replace(/^\/+/, '');
  return rel.split('/').map(s => encodeURIComponent(s)).join('/');
}

async function graphReq(method, path, token, body) {
  const opts = {
    method,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
  };
  if (body !== undefined) opts.body = JSON.stringify(body);
  const res = await fetch(`${GRAPH}${path}`, opts);
  if (!res.ok) {
    const err = await res.text().catch(() => '');
    throw new Error(`${method} ${path} → ${res.status}: ${err.slice(0, 200)}`);
  }
  return res;
}

// ── Find cell by searching a column for a label ───────────────────────────────
function findCellByLabel(rows, label) {
  const lbl = label.toLowerCase().trim();
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const cell = String(rows[r][c] || '').toLowerCase().trim();
      if (cell === lbl || cell.startsWith(lbl)) {
        return { row: r, col: c };
      }
    }
  }
  return null;
}

// ── Excel address helpers ──────────────────────────────────────────────────────
function colLetter(n) {
  let s = '';
  while (n >= 0) { s = String.fromCharCode((n % 26) + 65) + s; n = Math.floor(n / 26) - 1; }
  return s;
}
function cellAddr(row, col) { return `${colLetter(col)}${row + 1}`; }
function rangeAddr(r1, c1, r2, c2) { return `${cellAddr(r1,c1)}:${cellAddr(r2,c2)}`; }

// ── Update a single cell value (preserves formatting) ─────────────────────────
async function setCellValue(siteId, itemId, worksheetId, address, value, token, sessionId) {
  const headers = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
  };
  if (sessionId) headers['workbook-session-id'] = sessionId;

  const res = await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${address}')`,
    { method: 'PATCH', headers, body: JSON.stringify({ values: [[value]] }) }
  );
  if (!res.ok) {
    const err = await res.text().catch(() => '');
    console.warn(`setCellValue ${address}=${value}: ${res.status} ${err.slice(0,100)}`);
  }
}

// ── Set cell background color ─────────────────────────────────────────────────
async function setCellColor(siteId, itemId, worksheetId, address, hexColor, token, sessionId) {
  const headers = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
  };
  if (sessionId) headers['workbook-session-id'] = sessionId;

  await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${address}')/format/fill`,
    { method: 'PATCH', headers, body: JSON.stringify({ color: hexColor }) }
  ).catch(e => console.warn('setCellColor:', e.message));
}

// ── Delete a row by shifting up ───────────────────────────────────────────────
async function deleteRow(siteId, itemId, worksheetId, rowNum, totalCols, token, sessionId) {
  const headers = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
  };
  if (sessionId) headers['workbook-session-id'] = sessionId;

  const addr = `A${rowNum}:${colLetter(totalCols-1)}${rowNum}`;
  await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/range(address='${addr}')/delete`,
    { method: 'POST', headers, body: JSON.stringify({ shift: 'Up' }) }
  ).catch(e => console.warn(`deleteRow ${rowNum}:`, e.message));
}

// ── Get used range values ─────────────────────────────────────────────────────
async function getUsedRangeValues(siteId, itemId, worksheetId, token, sessionId) {
  const headers = { Authorization: `Bearer ${token}` };
  if (sessionId) headers['workbook-session-id'] = sessionId;

  const res = await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${worksheetId}/usedRange?$select=values,rowCount,columnCount`,
    { headers }
  );
  if (!res.ok) return { values: [], rowCount: 0, columnCount: 0 };
  return res.json();
}

// ── Color mapping ─────────────────────────────────────────────────────────────
function resultColorHex(color) {
  return color === 'green' ? '#00B050' : color === 'red' ? '#FF0000' : color === 'blue' ? '#0070C0' : '#E0E0E0';
}

// ── Fill a worksheet (COA or FHA) ─────────────────────────────────────────────
async function fillWorksheet(siteId, itemId, wsId, params, meta, labId, authorizedBy, reviewDate, today, token, sessionId) {
  const { values: rows, columnCount } = await getUsedRangeValues(siteId, itemId, wsId, token, sessionId);
  if (!rows || !rows.length) return;

  const totalCols = columnCount || 10;

  // ── Fill header fields ───────────────────────────────────────────────────────
  const headerMap = {
    'attention':             meta.customer || '',
    'lab id number':         labId,
    'date/time collected':   meta.dtCollected || '',
    'date/time received':    meta.dtReceived  || '',
    'date reported':         today,
    'location':              [meta.location, meta.city, meta.state, meta.zip].filter(Boolean).join(', '),
    'authorized by':         authorizedBy,
    'review date':           reviewDate,
  };

  for (const [label, value] of Object.entries(headerMap)) {
    const found = findCellByLabel(rows, label);
    if (!found) continue;
    // Value goes in the cell to the right of the label
    const valAddr = cellAddr(found.row, found.col + 1);
    await setCellValue(siteId, itemId, wsId, valAddr, value, token, sessionId);
  }

  // ── Find parameter rows ──────────────────────────────────────────────────────
  // Parameter names are in column A; results in the adjacent columns
  // Find the column indices for Result, EPA, Prep Date, Anal Date
  let headerRow = -1;
  let colResult = 2, colEPA = 3, colUnit = 4, colMethod = 5, colPrepDT = 6, colAnalDT = 7;

  for (let r = 0; r < rows.length; r++) {
    const rowLow = rows[r].map(c => String(c||'').toLowerCase().trim());
    if (rowLow.some(c => c.includes('your result') || c === 'result')) {
      headerRow = r;
      // Find column positions from header row
      rowLow.forEach((cell, i) => {
        if (cell.includes('your result') || cell === 'result') colResult = i;
        else if (cell.includes('epa')) colEPA = i;
        else if (cell === 'unit') colUnit = i;
        else if (cell === 'method') colMethod = i;
        else if (cell.includes('prep')) colPrepDT = i;
        else if (cell.includes('anal') || cell.includes('analysis')) colAnalDT = i;
      });
      break;
    }
  }

  // ── Build lookup of template params vs requested params ───────────────────────
  const requestedNames = new Set(params.map(p => p.name.toLowerCase().trim()));

  // Process rows from bottom up (so row deletion doesn't shift our indices)
  const rowsToDelete = [];

  for (let r = rows.length - 1; r > headerRow; r--) {
    const paramName = String(rows[r][0] || '').trim();
    if (!paramName) continue;
    const paramLow  = paramName.toLowerCase().trim();

    const matchedParam = params.find(p => p.name.toLowerCase().trim() === paramLow);

    if (!matchedParam) {
      // Parameter not in this test — mark for deletion
      rowsToDelete.push(r + 1); // 1-based row number for API
    } else {
      // Fill in the result values
      const color = resultColorHex(matchedParam.color || 'none');
      // Color indicator cell (column B = index 1)
      await setCellColor(siteId, itemId, wsId, cellAddr(r, 1), color, token, sessionId);
      // Result value
      await setCellValue(siteId, itemId, wsId, cellAddr(r, colResult), matchedParam.value || '', token, sessionId);
      // Prep date/time
      if (matchedParam.prepDT) await setCellValue(siteId, itemId, wsId, cellAddr(r, colPrepDT), matchedParam.prepDT, token, sessionId);
      // Analysis date/time
      if (matchedParam.analDT || matchedParam.time) await setCellValue(siteId, itemId, wsId, cellAddr(r, colAnalDT), matchedParam.analDT || matchedParam.time, token, sessionId);
    }
  }

  // Delete rows for unused parameters (bottom to top so indices stay valid)
  rowsToDelete.sort((a, b) => b - a);
  for (const rowNum of rowsToDelete) {
    await deleteRow(siteId, itemId, wsId, rowNum, totalCols, token, sessionId);
  }
}

// ── Main handler ──────────────────────────────────────────────────────────────
app.http('render-report-pdf', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    let tempFileId = null;
    const siteId   = process.env.SP_SITE_ID;

    try {
      const { reportData, authorizedBy, reviewDate } = await request.json();
      if (!reportData) return { status: 400, jsonBody: { error: 'reportData required' } };

      const token  = await getToken();
      const meta   = reportData.meta || {};
      const labId  = reportData.labId || '';
      const params     = reportData.activeParams || reportData.paramRows || [];
      const fhaParams  = reportData.fhaParams    || reportData.fhaRows   || [];
      const needsFHA   = reportData.needsFHA;
      const today      = new Date().toLocaleDateString('en-US',{month:'2-digit',day:'2-digit',year:'2-digit'});

      // ── Find template file ──────────────────────────────────────────────────
      const templatePath = process.env.SP_REPORT_TEMPLATE ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Report Templates.xlsx';
      const drivePath    = toDrivePath(templatePath);

      const tmplRes = await fetch(
        `${GRAPH}/sites/${siteId}/drive/root:/${drivePath}?$select=id,parentReference`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!tmplRes.ok) throw new Error(`Template not found at ${templatePath}`);
      const tmpl = await tmplRes.json();

      // ── Copy template to temp file ──────────────────────────────────────────
      const tempName = `TEMP_${labId}_${Date.now()}.xlsx`;
      const copyRes  = await fetch(
        `${GRAPH}/sites/${siteId}/drive/items/${tmpl.id}/copy`,
        {
          method:  'POST',
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body:    JSON.stringify({ parentReference: tmpl.parentReference, name: tempName }),
        }
      );
      if (!copyRes.ok) throw new Error(`Copy failed: ${copyRes.status}`);

      // Poll until copy is complete
      const monitorUrl = copyRes.headers.get('Location');
      let copyDone = false, tempId = null;
      for (let i = 0; i < 20; i++) {
        await new Promise(r => setTimeout(r, 1500));
        const pollRes  = await fetch(monitorUrl, { headers: { Authorization: `Bearer ${token}` } });
        const pollData = await pollRes.json();
        if (pollData.status === 'completed') { tempId = pollData.resourceId; copyDone = true; break; }
        if (pollData.status === 'failed') throw new Error('Copy operation failed');
      }
      if (!copyDone || !tempId) throw new Error('Copy timed out');
      tempFileId = tempId;

      context.log(`[render-pdf] Temp file: ${tempName} (${tempId})`);

      // ── Open persistent session ─────────────────────────────────────────────
      const sessRes = await fetch(
        `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/createSession`,
        {
          method:  'POST',
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body:    JSON.stringify({ persistChanges: true }),
        }
      );
      const sessData  = sessRes.ok ? await sessRes.json() : {};
      const sessionId = sessData.id || null;
      context.log(`[render-pdf] Session: ${sessionId ? 'OK' : 'none (changes may not persist)'}`);

      // ── Get worksheets ──────────────────────────────────────────────────────
      const wsRes  = await fetch(
        `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`,
        { headers: { Authorization: `Bearer ${token}`, ...(sessionId ? { 'workbook-session-id': sessionId } : {}) } }
      );
      const wsData = wsRes.ok ? await wsRes.json() : { value: [] };
      const sheets = wsData.value || [];
      context.log(`[render-pdf] Sheets: ${sheets.map(s=>s.name).join(', ')}`);

      // Find COA, FHA, and Notations sheets
      const coaSheet  = sheets.find(s => /coa|lab report|certificate/i.test(s.name) && !/fha|basic/i.test(s.name)) || sheets[0];
      const fhaSheet  = sheets.find(s => /fha|basic safety/i.test(s.name));
      const notSheet  = sheets.find(s => /notation/i.test(s.name));

      // ── Fill COA sheet ──────────────────────────────────────────────────────
      if (coaSheet) {
        await fillWorksheet(siteId, tempId, coaSheet.id, params, meta, labId, authorizedBy, reviewDate, today, token, sessionId);
        context.log('[render-pdf] COA sheet filled');
      }

      // ── Fill or hide FHA sheet ──────────────────────────────────────────────
      if (fhaSheet) {
        if (needsFHA && fhaParams.length) {
          await fillWorksheet(siteId, tempId, fhaSheet.id, fhaParams, meta, labId, authorizedBy, reviewDate, today, token, sessionId);
          context.log('[render-pdf] FHA sheet filled');
        } else {
          // Hide FHA sheet if not needed
          const hideHeaders = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
          if (sessionId) hideHeaders['workbook-session-id'] = sessionId;
          await fetch(
            `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${fhaSheet.id}`,
            { method: 'PATCH', headers: hideHeaders, body: JSON.stringify({ visibility: 'Hidden' }) }
          ).catch(e => context.log('Hide FHA:', e.message));
          context.log('[render-pdf] FHA sheet hidden (not needed)');
        }
      }

      // ── Close session ───────────────────────────────────────────────────────
      if (sessionId) {
        await fetch(
          `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/closeSession`,
          { method: 'POST', headers: { Authorization: `Bearer ${token}`, 'workbook-session-id': sessionId, 'Content-Type': 'application/json' } }
        ).catch(() => {});
      }

      // ── Wait for changes to settle ──────────────────────────────────────────
      await new Promise(r => setTimeout(r, 3000));

      // ── Export as PDF ───────────────────────────────────────────────────────
      const pdfRes = await fetch(
        `${GRAPH}/sites/${siteId}/drive/items/${tempId}/content?format=pdf`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!pdfRes.ok) throw new Error(`PDF export failed: ${pdfRes.status}`);
      const pdfBuffer = Buffer.from(await pdfRes.arrayBuffer());
      const pdfBase64 = pdfBuffer.toString('base64');

      context.log(`[render-pdf] PDF generated: ${pdfBuffer.length} bytes`);

      return {
        status: 200,
        jsonBody: { success: true, pdfBase64, fileName: `${labId}_COA.pdf` },
      };

    } catch(e) {
      context.log('[render-report-pdf] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    } finally {
      // ── Always clean up temp file ─────────────────────────────────────────
      if (tempFileId && siteId) {
        const token2 = await getToken().catch(() => null);
        if (token2) {
          await fetch(
            `${GRAPH}/sites/${siteId}/drive/items/${tempFileId}`,
            { method: 'DELETE', headers: { Authorization: `Bearer ${token2}` } }
          ).catch(e => console.warn('Cleanup failed:', e.message));
        }
      }
    }
  }
});
