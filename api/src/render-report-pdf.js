/**
 * render-report-pdf.js — Azure version
 * 1. Copies Report Templates.xlsx from SharePoint
 * 2. Fills in header fields and parameter results via Excel Workbook API
 * 3. Deletes rows for parameters not in this test
 * 4. Hides unused sheets (Radon/FHA if not needed)
 * 5. Downloads as PDF via Graph ?format=pdf
 * 6. Cleans up temp copy
 * Returns { success, pdfBase64, fileName }
 */
const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

const SHEET_LAB    = 'Lab Report - Template';
const SHEET_FHA    = 'FHA Lab Report - Template';
const SHEET_RADON  = 'Radon Lab Report - Template';
const SHEET_NOTES  = 'Notations - Template';

function toDrivePath(p) {
  const m = 'Shared Documents/';
  const i = p.indexOf(m);
  const r = i >= 0 ? p.slice(i + m.length) : p.replace(/^\/+/, '');
  return r.split('/').map(s => encodeURIComponent(s)).join('/');
}

function colLetter(n) {
  let s = '';
  while (n >= 0) { s = String.fromCharCode((n % 26) + 65) + s; n = Math.floor(n / 26) - 1; }
  return s;
}

// ── Graph helpers ─────────────────────────────────────────────────────────────
async function gGet(path, token, sessionId) {
  const h = { Authorization: `Bearer ${token}` };
  if (sessionId) h['workbook-session-id'] = sessionId;
  const r = await fetch(`${GRAPH}${path}`, { headers: h });
  if (!r.ok) throw new Error(`GET ${path} → ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json();
}

async function gPatch(path, token, body, sessionId) {
  const h = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
  if (sessionId) h['workbook-session-id'] = sessionId;
  const r = await fetch(`${GRAPH}${path}`, { method: 'PATCH', headers: h, body: JSON.stringify(body) });
  if (!r.ok) {
    const t = await r.text().catch(() => '');
    console.warn(`PATCH ${path} → ${r.status}: ${t.slice(0,120)}`);
  }
  return r;
}

async function gPost(path, token, body, sessionId) {
  const h = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
  if (sessionId) h['workbook-session-id'] = sessionId;
  const r = await fetch(`${GRAPH}${path}`, { method: 'POST', headers: h, body: JSON.stringify(body) });
  if (!r.ok) {
    const t = await r.text().catch(() => '');
    console.warn(`POST ${path} → ${r.status}: ${t.slice(0,120)}`);
  }
  return r;
}

// ── Set a cell value ──────────────────────────────────────────────────────────
async function setCell(siteId, itemId, wsId, addr, value, token, sid) {
  await gPatch(
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/range(address='${addr}')`,
    token, { values: [[value]] }, sid
  );
}

// ── Set cell fill color ───────────────────────────────────────────────────────
async function setCellColor(siteId, itemId, wsId, addr, hex, token, sid) {
  await gPatch(
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/range(address='${addr}')/format/fill`,
    token, { color: hex }, sid
  );
}

// ── Delete a row (shift up) ───────────────────────────────────────────────────
async function deleteRow(siteId, itemId, wsId, rowNum, colCount, token, sid) {
  const addr = `A${rowNum}:${colLetter(colCount - 1)}${rowNum}`;
  await gPost(
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/range(address='${addr}')/delete`,
    token, { shift: 'Up' }, sid
  );
}

// ── Get used range ────────────────────────────────────────────────────────────
async function getUsedRange(siteId, itemId, wsId, token, sid) {
  return gGet(
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/usedRange?$select=values,rowCount,columnCount`,
    token, sid
  );
}

// ── Search rows for a label, return { row, col } ──────────────────────────────
function findLabel(rows, label) {
  const l = label.toLowerCase().trim();
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < (rows[r]||[]).length; c++) {
      if (String(rows[r][c]||'').toLowerCase().trim() === l) return { row: r, col: c };
    }
  }
  return null;
}

// ── Fill one worksheet (COA or FHA) ──────────────────────────────────────────
async function fillSheet(siteId, itemId, wsId, params, meta, labId, authorizedBy, reviewDate, today, token, sid) {
  const { values: rows, columnCount } = await getUsedRange(siteId, itemId, wsId, token, sid);
  if (!rows || !rows.length) return;

  const nc = columnCount || 10;

  // Header fields — value goes in the cell to the RIGHT of the label
  const headers = {
    'attention:':             meta.customer || '',
    'lab id number:':         labId,
    'date/time collected:':   meta.dtCollected || '',
    'date/time received:':    meta.dtReceived  || '',
    'date reported:':         today,
    'authorized by:':         authorizedBy,
    'review date:':           reviewDate,
  };

  // Location goes in cell BELOW "Location:" label
  const locFound = findLabel(rows, 'location:');
  if (locFound) {
    const locAddr = `${colLetter(locFound.col)}${locFound.row + 2}`;
    await setCell(siteId, itemId, wsId, locAddr, meta.location || '', token, sid);
    const cityAddr = `${colLetter(locFound.col)}${locFound.row + 3}`;
    await setCell(siteId, itemId, wsId, cityAddr,
      [meta.city, meta.state, meta.zip].filter(Boolean).join(', '), token, sid);
  }

  for (const [label, value] of Object.entries(headers)) {
    const found = findLabel(rows, label);
    if (!found) continue;
    const addr = `${colLetter(found.col + 1)}${found.row + 1}`;
    await setCell(siteId, itemId, wsId, addr, value, token, sid);
  }

  // ── Find parameter table header row ──────────────────────────────────────────
  let hdrRow = -1;
  let colResult = -1, colPrepDT = -1, colAnalDT = -1, colColor = 1;

  for (let r = 0; r < rows.length; r++) {
    const rowLow = (rows[r] || []).map(c => String(c || '').toLowerCase().trim());
    if (rowLow.some(c => c.includes('your result') || c === 'result')) {
      hdrRow = r;
      rowLow.forEach((c, i) => {
        if (c.includes('your result') || c === 'result') colResult = i;
        else if (c.includes('prep'))   colPrepDT = i;
        else if (c.includes('anal') || c === 'analysis date/time') colAnalDT = i;
      });
      break;
    }
  }

  if (hdrRow < 0 || colResult < 0) return; // Can't find table

  // ── Map parameter names in the sheet ─────────────────────────────────────────
  const paramMap = {}; // name.toLowerCase() → row index (0-based)
  for (let r = hdrRow + 1; r < rows.length; r++) {
    const name = String((rows[r] || [])[0] || '').trim();
    if (name) paramMap[name.toLowerCase()] = r;
  }

  const requestedNames = new Set(params.map(p => p.name.toLowerCase().trim()));
  const colorMap = { green: '#00B050', red: '#FF0000', blue: '#0070C0', none: '#FFFFFF' };

  // Collect rows to delete (process after fills so row numbers stay valid)
  const toDelete = [];

  for (const [nameLow, rowIdx] of Object.entries(paramMap)) {
    const matchedParam = params.find(p => p.name.toLowerCase().trim() === nameLow);

    if (!matchedParam) {
      toDelete.push(rowIdx + 1); // 1-based
    } else {
      // Fill color cell (col B = index 1)
      const hex = colorMap[matchedParam.color || 'none'] || '#FFFFFF';
      await setCellColor(siteId, itemId, wsId,
        `${colLetter(colColor)}${rowIdx + 1}`, hex, token, sid);

      // Fill result value
      if (colResult >= 0) {
        await setCell(siteId, itemId, wsId,
          `${colLetter(colResult)}${rowIdx + 1}`, matchedParam.value || '', token, sid);
      }

      // Fill prep date/time
      if (colPrepDT >= 0 && matchedParam.prepDT) {
        await setCell(siteId, itemId, wsId,
          `${colLetter(colPrepDT)}${rowIdx + 1}`, matchedParam.prepDT, token, sid);
      }

      // Fill analysis date/time
      if (colAnalDT >= 0 && (matchedParam.analDT || matchedParam.time)) {
        await setCell(siteId, itemId, wsId,
          `${colLetter(colAnalDT)}${rowIdx + 1}`, matchedParam.analDT || matchedParam.time || '', token, sid);
      }
    }
  }

  // Delete unused rows — bottom to top so row numbers stay valid
  toDelete.sort((a, b) => b - a);
  for (const rowNum of toDelete) {
    await deleteRow(siteId, itemId, wsId, rowNum, nc, token, sid);
  }
}

// ── Hide a worksheet ──────────────────────────────────────────────────────────
async function hideSheet(siteId, itemId, wsId, token, sid) {
  await gPatch(
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}`,
    token, { visibility: 'Hidden' }, sid
  );
}

// ── Main handler ──────────────────────────────────────────────────────────────
app.http('render-report-pdf', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {
    let tempId  = null;
    const siteId = process.env.SP_SITE_ID;

    try {
      const { reportData, authorizedBy, reviewDate } = await request.json();
      if (!reportData) return { status: 400, jsonBody: { error: 'reportData required' } };

      const token      = await getToken();
      const meta       = reportData.meta || {};
      const labId      = reportData.labId || '';
      const params     = reportData.activeParams || reportData.paramRows || [];
      const fhaParams  = reportData.fhaParams    || reportData.fhaRows   || [];
      const needsFHA   = reportData.needsFHA;
      const isRadon    = reportData.isRadon;
      const today      = new Date().toLocaleDateString('en-US',{month:'2-digit',day:'2-digit',year:'2-digit'});

      // ── Find template ───────────────────────────────────────────────────────
      const tmplPath = process.env.SP_REPORT_TEMPLATE ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Report Templates.xlsx';
      const dp = toDrivePath(tmplPath);
      const tmpl = await gGet(`/sites/${siteId}/drive/root:/${dp}?$select=id,parentReference`, token);

      // ── Copy template ───────────────────────────────────────────────────────
      const tempName = `TEMP_${labId}_${Date.now()}.xlsx`;
      const copyRes  = await gPost(
        `/sites/${siteId}/drive/items/${tmpl.id}/copy`,
        token, { parentReference: tmpl.parentReference, name: tempName }
      );
      const monUrl = copyRes.headers?.get?.('Location');
      if (!monUrl) throw new Error('No monitor URL from copy operation');

      // Poll until copy complete
      for (let i = 0; i < 30; i++) {
        await new Promise(r => setTimeout(r, 2000));
        const poll = await fetch(monUrl, { headers: { Authorization: `Bearer ${token}` } });
        const pd   = await poll.json();
        if (pd.status === 'completed') { tempId = pd.resourceId; break; }
        if (pd.status === 'failed') throw new Error('Copy failed');
      }
      if (!tempId) throw new Error('Copy timed out');
      context.log(`[render-pdf] Copied: ${tempName}`);

      // ── Open session ────────────────────────────────────────────────────────
      const sessRes  = await gPost(`/sites/${siteId}/drive/items/${tempId}/workbook/createSession`, token, { persistChanges: true });
      const sessJson = sessRes.ok ? await sessRes.json() : {};
      const sid      = sessJson.id || null;

      // ── Get worksheets ──────────────────────────────────────────────────────
      const wsData = await gGet(`/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`, token, sid);
      const sheets = wsData.value || [];
      const ws     = name => sheets.find(s => s.name === name);

      const labSheet   = ws(SHEET_LAB);
      const fhaSheet   = ws(SHEET_FHA);
      const radonSheet = ws(SHEET_RADON);
      const notesSheet = ws(SHEET_NOTES);

      context.log(`[render-pdf] Sheets: ${sheets.map(s=>s.name).join(', ')}`);

      if (isRadon) {
        // Radon report: show Radon sheet + Notations, hide Lab + FHA
        if (labSheet)   await hideSheet(siteId, tempId, labSheet.id,   token, sid);
        if (fhaSheet)   await hideSheet(siteId, tempId, fhaSheet.id,   token, sid);
        // TODO: fill radon sheet when needed
      } else {
        // Standard report: hide Radon sheet
        if (radonSheet) await hideSheet(siteId, tempId, radonSheet.id, token, sid);

        // Fill Lab Report sheet
        if (labSheet) {
          await fillSheet(siteId, tempId, labSheet.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid);
          context.log('[render-pdf] Lab sheet filled');
        }

        // Fill or hide FHA sheet
        if (fhaSheet) {
          if (needsFHA && fhaParams.length) {
            await fillSheet(siteId, tempId, fhaSheet.id, fhaParams, meta, labId, authorizedBy, reviewDate, today, token, sid);
            context.log('[render-pdf] FHA sheet filled');
          } else {
            await hideSheet(siteId, tempId, fhaSheet.id, token, sid);
            context.log('[render-pdf] FHA sheet hidden');
          }
        }
      }

      // ── Close session ───────────────────────────────────────────────────────
      if (sid) {
        await gPost(`/sites/${siteId}/drive/items/${tempId}/workbook/closeSession`, token, {}, sid);
      }

      // Wait for changes to flush
      await new Promise(r => setTimeout(r, 4000));

      // ── Export PDF ──────────────────────────────────────────────────────────
      const h       = { Authorization: `Bearer ${token}` };
      const pdfRes  = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}/content?format=pdf`, { headers: h });
      if (!pdfRes.ok) throw new Error(`PDF export: ${pdfRes.status}`);
      const pdfB64  = Buffer.from(await pdfRes.arrayBuffer()).toString('base64');
      context.log(`[render-pdf] PDF: ${pdfB64.length} chars`);

      return {
        status: 200,
        jsonBody: { success: true, pdfBase64: pdfB64, fileName: `${labId}_COA.pdf` },
      };

    } catch(e) {
      context.log('[render-report-pdf] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    } finally {
      // Always delete temp file
      if (tempId && siteId) {
        const t2 = await getToken().catch(() => null);
        if (t2) {
          await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}`, {
            method: 'DELETE', headers: { Authorization: `Bearer ${t2}` }
          }).catch(() => {});
          context.log('[render-pdf] Temp file deleted');
        }
      }
    }
  }
});
