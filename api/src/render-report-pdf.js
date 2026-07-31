/**
 * render-report-pdf.js — Azure v3
 * Download template → modify with SheetJS → upload → export PDF
 * Much faster than Workbook API (no session, no per-cell roundtrips)
 */
const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');
const GRAPH        = 'https://graph.microsoft.com/v1.0';

let XLSX;
try { XLSX = require('xlsx'); } catch(e) { console.warn('xlsx unavailable'); }

const SHEET_LAB   = 'Lab Report - Template';
const SHEET_FHA   = 'FHA Lab Report - Template';
const SHEET_RADON = 'Radon Lab Report - Template';

function toDrivePath(p) {
  const m = 'Shared Documents/';
  const i = p.indexOf(m);
  const r = i >= 0 ? p.slice(i + m.length) : p.replace(/^\/+/, '');
  return r.split('/').map(s => encodeURIComponent(s)).join('/');
}

// ── Find cell containing label text, return {r, c} ────────────────────────────
function findCell(ws, label) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const l     = label.toLowerCase().trim();
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell) continue;
      const v = String(cell.v || '').toLowerCase().trim();
      if (v === l || v.startsWith(l)) return { r, c };
    }
  }
  return null;
}

// ── Set cell value (creates if needed) ────────────────────────────────────────
function setVal(ws, r, c, val) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const existing = ws[addr] || {};
  ws[addr] = { ...existing, t: 's', v: String(val), w: String(val) };
}

// ── Set cell background color ─────────────────────────────────────────────────
function setFill(ws, r, c, hex) {
  const addr = XLSX.utils.encode_cell({ r, c });
  if (!ws[addr]) ws[addr] = { t: 's', v: '' };
  ws[addr].s = {
    ...(ws[addr].s || {}),
    fill: { fgColor: { rgb: hex.replace('#', '') }, patternType: 'solid' },
  };
}

// ── Delete a row from worksheet ───────────────────────────────────────────────
function deleteRow(ws, delRow) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  // Shift rows above the deleted row down
  for (let r = delRow; r < range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const from = XLSX.utils.encode_cell({ r: r + 1, c });
      const to   = XLSX.utils.encode_cell({ r, c });
      if (ws[from]) ws[to] = ws[from];
      else delete ws[to];
    }
  }
  // Clear last row
  for (let c = range.s.c; c <= range.e.c; c++) {
    delete ws[XLSX.utils.encode_cell({ r: range.e.r, c })];
  }
  range.e.r--;
  ws['!ref'] = XLSX.utils.encode_range(range);
  // Fix merged cells
  if (ws['!merges']) {
    ws['!merges'] = ws['!merges']
      .filter(m => m.s.r !== delRow && m.e.r !== delRow)
      .map(m => {
        if (m.s.r > delRow) { m.s.r--; m.e.r--; }
        return m;
      });
  }
}

// ── Fill one worksheet ────────────────────────────────────────────────────────
function fillWorksheet(ws, params, meta, labId, authorizedBy, reviewDate, today) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const colorHex = { green: '00B050', red: 'FF0000', blue: '0070C0', none: 'E0E0E0' };

  // Header fields — fill cell to the right of the label
  const headers = {
    'attention:':           meta.customer || '',
    'lab id number:':       labId,
    'date/time collected:': meta.dtCollected || '',
    'date/time received:':  meta.dtReceived  || '',
    'date reported:':       today,
    'authorized by:':       authorizedBy,
    'review date:':         reviewDate,
  };

  for (const [lbl, val] of Object.entries(headers)) {
    const f = findCell(ws, lbl);
    if (f) setVal(ws, f.r, f.c + 1, val);
  }

  // Location — fill below label
  const locF = findCell(ws, 'location:');
  if (locF) {
    setVal(ws, locF.r + 1, locF.c, meta.location || '');
    setVal(ws, locF.r + 2, locF.c, [meta.city, meta.state, meta.zip].filter(Boolean).join(', '));
  }

  // Find parameter table header row
  let hdrRow = -1, colResult = -1, colPrepDT = -1, colAnalDT = -1;
  for (let r = range.s.r; r <= range.e.r; r++) {
    const rowCells = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      rowCells.push({ c, v: String(cell?.v || '').toLowerCase().trim() });
    }
    if (rowCells.some(x => x.v.includes('your result'))) {
      hdrRow = r;
      rowCells.forEach(x => {
        if (x.v.includes('your result'))       colResult = x.c;
        else if (x.v.includes('prep'))         colPrepDT = x.c;
        else if (x.v.includes('analysis'))     colAnalDT = x.c;
      });
      break;
    }
  }
  if (hdrRow < 0) return;

  // Map parameter name → row index
  const paramRowMap = {};
  for (let r = hdrRow + 1; r <= range.e.r; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: range.s.c })];
    const name = String(cell?.v || '').trim();
    if (name) paramRowMap[name.toLowerCase()] = r;
  }

  // Fill results and collect rows to delete (process bottom-up)
  const toDelete = [];
  for (const [nameLow, rowIdx] of Object.entries(paramRowMap)) {
    const p = params.find(x => x.name.toLowerCase().trim() === nameLow);
    if (!p) {
      toDelete.push(rowIdx);
    } else {
      // Color indicator in col B (or adjacent to param name)
      const hex = colorHex[p.color || 'none'] || 'E0E0E0';
      setFill(ws, rowIdx, range.s.c + 1, hex); // col B

      if (colResult >= 0 && p.value) setVal(ws, rowIdx, colResult, p.value);
      if (colPrepDT >= 0 && p.prepDT) setVal(ws, rowIdx, colPrepDT, p.prepDT);
      if (colAnalDT >= 0 && (p.analDT || p.time)) setVal(ws, rowIdx, colAnalDT, p.analDT || p.time);
    }
  }

  for (const r of toDelete.sort((a, b) => b - a)) {
    deleteRow(ws, r);
  }
}

// ── Hide a sheet in workbook ──────────────────────────────────────────────────
function hideSheet(wb, sheetName) {
  const idx = wb.SheetNames.indexOf(sheetName);
  if (idx < 0) return;
  if (!wb.Workbook) wb.Workbook = {};
  if (!wb.Workbook.Sheets) wb.Workbook.Sheets = [];
  while (wb.Workbook.Sheets.length <= idx) wb.Workbook.Sheets.push({});
  wb.Workbook.Sheets[idx].Hidden = 1;
}

app.http('render-report-pdf', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {
    if (!XLSX) return { status: 500, jsonBody: { error: 'xlsx package not available' } };

    const siteId = process.env.SP_SITE_ID;
    const body   = await request.json().catch(() => null);
    if (!body?.reportData) return { status: 400, jsonBody: { error: 'reportData required' } };

    const { reportData, authorizedBy = '', reviewDate = '' } = body;
    const meta      = reportData.meta || {};
    const labId     = reportData.labId || '';
    const params    = reportData.activeParams || reportData.paramRows || [];
    const fhaParams = reportData.fhaParams    || reportData.fhaRows   || [];
    const needsFHA  = reportData.needsFHA;
    const isRadon   = reportData.isRadon;
    const today     = new Date().toLocaleDateString('en-US', { month:'2-digit', day:'2-digit', year:'2-digit' });

    let token;
    try { token = await getToken(); }
    catch(e) { return { status: 500, jsonBody: { error: 'Auth: ' + e.message } }; }

    // ── Download template ───────────────────────────────────────────────────
    const tmplPath = process.env.SP_REPORT_TEMPLATE ||
      '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Report Templates.xlsx';
    const dp = toDrivePath(tmplPath);

    let tmplBuffer, tmplItem;
    try {
      // Get item metadata
      const metaRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${dp}?$select=id,parentReference`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!metaRes.ok) throw new Error(`Template not found (${metaRes.status})`);
      tmplItem = await metaRes.json();

      // Download content
      const dlRes = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tmplItem.id}/content`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!dlRes.ok) throw new Error(`Template download failed (${dlRes.status})`);
      tmplBuffer = Buffer.from(await dlRes.arrayBuffer());
      context.log(`[render-pdf] Template downloaded: ${tmplBuffer.length} bytes`);
    } catch(e) { return { status: 500, jsonBody: { error: e.message } }; }

    // ── Read and modify workbook ────────────────────────────────────────────
    let modBuffer;
    try {
      const wb = XLSX.read(tmplBuffer, { type: 'buffer', cellStyles: true });
      context.log('[render-pdf] Sheets:', wb.SheetNames.join(', '));

      if (isRadon) {
        hideSheet(wb, SHEET_LAB);
        hideSheet(wb, SHEET_FHA);
      } else {
        hideSheet(wb, SHEET_RADON);

        const labWs = wb.Sheets[SHEET_LAB];
        if (labWs) {
          fillWorksheet(labWs, params, meta, labId, authorizedBy, reviewDate, today);
          context.log('[render-pdf] Lab sheet filled');
        }

        if (needsFHA && fhaParams.length) {
          const fhaWs = wb.Sheets[SHEET_FHA];
          if (fhaWs) {
            fillWorksheet(fhaWs, fhaParams, meta, labId, authorizedBy, reviewDate, today);
            context.log('[render-pdf] FHA sheet filled');
          }
        } else {
          hideSheet(wb, SHEET_FHA);
        }
      }

      modBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
      context.log('[render-pdf] Modified workbook: ' + modBuffer.length + ' bytes');
    } catch(e) { return { status: 500, jsonBody: { error: 'Modify failed: ' + e.message } }; }

    // ── Upload modified file to SharePoint ─────────────────────────────────
    const tempName = `TEMP_${labId}_${Date.now()}.xlsx`;
    const tempPath = toDrivePath(tmplPath.replace(/[^/]+\.xlsx$/i, '') + tempName);
    let tempId;
    try {
      const upRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${tempPath}:/content`, {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        },
        body: modBuffer,
      });
      if (!upRes.ok) throw new Error(`Upload failed (${upRes.status})`);
      const upData = await upRes.json();
      tempId = upData.id;
      context.log('[render-pdf] Uploaded temp file:', tempId);
    } catch(e) { return { status: 500, jsonBody: { error: e.message } }; }

    // ── Export as PDF ──────────────────────────────────────────────────────
    let pdfBase64;
    try {
      // Small wait for SharePoint to index the file
      await new Promise(r => setTimeout(r, 2000));
      const pdfRes = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}/content?format=pdf`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!pdfRes.ok) throw new Error(`PDF export (${pdfRes.status})`);
      pdfBase64 = Buffer.from(await pdfRes.arrayBuffer()).toString('base64');
      context.log('[render-pdf] PDF: ' + pdfBase64.length + ' chars');
    } catch(e) {
      // Cleanup on error
      await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}`,
        { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } }).catch(() => {});
      return { status: 500, jsonBody: { error: e.message } };
    }

    // ── Cleanup ────────────────────────────────────────────────────────────
    await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}`,
      { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } }).catch(() => {});

    return { status: 200, jsonBody: { success: true, pdfBase64, fileName: `${labId}_COA.pdf` } };
  }
});
