/**
 * render-report-pdf.js — Azure v2
 * Copies Report Templates.xlsx, fills via Excel Workbook API, exports PDF.
 */
const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');
const GRAPH        = 'https://graph.microsoft.com/v1.0';

const SHEET_LAB   = 'Lab Report - Template';
const SHEET_FHA   = 'FHA Lab Report - Template';
const SHEET_RADON = 'Radon Lab Report - Template';

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

async function graphFetch(method, path, token, body, sid) {
  const h = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
  if (sid) h['workbook-session-id'] = sid;
  const opts = { method, headers: h };
  if (body !== undefined) opts.body = JSON.stringify(body);
  return fetch(`${GRAPH}${path}`, opts);
}

async function setCell(siteId, itemId, wsId, addr, value, token, sid) {
  await graphFetch('PATCH',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/range(address='${addr}')`,
    token, { values: [[value]] }, sid);
}

async function setCellFill(siteId, itemId, wsId, addr, hex, token, sid) {
  await graphFetch('PATCH',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/range(address='${addr}')/format/fill`,
    token, { color: hex }, sid);
}

async function deleteRow(siteId, itemId, wsId, row1based, cols, token, sid) {
  const addr = `A${row1based}:${colLetter(cols - 1)}${row1based}`;
  await graphFetch('POST',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/range(address='${addr}')/delete`,
    token, { shift: 'Up' }, sid);
}

async function hideSheet(siteId, itemId, wsId, token, sid) {
  await graphFetch('PATCH',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}`,
    token, { visibility: 'Hidden' }, sid);
}

async function getRange(siteId, itemId, wsId, token, sid) {
  const r = await graphFetch('GET',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/usedRange?$select=values,rowCount,columnCount`,
    token, undefined, sid);
  if (!r.ok) return { values: [], rowCount: 0, columnCount: 0 };
  return r.json();
}

function findLabel(rows, label) {
  const l = label.toLowerCase().trim();
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < (rows[r] || []).length; c++) {
      const v = String(rows[r][c] || '').toLowerCase().trim();
      if (v === l || v.startsWith(l)) return { row: r, col: c };
    }
  }
  return null;
}

async function fillSheet(siteId, itemId, wsId, params, meta, labId, authorizedBy, reviewDate, today, token, sid) {
  const range = await getRange(siteId, itemId, wsId, token, sid);
  const rows  = range.values || [];
  const nc    = range.columnCount || 10;
  if (!rows.length) return;

  // Fill header fields (value goes in cell to the right of label)
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
    const f = findLabel(rows, lbl);
    if (!f) continue;
    await setCell(siteId, itemId, wsId, `${colLetter(f.col + 1)}${f.row + 1}`, val, token, sid);
  }

  // Location (below label)
  const locF = findLabel(rows, 'location:');
  if (locF) {
    await setCell(siteId, itemId, wsId, `${colLetter(locF.col)}${locF.row + 2}`, meta.location || '', token, sid);
    await setCell(siteId, itemId, wsId, `${colLetter(locF.col)}${locF.row + 3}`,
      [meta.city, meta.state, meta.zip].filter(Boolean).join(', '), token, sid);
  }

  // Find parameter table header row
  let hdrRow = -1, colResult = -1, colPrepDT = -1, colAnalDT = -1;
  for (let r = 0; r < rows.length; r++) {
    const rLow = (rows[r] || []).map(c => String(c || '').toLowerCase().trim());
    if (rLow.some(c => c.includes('your result'))) {
      hdrRow = r;
      rLow.forEach((c, i) => {
        if (c.includes('your result'))        colResult = i;
        else if (c.includes('preparation') || c.includes('prep date')) colPrepDT = i;
        else if (c.includes('analysis date') || c.includes('anal'))    colAnalDT = i;
      });
      break;
    }
  }

  if (hdrRow < 0) return;

  // Map param name → row index
  const paramRowMap = {};
  for (let r = hdrRow + 1; r < rows.length; r++) {
    const name = String((rows[r] || [])[0] || '').trim();
    if (name) paramRowMap[name.toLowerCase()] = r;
  }

  const colorHex = { green: '#00B050', red: '#FF0000', blue: '#0070C0', none: '#FFFFFF' };
  const toDelete = [];

  for (const [nameLow, rowIdx] of Object.entries(paramRowMap)) {
    const p = params.find(x => x.name.toLowerCase().trim() === nameLow);
    if (!p) {
      toDelete.push(rowIdx + 1);
    } else {
      const hex = colorHex[p.color || 'none'] || '#FFFFFF';
      // Color cell (col B = index 1)
      await setCellFill(siteId, itemId, wsId, `B${rowIdx + 1}`, hex, token, sid);
      // Result value
      if (colResult >= 0 && p.value) {
        await setCell(siteId, itemId, wsId, `${colLetter(colResult)}${rowIdx + 1}`, p.value, token, sid);
      }
      // Prep date
      if (colPrepDT >= 0 && p.prepDT) {
        await setCell(siteId, itemId, wsId, `${colLetter(colPrepDT)}${rowIdx + 1}`, p.prepDT, token, sid);
      }
      // Anal date
      if (colAnalDT >= 0 && (p.analDT || p.time)) {
        await setCell(siteId, itemId, wsId, `${colLetter(colAnalDT)}${rowIdx + 1}`, p.analDT || p.time || '', token, sid);
      }
    }
  }

  // Delete unused rows bottom-to-top
  for (const r of toDelete.sort((a, b) => b - a)) {
    await deleteRow(siteId, itemId, wsId, r, nc, token, sid);
  }
}

async function cleanup(siteId, tempId) {
  try {
    const t = await getToken();
    await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}`,
      { method: 'DELETE', headers: { Authorization: `Bearer ${t}` } });
  } catch(e) { console.warn('[render-pdf] cleanup failed:', e.message); }
}

app.http('render-report-pdf', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const siteId = process.env.SP_SITE_ID;
    let tempId   = null;

    const body = await request.json().catch(() => null);
    if (!body?.reportData) {
      return { status: 400, jsonBody: { error: 'reportData required' } };
    }

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
    catch(e) { return { status: 500, jsonBody: { error: 'Auth failed: ' + e.message } }; }

    // Find template
    const tmplPath = process.env.SP_REPORT_TEMPLATE ||
      '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Report Templates.xlsx';
    const dp = toDrivePath(tmplPath);
    let tmpl;
    try {
      const r = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${dp}?$select=id,parentReference`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!r.ok) throw new Error(`Template not found (${r.status}): ${tmplPath}`);
      tmpl = await r.json();
    } catch(e) { return { status: 500, jsonBody: { error: e.message } }; }

    // Copy template
    const tempName = `TEMP_${labId}_${Date.now()}.xlsx`;
    try {
      const cr = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tmpl.id}/copy`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ parentReference: tmpl.parentReference, name: tempName }),
      });
      if (!cr.ok) throw new Error(`Copy failed: ${cr.status}`);
      const monUrl = cr.headers.get('Location');
      if (!monUrl) throw new Error('No monitor URL from copy');

      // Poll for completion
      for (let i = 0; i < 25; i++) {
        await new Promise(r => setTimeout(r, 2000));
        const pr  = await fetch(monUrl, { headers: { Authorization: `Bearer ${token}` } });
        const pd  = await pr.json();
        if (pd.status === 'completed') { tempId = pd.resourceId; break; }
        if (pd.status === 'failed')    throw new Error('Copy operation failed');
      }
      if (!tempId) throw new Error('Copy timed out after 50s');
    } catch(e) {
      return { status: 500, jsonBody: { error: e.message } };
    }

    context.log(`[render-pdf] Temp file: ${tempId}`);

    // Open session
    let sid = null;
    try {
      const sr = await graphFetch('POST',
        `/sites/${siteId}/drive/items/${tempId}/workbook/createSession`,
        token, { persistChanges: true });
      if (sr.ok) sid = (await sr.json()).id || null;
    } catch(e) { context.log('[render-pdf] Session warning:', e.message); }

    // Get sheets
    let sheets = [];
    try {
      const wr = await graphFetch('GET',
        `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`,
        token, undefined, sid);
      sheets = wr.ok ? (await wr.json()).value || [] : [];
    } catch(e) { context.log('[render-pdf] Sheets error:', e.message); }

    const ws = name => sheets.find(s => s.name === name);
    context.log('[render-pdf] Sheets:', sheets.map(s => s.name).join(', '));

    try {
      if (isRadon) {
        const lab = ws(SHEET_LAB), fha = ws(SHEET_FHA);
        if (lab) await hideSheet(siteId, tempId, lab.id, token, sid);
        if (fha) await hideSheet(siteId, tempId, fha.id, token, sid);
      } else {
        const radon = ws(SHEET_RADON);
        if (radon) await hideSheet(siteId, tempId, radon.id, token, sid);

        const lab = ws(SHEET_LAB);
        if (lab) {
          await fillSheet(siteId, tempId, lab.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid);
          context.log('[render-pdf] Lab sheet filled');
        }

        const fha = ws(SHEET_FHA);
        if (fha) {
          if (needsFHA && fhaParams.length) {
            await fillSheet(siteId, tempId, fha.id, fhaParams, meta, labId, authorizedBy, reviewDate, today, token, sid);
            context.log('[render-pdf] FHA sheet filled');
          } else {
            await hideSheet(siteId, tempId, fha.id, token, sid);
          }
        }
      }
    } catch(e) {
      context.log('[render-pdf] Fill error:', e.message);
      await cleanup(siteId, tempId);
      return { status: 500, jsonBody: { error: 'Fill error: ' + e.message } };
    }

    // Close session
    if (sid) {
      await graphFetch('POST',
        `/sites/${siteId}/drive/items/${tempId}/workbook/closeSession`,
        token, {}, sid).catch(() => {});
    }

    // Wait for flush
    await new Promise(r => setTimeout(r, 4000));

    // Export PDF
    let pdfBase64;
    try {
      const pr = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}/content?format=pdf`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!pr.ok) throw new Error(`PDF export: ${pr.status}`);
      pdfBase64 = Buffer.from(await pr.arrayBuffer()).toString('base64');
      context.log('[render-pdf] PDF bytes:', pdfBase64.length);
    } catch(e) {
      await cleanup(siteId, tempId);
      return { status: 500, jsonBody: { error: 'PDF export failed: ' + e.message } };
    }

    // Cleanup and return
    await cleanup(siteId, tempId);
    return { status: 200, jsonBody: { success: true, pdfBase64, fileName: `${labId}_COA.pdf` } };
  }
});
