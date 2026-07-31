/**
 * render-report-pdf.js — Azure v4
 * Uses Graph API $batch for fast parallel cell updates.
 * Preserves all template formatting (logo, colors, merged cells).
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

// ── Graph $batch — sends up to 20 requests in one HTTP call ──────────────────
async function graphBatch(requests, token, sid) {
  const batchRequests = requests.map((req, i) => ({
    id:      String(i + 1),
    method:  req.method || 'PATCH',
    url:     req.url,
    headers: { 'Content-Type': 'application/json', ...(sid ? { 'workbook-session-id': sid } : {}) },
    body:    req.body,
  }));

  const res = await fetch(`${GRAPH}/$batch`, {
    method:  'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body:    JSON.stringify({ requests: batchRequests }),
  });
  if (!res.ok) console.warn('$batch failed:', res.status);
  return res.ok ? (await res.json()).responses || [] : [];
}

// ── Single Graph call ─────────────────────────────────────────────────────────
async function gReq(method, path, token, body, sid) {
  const h = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
  if (sid) h['workbook-session-id'] = sid;
  const opts = { method, headers: h };
  if (body !== undefined) opts.body = JSON.stringify(body);
  return fetch(`${GRAPH}${path}`, opts);
}

// ── Find label cell in 2D array ───────────────────────────────────────────────
function findLabel(rows, label) {
  const l = label.toLowerCase().trim();
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < (rows[r] || []).length; c++) {
      const v = String(rows[r][c] || '').toLowerCase().trim();
      if (v === l || v.startsWith(l)) return { r, c };
    }
  }
  return null;
}

// ── Build all update requests for a worksheet ────────────────────────────────
async function buildSheetUpdates(siteId, itemId, wsId, params, meta, labId, authorizedBy, reviewDate, today, token, sid) {
  // Get used range
  const rr = await gReq('GET',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/usedRange?$select=values,columnCount`,
    token, undefined, sid);
  if (!rr.ok) return { cellUpdates: [], colorUpdates: [], toDelete: [] };
  const { values: rows, columnCount: nc } = await rr.json();
  if (!rows?.length) return { cellUpdates: [], colorUpdates: [], toDelete: [] };

  const base = `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}`;
  const cellUpdates = [], colorUpdates = [];

  const setCell = (r, c, val) => cellUpdates.push({
    url:  `${base}/range(address='${colLetter(c)}${r + 1}')`,
    body: { values: [[String(val || '')]] },
  });

  // Header fields
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
    if (f) setCell(f.r, f.c + 1, val);
  }

  // Location
  const locF = findLabel(rows, 'location:');
  if (locF) {
    setCell(locF.r + 1, locF.c, meta.location || '');
    setCell(locF.r + 2, locF.c, [meta.city, meta.state, meta.zip].filter(Boolean).join(', '));
  }

  // Find parameter table
  let hdrRow = -1, colResult = -1, colPrepDT = -1, colAnalDT = -1;
  for (let r = 0; r < rows.length; r++) {
    const rLow = (rows[r] || []).map(c => String(c || '').toLowerCase().trim());
    if (rLow.some(c => c.includes('your result'))) {
      hdrRow = r;
      rLow.forEach((c, i) => {
        if (c.includes('your result'))                   colResult = i;
        else if (c.includes('preparation') || c.includes('prep date')) colPrepDT = i;
        else if (c.includes('analysis date'))            colAnalDT = i;
      });
      break;
    }
  }

  const toDelete = [];
  if (hdrRow >= 0) {
    const paramRowMap = {};
    for (let r = hdrRow + 1; r < rows.length; r++) {
      const name = String((rows[r] || [])[0] || '').trim();
      if (name) paramRowMap[name.toLowerCase()] = r;
    }

    const colorHex = { green: '#00B050', red: '#FF0000', blue: '#0070C0', none: '#FFFFFF' };

    for (const [nameLow, rowIdx] of Object.entries(paramRowMap)) {
      const p = params.find(x => x.name.toLowerCase().trim() === nameLow);
      if (!p) {
        toDelete.push(rowIdx + 1); // 1-based
      } else {
        // Color cell (col B = index 1)
        const hex = colorHex[p.color || 'none'] || '#FFFFFF';
        colorUpdates.push({
          url:  `${base}/range(address='B${rowIdx + 1}')/format/fill`,
          body: { color: hex },
        });
        if (colResult >= 0 && p.value)             setCell(rowIdx, colResult, p.value);
        if (colPrepDT >= 0 && p.prepDT)            setCell(rowIdx, colPrepDT, p.prepDT);
        if (colAnalDT >= 0 && (p.analDT||p.time))  setCell(rowIdx, colAnalDT, p.analDT || p.time);
      }
    }
  }

  return { cellUpdates, colorUpdates, toDelete, nc: nc || 10 };
}

// ── Send updates in batches of 20 ────────────────────────────────────────────
async function sendBatches(updates, token, sid) {
  for (let i = 0; i < updates.length; i += 20) {
    await graphBatch(updates.slice(i, i + 20), token, sid);
  }
}

// ── Delete rows bottom-to-top ─────────────────────────────────────────────────
async function deleteRows(siteId, itemId, wsId, rows, nc, token, sid) {
  const sorted = [...rows].sort((a, b) => b - a);
  for (const r of sorted) {
    const addr = `A${r}:${colLetter(nc - 1)}${r}`;
    await gReq('POST',
      `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/range(address='${addr}')/delete`,
      token, { shift: 'Up' }, sid);
  }
}

// ── Hide sheet ────────────────────────────────────────────────────────────────
async function hideSheet(siteId, itemId, wsId, token, sid) {
  await gReq('PATCH',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}`,
    token, { visibility: 'Hidden' }, sid);
}

// ─────────────────────────────────────────────────────────────────────────────
app.http('render-report-pdf', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {
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

    // ── Find + copy template ────────────────────────────────────────────────
    const tmplPath = process.env.SP_REPORT_TEMPLATE ||
      '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Report Templates.xlsx';
    const dp = toDrivePath(tmplPath);

    let tmplId, parentRef;
    try {
      const r = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${dp}?$select=id,parentReference`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!r.ok) throw new Error(`Template not found (${r.status})`);
      const d = await r.json();
      tmplId = d.id; parentRef = d.parentReference;
    } catch(e) { return { status: 500, jsonBody: { error: e.message } }; }

    const tempName = `TEMP_${labId}_${Date.now()}.xlsx`;
    let tempId;
    try {
      const cr = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tmplId}/copy`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ parentReference: parentRef, name: tempName }),
      });
      if (!cr.ok) throw new Error(`Copy failed (${cr.status})`);

      const monUrl = cr.headers.get('Location');
      // Poll until complete (up to 30s)
      for (let i = 0; i < 15; i++) {
        await new Promise(r => setTimeout(r, 2000));
        const pr = await fetch(monUrl, { headers: { Authorization: `Bearer ${token}` } });
        const pd = await pr.json().catch(() => ({}));
        context.log('[render-pdf] Copy status:', pd.status);
        if (pd.status === 'completed') { tempId = pd.resourceId; break; }
        if (pd.status === 'failed')    throw new Error('Copy failed');
      }
      if (!tempId) throw new Error('Copy timed out');
    } catch(e) { return { status: 500, jsonBody: { error: e.message } }; }

    context.log('[render-pdf] Temp:', tempId);

    // ── Open session ────────────────────────────────────────────────────────
    let sid = null;
    try {
      const sr = await gReq('POST',
        `/sites/${siteId}/drive/items/${tempId}/workbook/createSession`,
        token, { persistChanges: true });
      if (sr.ok) sid = (await sr.json()).id || null;
      context.log('[render-pdf] Session:', sid ? 'OK' : 'none');
    } catch(e) { context.log('Session warning:', e.message); }

    // ── Get sheets ──────────────────────────────────────────────────────────
    let sheets = [];
    try {
      const wr = await gReq('GET', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`, token, undefined, sid);
      if (wr.ok) sheets = (await wr.json()).value || [];
      context.log('[render-pdf] Sheets:', sheets.map(s=>s.name).join(', '));
    } catch(e) { context.log('Sheets warning:', e.message); }

    const ws = name => sheets.find(s => s.name === name);

    try {
      if (isRadon) {
        const lab = ws(SHEET_LAB), fha = ws(SHEET_FHA);
        if (lab) await hideSheet(siteId, tempId, lab.id, token, sid);
        if (fha) await hideSheet(siteId, tempId, fha.id, token, sid);
      } else {
        const radon = ws(SHEET_RADON);
        if (radon) await hideSheet(siteId, tempId, radon.id, token, sid);

        // Fill Lab sheet
        const lab = ws(SHEET_LAB);
        if (lab) {
          const { cellUpdates, colorUpdates, toDelete, nc } =
            await buildSheetUpdates(siteId, tempId, lab.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid);
          await sendBatches(cellUpdates, token, sid);
          await sendBatches(colorUpdates, token, sid);
          await deleteRows(siteId, tempId, lab.id, toDelete, nc, token, sid);
          context.log(`[render-pdf] Lab: ${cellUpdates.length} cells, ${colorUpdates.length} colors, ${toDelete.length} deleted`);
        }

        // Fill FHA sheet
        const fha = ws(SHEET_FHA);
        if (fha) {
          if (needsFHA && fhaParams.length) {
            const { cellUpdates, colorUpdates, toDelete, nc } =
              await buildSheetUpdates(siteId, tempId, fha.id, fhaParams, meta, labId, authorizedBy, reviewDate, today, token, sid);
            await sendBatches(cellUpdates, token, sid);
            await sendBatches(colorUpdates, token, sid);
            await deleteRows(siteId, tempId, fha.id, toDelete, nc, token, sid);
            context.log('[render-pdf] FHA filled');
          } else {
            await hideSheet(siteId, tempId, fha.id, token, sid);
          }
        }
      }
    } catch(e) {
      context.log('[render-pdf] Fill error:', e.message);
    }

    // ── Close session ───────────────────────────────────────────────────────
    if (sid) {
      await gReq('POST', `/sites/${siteId}/drive/items/${tempId}/workbook/closeSession`, token, {}, sid).catch(()=>{});
    }
    await new Promise(r => setTimeout(r, 3000));

    // ── Export PDF ──────────────────────────────────────────────────────────
    let pdfBase64;
    try {
      const pr = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}/content?format=pdf`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!pr.ok) throw new Error(`PDF export (${pr.status}): ${await pr.text().catch(()=>'')}`);
      pdfBase64 = Buffer.from(await pr.arrayBuffer()).toString('base64');
      context.log('[render-pdf] PDF:', pdfBase64.length, 'chars');
    } catch(e) {
      await gReq('DELETE', `/sites/${siteId}/drive/items/${tempId}`, token).catch(()=>{});
      return { status: 500, jsonBody: { error: e.message } };
    }

    // ── Cleanup ─────────────────────────────────────────────────────────────
    await gReq('DELETE', `/sites/${siteId}/drive/items/${tempId}`, token).catch(()=>{});

    return { status: 200, jsonBody: { success: true, pdfBase64, fileName: `${labId}_COA.pdf` } };
  }
});
