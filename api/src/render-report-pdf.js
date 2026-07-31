/**
 * render-report-pdf.js — Azure v5
 * Download template → re-upload as temp → fill via Workbook API $batch → export PDF
 * No async copy polling — synchronous download+upload is fast and reliable.
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

async function gReq(method, path, token, body, sid) {
  const h = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
  if (sid) h['workbook-session-id'] = sid;
  const o = { method, headers: h };
  if (body !== undefined) o.body = JSON.stringify(body);
  return fetch(`${GRAPH}${path}`, o);
}

async function graphBatch(reqs, token, sid) {
  const r = await fetch(`${GRAPH}/$batch`, {
    method:  'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body:    JSON.stringify({
      requests: reqs.map((r, i) => ({
        id:      String(i + 1),
        method:  r.method || 'PATCH',
        url:     r.url,
        headers: { 'Content-Type': 'application/json', ...(sid ? { 'workbook-session-id': sid } : {}) },
        body:    r.body,
      })),
    }),
  });
  return r.ok ? (await r.json()).responses || [] : [];
}

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

async function fillSheet(siteId, itemId, wsId, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context) {
  const rr = await gReq('GET',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/usedRange?$select=values,columnCount`,
    token, undefined, sid);
  if (!rr.ok) { context.log('usedRange failed:', rr.status); return; }
  const { values: rows, columnCount: nc } = await rr.json();
  if (!rows?.length) return;

  const base = `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}`;
  const cellUpdates = [], colorUpdates = [];

  const addCell = (r, c, val) => cellUpdates.push({
    url:  `${base}/range(address='${colLetter(c)}${r + 1}')`,
    body: { values: [[String(val || '')]] },
  });

  // Headers
  const hdrs = {
    'attention:':           meta.customer || '',
    'lab id number:':       labId,
    'date/time collected:': meta.dtCollected || '',
    'date/time received:':  meta.dtReceived  || '',
    'date reported:':       today,
    'authorized by:':       authorizedBy,
    'review date:':         reviewDate,
  };
  for (const [lbl, val] of Object.entries(hdrs)) {
    const f = findLabel(rows, lbl);
    if (f) addCell(f.r, f.c + 1, val);
  }

  // Location
  const lf = findLabel(rows, 'location:');
  if (lf) {
    addCell(lf.r + 1, lf.c, meta.location || '');
    addCell(lf.r + 2, lf.c, [meta.city, meta.state, meta.zip].filter(Boolean).join(', '));
  }

  // Parameter table
  let hdrRow = -1, colResult = -1, colPrepDT = -1, colAnalDT = -1;
  for (let r = 0; r < rows.length; r++) {
    const rl = (rows[r] || []).map(c => String(c || '').toLowerCase().trim());
    if (rl.some(c => c.includes('your result'))) {
      hdrRow = r;
      rl.forEach((c, i) => {
        if (c.includes('your result'))                         colResult = i;
        else if (c.includes('preparation') || c.includes('prep date')) colPrepDT = i;
        else if (c.includes('analysis date'))                  colAnalDT = i;
      });
      break;
    }
  }

  const toDelete = [];
  if (hdrRow >= 0) {
    const pMap = {};
    for (let r = hdrRow + 1; r < rows.length; r++) {
      const name = String((rows[r] || [])[0] || '').trim();
      if (name) pMap[name.toLowerCase()] = r;
    }
    const cx = { green: '#00B050', red: '#FF0000', blue: '#0070C0', none: '#FFFFFF' };
    for (const [nl, ri] of Object.entries(pMap)) {
      const p = params.find(x => x.name.toLowerCase().trim() === nl);
      if (!p) { toDelete.push(ri + 1); }
      else {
        colorUpdates.push({ url: `${base}/range(address='B${ri + 1}')/format/fill`, body: { color: cx[p.color||'none']||'#FFFFFF' } });
        if (colResult >= 0 && p.value)            addCell(ri, colResult, p.value);
        if (colPrepDT >= 0 && p.prepDT)           addCell(ri, colPrepDT, p.prepDT);
        if (colAnalDT >= 0 && (p.analDT||p.time)) addCell(ri, colAnalDT, p.analDT || p.time);
      }
    }
  }

  // Send cell + color updates in batches of 20
  for (let i = 0; i < cellUpdates.length; i += 20) await graphBatch(cellUpdates.slice(i, i + 20), token, sid);
  for (let i = 0; i < colorUpdates.length; i += 20) await graphBatch(colorUpdates.slice(i, i + 20), token, sid);

  // Delete unused rows bottom-to-top (must be sequential)
  for (const r of toDelete.sort((a, b) => b - a)) {
    const addr = `A${r}:${colLetter((nc||10) - 1)}${r}`;
    await gReq('POST', `${base}/range(address='${addr}')/delete`, token, { shift: 'Up' }, sid);
  }
  context.log(`[pdf] Sheet done: ${cellUpdates.length} cells, ${colorUpdates.length} colors, ${toDelete.length} deleted`);
}

async function hideSheet(siteId, itemId, wsId, token, sid) {
  await gReq('PATCH', `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}`,
    token, { visibility: 'Hidden' }, sid);
}

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

    // ── Download template ───────────────────────────────────────────────────
    const tmplPath = process.env.SP_REPORT_TEMPLATE ||
      '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Report Templates.xlsx';
    const dp = toDrivePath(tmplPath);

    let tmplBuffer, parentPath;
    try {
      const metaR = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${dp}?$select=id,parentReference`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!metaR.ok) throw new Error(`Template not found (${metaR.status})`);
      const metaD = await metaR.json();
      parentPath = dp.replace(/\/[^/]+$/, ''); // strip filename to get folder path

      const dlR = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${metaD.id}/content`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!dlR.ok) throw new Error(`Template download failed (${dlR.status})`);
      tmplBuffer = Buffer.from(await dlR.arrayBuffer());
      context.log('[pdf] Template downloaded:', tmplBuffer.length, 'bytes');
    } catch(e) { return { status: 500, jsonBody: { error: e.message } }; }

    // ── Upload as temp file (synchronous — no polling!) ─────────────────────
    const tempName = `TEMP_${labId}_${Date.now()}.xlsx`;
    let tempId;
    try {
      const upR = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${parentPath}/${encodeURIComponent(tempName)}:/content`, {
        method:  'PUT',
        headers: {
          Authorization:  `Bearer ${token}`,
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        },
        body: tmplBuffer,
      });
      if (!upR.ok) throw new Error(`Upload failed (${upR.status}): ${await upR.text().catch(()=>'')}`);
      const upD = await upR.json();
      tempId = upD.id;
      context.log('[pdf] Uploaded temp:', tempId);
    } catch(e) { return { status: 500, jsonBody: { error: e.message } }; }

    // ── Open Workbook session ───────────────────────────────────────────────
    let sid = null;
    try {
      const sr = await gReq('POST',
        `/sites/${siteId}/drive/items/${tempId}/workbook/createSession`,
        token, { persistChanges: true });
      if (sr.ok) sid = (await sr.json()).id || null;
      context.log('[pdf] Session:', sid ? 'OK' : 'none');
    } catch(e) { context.log('Session warn:', e.message); }

    // ── Get sheets ──────────────────────────────────────────────────────────
    let sheets = [];
    const wr = await gReq('GET', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`, token, undefined, sid);
    if (wr.ok) sheets = (await wr.json()).value || [];
    context.log('[pdf] Sheets:', sheets.map(s => s.name).join(', '));
    const ws = name => sheets.find(s => s.name === name);

    // ── Fill sheets ─────────────────────────────────────────────────────────
    if (isRadon) {
      const lab = ws(SHEET_LAB), fha = ws(SHEET_FHA);
      if (lab) await hideSheet(siteId, tempId, lab.id, token, sid);
      if (fha) await hideSheet(siteId, tempId, fha.id, token, sid);
    } else {
      const radon = ws(SHEET_RADON);
      if (radon) await hideSheet(siteId, tempId, radon.id, token, sid);

      const lab = ws(SHEET_LAB);
      if (lab) await fillSheet(siteId, tempId, lab.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context);

      const fha = ws(SHEET_FHA);
      if (fha) {
        if (needsFHA && fhaParams.length)
          await fillSheet(siteId, tempId, fha.id, fhaParams, meta, labId, authorizedBy, reviewDate, today, token, sid, context);
        else
          await hideSheet(siteId, tempId, fha.id, token, sid);
      }
    }

    // ── Close session + wait ────────────────────────────────────────────────
    if (sid) await gReq('POST', `/sites/${siteId}/drive/items/${tempId}/workbook/closeSession`, token, {}, sid).catch(()=>{});
    await new Promise(r => setTimeout(r, 3000));

    // ── Export PDF ──────────────────────────────────────────────────────────
    let pdfBase64;
    try {
      const pr = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}/content?format=pdf`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!pr.ok) throw new Error(`PDF export (${pr.status})`);
      pdfBase64 = Buffer.from(await pr.arrayBuffer()).toString('base64');
      context.log('[pdf] PDF size:', pdfBase64.length);
    } catch(e) {
      await gReq('DELETE', `/sites/${siteId}/drive/items/${tempId}`, token).catch(()=>{});
      return { status: 500, jsonBody: { error: e.message } };
    }

    await gReq('DELETE', `/sites/${siteId}/drive/items/${tempId}`, token).catch(()=>{});
    return { status: 200, jsonBody: { success: true, pdfBase64, fileName: `${labId}_COA.pdf` } };
  }
});
