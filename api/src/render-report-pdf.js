/**
 * render-report-pdf.js — Azure v6
 * 1. Download Report Templates.xlsx
 * 2. Re-upload as temp copy
 * 3. Delete unneeded sheets
 * 4. Rename tabs (remove "- Template")
 * 5. Delete rows for parameters not in this test
 * 6. Fill all data
 * 7. Export as PDF → return base64
 */
const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');
const GRAPH        = 'https://graph.microsoft.com/v1.0';

const TMPL_LAB   = 'Lab Report - Template';
const TMPL_FHA   = 'FHA Lab Report - Template';
const TMPL_RADON = 'Radon Lab Report - Template';
const TMPL_NOTES = 'Notations - Template';

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
function normalizeCell(v) {
  return String(v || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').toLowerCase().trim();
}
function findLabel(rows, label) {
  const l = normalizeCell(label);
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < (rows[r] || []).length; c++) {
      const v = normalizeCell(rows[r][c]);
      if (v === l || v.startsWith(l)) return { r, c };
    }
  }
  return null;
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
    body: JSON.stringify({
      requests: reqs.map((r, i) => ({
        id:      String(i + 1),
        method:  r.method || 'PATCH',
        url:     r.url,
        headers: { 'Content-Type': 'application/json', ...(sid ? { 'workbook-session-id': sid } : {}) },
        body:    r.body,
      })),
    }),
  });
  if (!r.ok) return [];
  const d = await r.json();
  return d.responses || [];
}

async function fillSheet(siteId, itemId, wsId, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context) {
  const rr = await gReq('GET',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/usedRange?$select=values,columnCount`,
    token, undefined, sid);
  if (!rr.ok) { context.log('usedRange failed:', rr.status); return; }
  const { values: rows, columnCount: nc } = await rr.json();
  if (!rows?.length) return;

  const base        = `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}`;
  const cellUpdates = [];
  const colorUpdates = [];

  const addCell = (r, c, val) => cellUpdates.push({
    url:  `${base}/range(address='${colLetter(c)}${r + 1}')`,
    body: { values: [[String(val || '')]] },
  });

  // ── Clear all conditional formatting so our fills show correctly ──────────
  const cfRes = await gReq('GET',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/conditionalFormats`,
    token, undefined, sid);
  if (cfRes.ok) {
    const cfs = (await cfRes.json()).value || [];
    for (const cf of cfs) {
      await gReq('DELETE',
        `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/conditionalFormats/${cf.id}`,
        token, undefined, sid).catch(() => {});
    }
    context.log(`[pdf] Cleared ${cfs.length} conditional formats`);
  }

  // ── Right-side header fields (Lab ID, dates) ─────────────────────────────
  const rightHdrs = {
    'lab id number:':       labId,
    'date/time collected:': meta.dtCollected || '',
    'date/time received:':  meta.dtReceived  || '',
    'date reported:':       today,
  };
  for (const [lbl, val] of Object.entries(rightHdrs)) {
    const f = findLabel(rows, lbl);
    if (!f) { context.log(`[pdf] Right label not found: "${lbl}"`); continue; }
    // Scan right for the FIRST empty cell — that's the value input cell
    let targetCol = f.c + 1;
    for (let dc = 1; dc <= 6; dc++) {
      const cv = normalizeCell((rows[f.r] || [])[f.c + dc]);
      if (!cv) { targetCol = f.c + dc; break; } // stop at first empty
    }
    addCell(f.r, targetCol, val);
    context.log(`[pdf] "${lbl}" → ${colLetter(targetCol)}${f.r+1} = "${val}"`);
  }

  // ── Attention block — name + address + email below ────────────────────────
  const attF = findLabel(rows, 'attention:');
  if (attF) {
    addCell(attF.r, attF.c + 1, meta.customer || '');
    // Row below name: client email
    const emailLine = meta.email || '';
    if (emailLine) addCell(attF.r + 1, attF.c + 1, emailLine);
  }

  // ── Left-side fields ──────────────────────────────────────────────────────
  const leftHdrs = {
    'date reported:': today,
    'authorized by:': authorizedBy,
    'review date:':   reviewDate,
  };
  for (const [lbl, val] of Object.entries(leftHdrs)) {
    const f = findLabel(rows, lbl);
    if (!f) continue;
    addCell(f.r, f.c + 1, val);
  }

  // ── Location — address to the RIGHT of "Location:" label, city/state/zip below ──
  const lf = findLabel(rows, 'location:');
  if (lf) {
    // Scan right for first empty cell on the Location row
    let locValCol = lf.c + 1;
    for (let dc = 1; dc <= 4; dc++) {
      const cv = normalizeCell((rows[lf.r] || [])[lf.c + dc]);
      if (!cv) { locValCol = lf.c + dc; break; }
    }
    const cityLine = [meta.city, meta.state, meta.zip].filter(Boolean).join(', ');
    addCell(lf.r,     locValCol, meta.location || '');  // address on same row as Location:
    addCell(lf.r + 1, locValCol, cityLine);              // city/state/zip on next row
  }

  // Find parameter table header row
  let hdrRow = -1, colResult = -1, colPrepDT = -1, colAnalDT = -1;
  for (let r = 0; r < rows.length; r++) {
    const rl = (rows[r] || []).map(c => normalizeCell(c));
    if (rl.some(c => c.includes('your result') || c === 'result')) {
      hdrRow = r;
      rl.forEach((c, i) => {
        if (c.includes('your result') || c === 'result')          colResult = i;
        else if (c.includes('preparation') || c.includes('prep')) colPrepDT = i;
        else if (c.includes('analysis') || c.includes('anal'))    colAnalDT = i;
      });
      break;
    }
  }
  context.log(`[pdf] hdrRow=${hdrRow} colResult=${colResult} colPrepDT=${colPrepDT} colAnalDT=${colAnalDT}`);

  // Map parameter names in sheet → row index
  const pMap = {};
  if (hdrRow >= 0) {
    for (let r = hdrRow + 1; r < rows.length; r++) {
      const name = normalizeCell((rows[r] || [])[0]);
      if (name) pMap[name] = r;
    }
  }

  const colorHex = { green: '#00B050', red: '#FF0000', blue: '#0070C0', none: '#FFFFFF' };
  const toDelete = [];

  for (const [nameLow, ri] of Object.entries(pMap)) {
    const p = params.find(x => normalizeCell(x.name) === nameLow);
    if (!p) {
      toDelete.push(ri + 1); // 1-based row number
    } else {
      // Color cell (col B = index 1)
      colorUpdates.push({
        url:  `${base}/range(address='B${ri + 1}')/format/fill`,
        body: { color: colorHex[p.color || 'none'] || '#FFFFFF' },
      });
      if (colResult >= 0 && p.value)             addCell(ri, colResult, p.value);
      if (colPrepDT >= 0 && p.prepDT)            addCell(ri, colPrepDT, p.prepDT);
      if (colAnalDT >= 0 && (p.analDT || p.time)) addCell(ri, colAnalDT, p.analDT || p.time);
    }
  }

  context.log(`[pdf] ${cellUpdates.length} cells, ${colorUpdates.length} colors, ${toDelete.length} rows to delete`);

  // Send cell value updates first
  for (let i = 0; i < cellUpdates.length; i += 20) {
    const resp = await graphBatch(cellUpdates.slice(i, i + 20), token, sid);
    const errs = resp.filter(r => parseInt(r.status) >= 400);
    if (errs.length) context.log('[pdf] Cell batch errors:', JSON.stringify(errs.slice(0, 2)));
  }

  // Delete unused rows bottom-to-top BEFORE setting colors
  // (so color row addresses are correct for the final sheet layout)
  const toDeleteSorted = toDelete.sort((a, b) => b - a);
  for (const r of toDeleteSorted) {
    const addr = `A${r}:${colLetter((nc || 10) - 1)}${r}`;
    const dr = await gReq('POST', `${base}/range(address='${addr}')/delete`, token, { shift: 'Up' }, sid);
    if (!dr.ok) context.log(`[pdf] Delete row ${r} failed:`, dr.status);
  }

  // Re-build color updates with FINAL row numbers (after deletions shifted rows up)
  // Count how many rows were deleted ABOVE each color row and adjust
  const finalColorUpdates = colorUpdates.map(cu => {
    // Extract row number from URL like .../range(address='B17')/...
    const m = cu.url.match(/address='B(\d+)'/);
    if (!m) return cu;
    const origRow = parseInt(m[1]);
    // Count deleted rows above this row
    const deletedAbove = toDeleteSorted.filter(d => d < origRow).length;
    const newRow = origRow - deletedAbove;
    return { ...cu, url: cu.url.replace(`B${origRow}`, `B${newRow}`) };
  });

  // Set colors LAST — after row deletions and after CF was cleared
  for (let i = 0; i < finalColorUpdates.length; i += 20) {
    await graphBatch(finalColorUpdates.slice(i, i + 20), token, sid);
  }
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

    // ── Step 1: Download template ───────────────────────────────────────────
    const tmplPath = process.env.SP_REPORT_TEMPLATE ||
      '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Report Templates.xlsx';
    const dp = toDrivePath(tmplPath);

    let tmplBuffer, folderPath;
    try {
      const metaR = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${dp}?$select=id`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!metaR.ok) throw new Error(`Template not found (${metaR.status})`);
      const { id: tmplId } = await metaR.json();

      folderPath = dp.replace(/\/[^/]+$/, '');

      const dlR = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tmplId}/content`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!dlR.ok) throw new Error(`Template download failed (${dlR.status})`);
      tmplBuffer = Buffer.from(await dlR.arrayBuffer());
      context.log('[pdf] Template downloaded:', tmplBuffer.length, 'bytes');
    } catch(e) { return { status: 500, jsonBody: { error: e.message } }; }

    // ── Step 2: Upload as temp copy ─────────────────────────────────────────
    const tempName = `TEMP_${labId}_${Date.now()}.xlsx`;
    let tempId;
    try {
      const upR = await fetch(
        `${GRAPH}/sites/${siteId}/drive/root:/${folderPath}/${encodeURIComponent(tempName)}:/content`, {
          method:  'PUT',
          headers: {
            Authorization:  `Bearer ${token}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          },
          body: tmplBuffer,
        });
      if (!upR.ok) throw new Error(`Upload failed (${upR.status})`);
      tempId = (await upR.json()).id;
      context.log('[pdf] Temp uploaded:', tempId);
    } catch(e) { return { status: 500, jsonBody: { error: e.message } }; }

    // ── Step 3: Open Workbook session ───────────────────────────────────────
    let sid = null;
    const sr = await gReq('POST',
      `/sites/${siteId}/drive/items/${tempId}/workbook/createSession`,
      token, { persistChanges: true });
    if (sr.ok) sid = (await sr.json()).id || null;
    context.log('[pdf] Session:', sid ? 'OK' : 'none');

    // ── Step 4: Get all sheets ──────────────────────────────────────────────
    const wr     = await gReq('GET', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`, token, undefined, sid);
    const sheets = wr.ok ? (await wr.json()).value || [] : [];
    context.log('[pdf] Sheets:', sheets.map(s => s.name).join(', '));
    const ws = name => sheets.find(s => s.name === name);

    // ── Step 5: Delete unneeded sheets ──────────────────────────────────────
    const sheetsToDelete = [];
    if (isRadon) {
      if (ws(TMPL_LAB))   sheetsToDelete.push(ws(TMPL_LAB));
      if (ws(TMPL_FHA))   sheetsToDelete.push(ws(TMPL_FHA));
    } else {
      if (ws(TMPL_RADON)) sheetsToDelete.push(ws(TMPL_RADON));
      if (!needsFHA || !fhaParams.length) {
        if (ws(TMPL_FHA)) sheetsToDelete.push(ws(TMPL_FHA));
      }
    }
    for (const sheet of sheetsToDelete) {
      const dr = await gReq('DELETE',
        `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${sheet.id}`,
        token, undefined, sid);
      context.log(`[pdf] Deleted sheet "${sheet.name}": ${dr.status}`);
    }

    // ── Step 6: Rename remaining sheets (remove "- Template") ──────────────
    const remainingWr = await gReq('GET',
      `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`,
      token, undefined, sid);
    const remaining = remainingWr.ok ? (await remainingWr.json()).value || [] : [];

    for (const sheet of remaining) {
      const newName = sheet.name.replace(/\s*-\s*Template\s*$/i, '').trim();
      if (newName !== sheet.name) {
        await gReq('PATCH',
          `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${sheet.id}`,
          token, { name: newName }, sid);
        context.log(`[pdf] Renamed "${sheet.name}" → "${newName}"`);
      }
    }

    // ── Step 7: Fill sheets ─────────────────────────────────────────────────
    // Re-fetch sheets after rename
    const finalWr    = await gReq('GET',
      `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`,
      token, undefined, sid);
    const finalSheets = finalWr.ok ? (await finalWr.json()).value || [] : [];
    context.log('[pdf] Final sheets:', finalSheets.map(s => s.name).join(', '));

    const labSheet   = finalSheets.find(s => /^lab report$/i.test(s.name) || /^lab report/i.test(s.name));
    const fhaSheet   = finalSheets.find(s => /^fha/i.test(s.name));
    const radonSheet = finalSheets.find(s => /^radon/i.test(s.name));

    const fitOnePage = async (wsId) => {
      // fitToWidth:1 ensures all columns fit on 1 page wide
      // fitToHeight:0 means unlimited rows (template already designed to fit 1 page)
      await gReq('PATCH',
        `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${wsId}/pageLayout`,
        token, { fitToPage: true, fitToWidth: 1, fitToHeight: 0, orientation: 'portrait', paperSize: 1 }, sid
      ).catch(() => {});
    };

    if (isRadon && radonSheet) {
      await fillSheet(siteId, tempId, radonSheet.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context);
      await fitOnePage(radonSheet.id);
    } else if (labSheet) {
      await fillSheet(siteId, tempId, labSheet.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context);
      await fitOnePage(labSheet.id);
    }

    if (fhaSheet && needsFHA && fhaParams.length) {
      await fillSheet(siteId, tempId, fhaSheet.id, fhaParams, meta, labId, authorizedBy, reviewDate, today, token, sid, context);
      await fitOnePage(fhaSheet.id);
    }

    // ── Step 8: Close session + wait ────────────────────────────────────────
    if (sid) {
      await gReq('POST',
        `/sites/${siteId}/drive/items/${tempId}/workbook/closeSession`,
        token, {}, sid).catch(() => {});
    }
    await new Promise(r => setTimeout(r, 3000));

    // ── Step 9: Export as PDF ───────────────────────────────────────────────
    let pdfBase64;
    try {
      const pr = await fetch(
        `${GRAPH}/sites/${siteId}/drive/items/${tempId}/content?format=pdf`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (!pr.ok) throw new Error(`PDF export (${pr.status})`);
      pdfBase64 = Buffer.from(await pr.arrayBuffer()).toString('base64');
      context.log('[pdf] PDF size:', pdfBase64.length);
    } catch(e) {
      await gReq('DELETE', `/sites/${siteId}/drive/items/${tempId}`, token).catch(() => {});
      return { status: 500, jsonBody: { error: e.message } };
    }

    // ── Step 10: Delete temp file ───────────────────────────────────────────
    await gReq('DELETE', `/sites/${siteId}/drive/items/${tempId}`, token).catch(() => {});
    context.log('[pdf] Temp file deleted');

    return { status: 200, jsonBody: { success: true, pdfBase64, fileName: `${labId}_COA.pdf` } };
  }
});
