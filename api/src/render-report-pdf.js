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
const TMPL_SPEC  = 'Arsenic Spec Report - Template';
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

async function hideSheet(siteId, itemId, wsId, token, sid) {
  // Set worksheet visibility to hidden so it doesn't appear in the PDF
  // Use relative path — gReq prepends ${GRAPH} automatically
  await gReq('PATCH',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}`,
    token, { visibility: 'hidden' }, sid
  ).catch(() => {});
}

async function fillSheet(siteId, itemId, wsId, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context, comments='', commentsCell='') {
  const _dbg = []; // debug collector
  _dbg.push('dtCollected=' + (meta?.dtCollected||'EMPTY') + ' location=' + (meta?.location||'EMPTY') + ' city=' + (meta?.city||'EMPTY'));
  const rr = await gReq('GET',
    `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/usedRange?$select=values,columnCount,address`,
    token, undefined, sid);
  if (!rr.ok) { context.log('usedRange failed:', rr.status); return; }
  const { values: rows, columnCount: nc, address: rangeAddr } = await rr.json();
  // Parse the starting row from the range address (e.g. "Sheet1!B12:J60" → startRow = 12)
  const startRowMatch = (rangeAddr || '').match(/[A-Z]+(\d+):/);
  const startRow = startRowMatch ? parseInt(startRowMatch[1]) : 1;
  if (!rows?.length) return;

  const base        = `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}`;
  const cellUpdates = [];
  const colorUpdates = [];

  const addCell = (r, c, val) => cellUpdates.push({
    url:  `${base}/range(address='${colLetter(c)}${r + 1}')`,
    body: { values: [[String(val || '')]] },
  });

  // ── Clear all conditional formatting so our fills show correctly ──────────
  // NOTE: We do NOT clear conditional formatting — the template's built-in
  // conditional rules color cells correctly when values are written.
  // Clearing them caused colors to disappear entirely.

  // ── Right-side header fields (Lab ID, dates) ─────────────────────────────
  // Template has: label | (merge gap) | value | (time cell for date fields)
  // Header cells written per-sheet outside fillSheet

  // ── Left-side fields — handled via direct writes (E24, J24, I10) ─────────
  // Attention block — client name, billing address, report email
  const attLbl = findLabel(rows, 'attention');
  // ── Write comments to specific cell ────────────────────────────────────────
  if (comments && commentsCell) {
    // commentsCell format: 'A48' → row=47 (0-based), col=0
    const ccMatch = commentsCell.match(/^([A-Z]+)(\d+)$/i);
    if (ccMatch) {
      const ccCol = ccMatch[1].toUpperCase().split('').reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
      const ccRow = parseInt(ccMatch[2], 10) - 1;
      addCell(ccRow, ccCol, comments);
    }
  }

  if (attLbl) {
    const m             = meta || {};
    const attCol        = attLbl.c + 1;
    const cityStateZip  = [m.city, m.state, m.zip].filter(Boolean).join(', ');
    const fullSampleAddr = [m.location, cityStateZip].filter(Boolean).join(', ');
    const attAddress    = m.billingAddress || fullSampleAddr || '';
    const rawClientName = m.clientName || m.customer || '';
    const isPublicClient = /^Public-/i.test(rawClientName);
    const rawNameStr    = rawClientName.replace(/^Public-/i, '').trim();
    const cleanName     = isPublicClient && rawNameStr.includes(', ')
      ? rawNameStr.split(', ').reverse().join(' ')
      : rawNameStr;
    addCell(attLbl.r,     attCol, cleanName);
    if (attAddress)  addCell(attLbl.r + 1, attCol, attAddress);
    const emailLine = m.reportEmail || m.email || '';
    if (emailLine)   addCell(attLbl.r + 2, attCol, emailLine);
  }



  // ── Location — address to the RIGHT of "Location:" label, city/state/zip below ──
  const lf = findLabel(rows, 'location:');
  if (lf) {
    // Scan right for first empty cell on the Location row
    let locValCol = lf.c + 1;
    for (let dc = 1; dc <= 6; dc++) {
      const cv = normalizeCell((rows[lf.r] || [])[lf.c + dc]);
      if (!cv) { locValCol = lf.c + dc; break; }
    }
    const cityLine = [meta.city, meta.state, meta.zip].filter(Boolean).join(', ');
    addCell(lf.r,     locValCol, meta.location || '');  // address on same row as Location:
    addCell(lf.r + 1, locValCol, cityLine);              // city/state/zip on next row
  }

  // Find parameter table header row
  let hdrRow = -1, colResult = -1, colPrepDT = -1, colAnalDT = -1, colQualifier = 5; // col F default
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
  context.log(`[pdf] hdrRow=${hdrRow} colResult=${colResult} colPrepDT=${colPrepDT} colAnalDT=${colAnalDT} colQualifier=${colQualifier}`);

  // Map parameter names in sheet → row index
  // Only scan rows immediately after the header — stop at blank row or notation row
  const pMap = {};
  if (hdrRow >= 0) {
    let emptyStreak = 0;
    for (let r = hdrRow + 1; r < rows.length; r++) {
      const row   = rows[r] || [];
      const nameA = normalizeCell(row[0]);
      const nameB = normalizeCell(row[1]);
      const name  = nameA || nameB;
      // Stop if we hit a notation/comments/footer row
      if (name && (name.startsWith('notation') || name.startsWith('comment') || 
                   name.startsWith('this report') || name.startsWith('analytical') ||
                   name.startsWith('authorized') || name.startsWith('reporting limit'))) break;
      if (!name) { emptyStreak++; if (emptyStreak >= 5) break; continue; }
      emptyStreak = 0;
      pMap[name] = r;
    }
  }
  const paramNames = params.map(p => normalizeCell(p.name));
  context.log(`[pdf] hdrRow=${hdrRow} colResult=${colResult} nc=${nc}`);
  context.log(`[pdf] Template param rows (col A): ${JSON.stringify(Object.keys(pMap))}`);
  context.log(`[pdf] Looking for (from params): ${JSON.stringify(paramNames)}`);
  // Also log what's in cols 0,1,2 of the first few rows after header
  for (let r = hdrRow+1; r < Math.min(hdrRow+5, rows.length); r++) {
    context.log(`[pdf] Row ${r}: col0="${(rows[r]||[])[0]}" col1="${(rows[r]||[])[1]}" col2="${(rows[r]||[])[2]}"`);
  }

  const toDelete = [];

  for (const [nameLow, ri] of Object.entries(pMap)) {
    // Match by exact name OR partial match (template may omit ", Total" suffix)
    const p = params.find(x => {
      const pn = normalizeCell(x.name);
      return pn === nameLow || pn.startsWith(nameLow) || nameLow.startsWith(pn);
    });
    if (!p) {
      toDelete.push(startRow + ri); // correct Excel row = usedRange start + row index
    } else {
      // Write result value and dates
      if (colResult >= 0 && p.value)              addCell(ri, colResult,    p.value);
      if (colPrepDT >= 0 && p.prepDT)             addCell(ri, colPrepDT,    p.prepDT);
      if (colAnalDT >= 0 && (p.analDT || p.time)) addCell(ri, colAnalDT,    p.analDT || p.time);
      if (p.qualifier)                             addCell(ri, colQualifier, p.qualifier);

    }
  }

  context.log(`[pdf] ${cellUpdates.length} cells, ${colorUpdates.length} colors, ${toDelete.length} rows to delete`);

  // Send cell value updates first
  for (let i = 0; i < cellUpdates.length; i += 20) {
    const resp = await graphBatch(cellUpdates.slice(i, i + 20), token, sid);
    const errs = resp.filter(r => parseInt(r.status) >= 400);
    if (errs.length) context.log('[pdf] Cell batch errors:', JSON.stringify(errs.slice(0, 2)));
  }

  // Delete unused rows (bottom to top), then re-apply borders to all remaining param rows
  if (toDelete.length > 0) {
    const deleteSet = new Set(toDelete);
    const descending = [...toDelete].sort((a, b) => b - a);

    // Step 1: Read border style from a non-deleted param row before we start
    let savedBorder = null;
    const sampleRow = [...Object.values(pMap)].find(ri => !deleteSet.has(startRow + ri));
    if (sampleRow !== undefined) {
      try {
        const sampleRowNum = startRow + sampleRow;
        const borderR = await fetch(
          `${GRAPH}/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/range(address='A${sampleRowNum}')/format/borders/Bottom`,
          { headers: { Authorization: `Bearer ${token}`, 'workbook-session-id': sid } }
        );
        if (borderR.ok) {
          const bd = await borderR.json();
          if (bd.style && bd.style !== 'None') savedBorder = { style: bd.style, color: bd.color, weight: bd.weight };
        }
      } catch(e) {}
    }

    // Step 2: Delete rows bottom to top
    for (const rowNum of descending) {
      const addr = `A${rowNum}:${colLetter((nc||10)-1)}${rowNum}`;
      await fetch(
        `${GRAPH}/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/range(address='${addr}')/delete`,
        {
          method: 'POST',
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', 'workbook-session-id': sid },
          body: JSON.stringify({ shift: 'Up' }),
        }
      ).catch(e => context.log('[fillSheet] delete error:', e.message));
    }
    context.log(`[pdf] Deleted ${toDelete.length} unused rows`);

    // Step 3: Re-apply bottom border to all remaining param rows (row numbers have shifted)
    if (savedBorder) {
      const remainingRows = Object.values(pMap)
        .filter(ri => !deleteSet.has(startRow + ri))
        .map(ri => {
          // Recalculate row number after deletions: count how many deleted rows were above this row
          const excelRow = startRow + ri;
          const deletedAbove = toDelete.filter(d => d < excelRow).length;
          return excelRow - deletedAbove;
        })
        .sort((a, b) => a - b);

      const lastCol = colLetter((nc||10)-1);
      const borderReqs = remainingRows.map(rowNum => ({
        url:  `${GRAPH}/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/range(address='A${rowNum}:${lastCol}${rowNum}')/format/borders/Bottom`,
        body: savedBorder,
      }));
      for (let i = 0; i < borderReqs.length; i += 20) {
        const batch = borderReqs.slice(i, i + 20);
        await fetch(`${GRAPH}/$batch`, {
          method: 'POST',
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({
            requests: batch.map((r, idx) => ({
              id: String(i + idx + 1), method: 'PATCH', url: r.url,
              headers: { 'Content-Type': 'application/json', 'workbook-session-id': sid },
              body: r.body,
            })),
          }),
        }).catch(() => {});
      }
      context.log(`[pdf] Re-applied borders to ${remainingRows.length} remaining rows`);
    }
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
    const isRadon       = reportData.isRadon;
    const isArsenicSpec = reportData.isArsenicSpec;
    const today     = new Date().toLocaleDateString('en-US', { timeZone: 'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });

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

    // Find Arsenic Spec sheet if present
    const specSheet = finalSheets.find(s => /arsenic.*spec/i.test(s.name));


  // Helper: write header cells for a specific worksheet
  const writeHeaders = async (wsId2, cells) => {
    const wsBase = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${wsId2}`;
    const wsHdr  = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
    const splitDT = (dt) => {
      if (!dt) return ['', ''];
      const i = dt.indexOf(' ');
      return i > 0 ? [dt.slice(0, i), dt.slice(i + 1)] : [dt, ''];
    };
    const [dc, tc] = splitDT(meta.dtCollected || meta.dateDrawn || '');
    const [dr, tr] = splitDT(meta.dtReceived  || meta.dateReceived || '');
    const vals = { labId, dc, tc, dr, tr, today: meta.dateReported || today };
    for (const [addr, key] of cells) {
      const val = vals[key] || '';
      if (val) await fetch(`${wsBase}/range(address='${addr}')`, {
        method: 'PATCH', headers: wsHdr, body: JSON.stringify({ values: [[val]] })
      }).catch(() => {});
    }
  };

    if (isRadon && radonSheet) {
      // Radon template: Lab ID=I7, Date=I8, Time=J8, DateRec=I9, TimeRec=J9, DateRep=I10
      await writeHeaders(radonSheet.id, [
        ['I7','labId'], ['I8','dc'], ['J8','tc'], ['I9','dr'], ['J9','tr'], ['I10','today']
      ]);
      await fillSheet(siteId, tempId, radonSheet.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context, reportData._comments || '', '');

      // Write known cells directly: authorized by, review date, analysis date/time
      const wsBase3 = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${radonSheet.id}`;
      const wbHdr3  = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
      const directWrites = [
        ['E24', authorizedBy || ''],
        ['J24', reviewDate   || ''],
      ];
      for (const [addr, val] of directWrites) {
        if (val) await fetch(`${wsBase3}/range(address='${addr}')`, {
          method: 'PATCH', headers: wbHdr3, body: JSON.stringify({ values: [[val]] })
        }).catch(()=>{});
      }
      await fitOnePage(radonSheet.id);
      // Delete Spec sheet for radon reports
      if (specSheet) {
        await gReq('DELETE', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${specSheet.id}`, token, null, sid);
      }
    } else if (isArsenicSpec && specSheet) {
      // Arsenic Speciation: write headers first, then fill params
      await writeHeaders(specSheet.id, [
        ['H7','labId'], ['H8','dc'], ['I8','tc'], ['H9','dr'], ['I9','tr'], ['H10','today']
      ]);
      await fillSheet(siteId, tempId, specSheet.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context, reportData._comments || '', 'A24');
      // Write authorized by and review date to D32/I32 for Arsenic Spec report
      const wsBaseSpec = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${specSheet.id}`;
      const wbHdrSpec  = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
      for (const [addr, val] of [['D32', authorizedBy||''],['I32', reviewDate||'']]) {
        if (val) await fetch(`${wsBaseSpec}/range(address='${addr}')`, {
          method: 'PATCH', headers: wbHdrSpec, body: JSON.stringify({ values: [[val]] })
        }).catch(()=>{});
      }
      await fitOnePage(specSheet.id);
      // Hide Lab Report and FHA sheets
      if (labSheet)   await hideSheet(siteId, tempId, labSheet.id,   token, sid);
      if (fhaSheet)   await hideSheet(siteId, tempId, fhaSheet.id,   token, sid);
      if (radonSheet) await hideSheet(siteId, tempId, radonSheet.id, token, sid);
    } else if (labSheet) {
      // Delete spec sheet entirely when not a speciation test
      if (specSheet) {
        await gReq('DELETE', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${specSheet.id}`, token, null, sid);
      }
      // Standard template: Lab ID=H7, Date=H8, Time=I8, DateRec=H9, TimeRec=I9, DateRep=H10
      await writeHeaders(labSheet.id, [
        ['H7','labId'], ['H8','dc'], ['I8','tc'], ['H9','dr'], ['I9','tr'], ['H10','today']
      ]);
      await fillSheet(siteId, tempId, labSheet.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context, reportData._comments || '', 'A48');

      // Write authorized by and review date directly to known cell positions
      const wsBaseLab = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${labSheet.id}`;
      const wbHdrLab  = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
      for (const [addr, val] of [['D57', authorizedBy||''],['I57', reviewDate||'']]) {
        if (val) await fetch(`${wsBaseLab}/range(address='${addr}')`, {
          method: 'PATCH', headers: wbHdrLab, body: JSON.stringify({ values: [[val]] })
        }).catch(()=>{});
      }
      await fitOnePage(labSheet.id);
    }

    if (fhaSheet && needsFHA && fhaParams.length) {
      // Write same header cells as standard template
      await writeHeaders(fhaSheet.id, [
        ['H7','labId'], ['H8','dc'], ['I8','tc'], ['H9','dr'], ['I9','tr'], ['H10','today']
      ]);
      await fillSheet(siteId, tempId, fhaSheet.id, fhaParams, meta, labId, authorizedBy, reviewDate, today, token, sid, context, reportData._comments || '', 'A27');
      // Write authorized by and review date
      const wsBaseFHA = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${fhaSheet.id}`;
      const wbHdrFHA  = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
      for (const [addr, val] of [['D35', authorizedBy||''],['I35', reviewDate||'']]) {
        if (val) await fetch(`${wsBaseFHA}/range(address='${addr}')`, {
          method: 'PATCH', headers: wbHdrFHA, body: JSON.stringify({ values: [[val]] })
        }).catch(()=>{});
      }
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

    // Build pretty filename: [address]_[abbrev]_[labId] Report.pdf
    const safe = s => (s || '').trim().replace(/[<>:"/\\|?*]/g, '').trim();
    const addr  = safe(reportData?.meta?.location || '');
    const abbr  = safe(reportData?.meta?.abbrev || reportData?.meta?.clientCode || '');
    const parts = [addr, abbr, labId].filter(Boolean);
    const reportFileName = parts.join('_') + (isRadon ? ' RW Report.pdf' : ' Report.pdf');
    return { status: 200, jsonBody: { success: true, pdfBase64, fileName: reportFileName } };
  }
});
