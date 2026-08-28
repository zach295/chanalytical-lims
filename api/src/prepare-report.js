/**
 * prepare-report.js — Phase 2
 * Creates a filled copy of the Excel report template and keeps it alive.
 * Does NOT convert to PDF. Returns tempId (the kept file) and cellColors
 * (exact fill colors written to the color indicator cells).
 *
 * Called by the dashboard when a report is loaded.
 * The kept file is later used by export-pdf.js (Phase 3) for PDF generation
 * and patch-report-cell.js (Phase 4) for live edits.
 */

const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');
const GRAPH        = 'https://graph.microsoft.com/v1.0';

const TMPL_LAB   = 'Lab Report - Template';
const TMPL_FHA   = 'FHA Lab Report - Template';
const TMPL_RADON = 'Radon Lab Report - Template';
const TMPL_SPEC  = 'Arsenic Spec Report - Template';

// ── Helpers ───────────────────────────────────────────────────────────────────

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

// ── Color rules — exact conditional formatting from the Excel template ─────────
// Each parameter's rules match the template formulas exactly.
// Green = meets limits, Blue = see notation / advisory, Red = exceeds limit
function calcFillColor(paramName, displayVal) {
  if (!displayVal && displayVal !== 0) return null;
  const s   = String(displayVal).trim();
  if (!s) return null;
  const n   = parseFloat(s);
  const num = !isNaN(n);   // true when value is a real number
  const rl  = s.startsWith('<'); // <RL / below detection

  switch (paramName) {
    case 'Chloride, Total':
      // green: <2 (string) OR <250;  blue: >=250
      if (rl || (num && n < 250))  return '#00CC44';
      if (num && n >= 250)          return '#0070C0';
      return null;

    case 'Fluoride, Total':
      // green: <0.2 (string) OR <1.9;  blue: 1.9–3.9;  red: >=4
      if (rl || (num && n < 1.9))            return '#00CC44';
      if (num && n >= 1.9 && n <= 3.9)       return '#0070C0';
      if (num && n >= 4)                      return '#FF0000';
      return null;

    case 'Nitrite-Nitrogen, Total':
      // green: <0.2 (string) OR <1;  red: >=1
      if (rl || (num && n < 1))   return '#00CC44';
      if (num && n >= 1)           return '#FF0000';
      return null;

    case 'Nitrate-Nitrogen, Total':
      // green: <1 (string) OR <10;  red: >=10
      if (rl || (num && n < 10))  return '#00CC44';
      if (num && n >= 10)          return '#FF0000';
      return null;

    case 'Arsenic, Total':
      // green: <1 (string) OR <10;  red: >=10
      if (rl || (num && n < 10))  return '#00CC44';
      if (num && n >= 10)          return '#FF0000';
      return null;

    case 'Lead, Total':
      // green: <1 (string) OR <15;  red: >=15
      if (rl || (num && n < 15))  return '#00CC44';
      if (num && n >= 15)          return '#FF0000';
      return null;

    case 'Uranium, Total':
      // green: <1 (string) OR <30;  red: >=30
      if (rl || (num && n < 30))  return '#00CC44';
      if (num && n >= 30)          return '#FF0000';
      return null;

    case 'Copper, Total':
      // green: <0.001 (string) OR <0.9;  blue: 0.9–1.29;  red: >=1.3
      if (rl || (num && n < 0.9))            return '#00CC44';
      if (num && n >= 0.9 && n <= 1.29)      return '#0070C0';
      if (num && n >= 1.3)                    return '#FF0000';
      return null;

    case 'Iron, Total':
      // green: <0.05 (string) OR <0.3;  blue: >=0.3
      if (rl || (num && n < 0.3))  return '#00CC44';
      if (num && n >= 0.3)          return '#0070C0';
      return null;

    case 'Manganese, Total':
      // green: <0.001 (string) OR <0.05;  blue: >=0.05
      if (rl || (num && n < 0.05))  return '#00CC44';
      if (num && n >= 0.05)          return '#0070C0';
      return null;

    case 'Sodium, Total':
      // blue: >=20 only — no color defined below 20 in template
      if (num && n >= 20)  return '#0070C0';
      return null;

    case 'Antimony, Total':
      // green: <0.0005 (string) OR <0.006;  red: >=0.006
      if (rl || (num && n < 0.006))  return '#00CC44';
      if (num && n >= 0.006)          return '#FF0000';
      return null;

    case 'Cadmium, Total':
      // green: <0.002 (string) OR <0.005;  red: >=0.005
      if (rl || (num && n < 0.005))  return '#00CC44';
      if (num && n >= 0.005)          return '#FF0000';
      return null;

    case 'Chromium, Total':
      // green: <0.002 (string) OR <0.1;  red: >=0.1
      if (rl || (num && n < 0.1))  return '#00CC44';
      if (num && n >= 0.1)          return '#FF0000';
      return null;

    case 'pH Electrometric':
      // green: 6.5–8.5;  blue: outside range
      if (num && n >= 6.5 && n <= 8.5)  return '#00CC44';
      if (num && (n < 6.5 || n > 8.5))  return '#0070C0';
      return null;

    case 'Sulfate':
      // green: <40 (string) OR <250;  red: >=250
      if (rl || (num && n < 250))  return '#00CC44';
      if (num && n >= 250)          return '#FF0000';
      return null;

    case 'Total Coliform':
      // green: <1 strings OR <1;  blue: >=1
      if (rl || (num && n < 1))  return '#00CC44';
      if (num && n >= 1)          return '#0070C0';
      return null;

    case 'E. Coli':
      // green: <1 strings OR <1;  red: >=1  (E.coli is RED, not blue)
      if (rl || (num && n < 1))  return '#00CC44';
      if (num && n >= 1)          return '#FF0000';
      return null;

    case 'Radon Water':
      // green: <100 (string) OR <4000;  blue: >=4000
      if (rl || (num && n < 4000))   return '#00CC44';
      if (num && n >= 4000)           return '#0070C0';
      return null;

    case 'Turbidity':
      // green: <1 string OR <1;  blue: >=1
      if (rl || (num && n < 1))  return '#00CC44';
      if (num && n >= 1)          return '#0070C0';
      return null;
    default:
      return null;
  }
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
  return (await r.json()).responses || [];
}

// ── Fill one sheet ─────────────────────────────────────────────────────────────
// Returns { paramName: hexColor } for all params written to this sheet
async function fillSheet(siteId, itemId, wsId, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context, comments, commentsCell) {
  const base         = `/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}`;
  const cellUpdates  = [];
  const colorUpdates = [];
  const cellColors   = {}; // return value: paramName → hex

  const rr = await gReq('GET', `${base}/usedRange?$select=values,columnCount`, token, undefined, sid);
  if (!rr.ok) { context.log('[prepare] usedRange failed:', rr.status); return cellColors; }
  const { values: rows, columnCount: nc } = await rr.json();
  if (!rows?.length) return cellColors;

  const addCell = (r, c, val) => cellUpdates.push({
    url:  `${base}/range(address='${colLetter(c)}${r + 1}')`,
    body: { values: [[String(val || '')]] },
  });

  // ── Attention block ───────────────────────────────────────────────────────
  const attLbl = findLabel(rows, 'attention');
  if (attLbl) {
    const m            = meta || {};
    const attCol       = attLbl.c + 1;
    const cityStateZip = [m.city, m.state, m.zip].filter(Boolean).join(', ');
    const attAddress   = m.billingAddress || [m.location, cityStateZip].filter(Boolean).join(', ') || '';
    const rawName      = (m.clientName || m.customer || '').replace(/^Public-/i, '').trim();
    const cleanName    = rawName.includes(', ') ? rawName.split(', ').reverse().join(' ') : rawName;
    addCell(attLbl.r,     attCol, cleanName);
    if (attAddress)  addCell(attLbl.r + 1, attCol, attAddress);
    const emailLine = m.reportEmail || m.email || '';
    if (emailLine)   addCell(attLbl.r + 2, attCol, emailLine);
  }

  // ── Location block ────────────────────────────────────────────────────────
  const lf = findLabel(rows, 'location:');
  if (lf) {
    let locValCol = lf.c + 1;
    for (let dc = 1; dc <= 6; dc++) {
      if (!normalizeCell((rows[lf.r] || [])[lf.c + dc])) { locValCol = lf.c + dc; break; }
    }
    const cityLine = [meta.city, meta.state, meta.zip].filter(Boolean).join(', ');
    addCell(lf.r,     locValCol, meta.location || '');
    addCell(lf.r + 1, locValCol, cityLine);
  }

  // ── Comments ──────────────────────────────────────────────────────────────
  if (comments && commentsCell) {
    const cm = commentsCell.match(/^([A-Z]+)(\d+)$/i);
    if (cm) {
      const cc = cm[1].toUpperCase().split('').reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
      const cr = parseInt(cm[2], 10) - 1;
      addCell(cr, cc, comments);
    }
  }

  // ── Find parameter table header ───────────────────────────────────────────
  let hdrRow = -1, colResult = -1, colPrepDT = -1, colAnalDT = -1, colQualifier = 5;
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

  // ── Map parameter names to rows ───────────────────────────────────────────
  const pMap = {};
  if (hdrRow >= 0) {
    let emptyStreak = 0;
    for (let r = hdrRow + 1; r < rows.length; r++) {
      const name = normalizeCell((rows[r] || [])[0]) || normalizeCell((rows[r] || [])[1]);
      if (name && (name.startsWith('notation') || name.startsWith('comment') || name.startsWith('authorized'))) break;
      if (!name) { if (++emptyStreak >= 5) break; continue; }
      emptyStreak = 0;
      pMap[name] = r;
    }
  }

  const toHide = [];

  for (const [nameLow, ri] of Object.entries(pMap)) {
    const p = params.find(x => {
      const pn = normalizeCell(x.name);
      return pn === nameLow || pn.startsWith(nameLow) || nameLow.startsWith(pn);
    });
    if (!p) {
      toHide.push(ri + 1);
    } else {
      if (colResult >= 0 && p.value)               addCell(ri, colResult,    p.value);
      if (colPrepDT >= 0 && p.prepDT)              addCell(ri, colPrepDT,    p.prepDT);
      if (colAnalDT >= 0 && (p.analDT || p.time))  addCell(ri, colAnalDT,    p.analDT || p.time);
      if (p.qualifier)                              addCell(ri, colQualifier, p.qualifier);

      // Write explicit fill color to color indicator cell (col before result)
      if (colResult > 0) {
        const hex = calcFillColor(p.name, p.value);
        if (hex) {
          colorUpdates.push({ url: `${base}/range(address='${colLetter(colResult - 1)}${ri + 1}')/format/fill`, body: { color: hex } });
          cellColors[p.name] = hex;
        }
      }
    }
  }

  context.log(`[prepare] ${cellUpdates.length} cell writes, ${colorUpdates.length} color writes, ${toHide.length} rows to hide`);

  // Batch write values
  for (let i = 0; i < cellUpdates.length; i += 20) {
    await graphBatch(cellUpdates.slice(i, i + 20), token, sid);
  }

  // Batch write fill colors
  for (let i = 0; i < colorUpdates.length; i += 20) {
    await graphBatch(colorUpdates.slice(i, i + 20), token, sid);
  }

  // Delete unused rows (bottom to top so row numbers stay valid after each delete)
  if (toHide.length > 0) {
    const descending = [...toHide].sort((a, b) => b - a);
    const ranges = [];
    let re3 = descending[0], rs3 = descending[0];
    for (let i = 1; i < descending.length; i++) {
      if (descending[i] === rs3 - 1) rs3 = descending[i];
      else { ranges.push([rs3, re3]); re3 = rs3 = descending[i]; }
    }
    ranges.push([rs3, re3]);
    for (const [s, e] of ranges) {
      const addr = `A${s}:${colLetter((nc || 10) - 1)}${e}`;
      await fetch(`${GRAPH}/sites/${siteId}/drive/items/${itemId}/workbook/worksheets/${wsId}/range(address='${addr}')/delete`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', 'workbook-session-id': sid },
        body: JSON.stringify({ shift: 'Up' }),
      }).catch(e => context.log('[prepare] row delete error:', e.message));
    }
  }

  return cellColors;
}

// ── Main handler ──────────────────────────────────────────────────────────────
app.http('prepare-report', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {
    try {
    const siteId = process.env.SP_SITE_ID;
    const body   = await request.json().catch(() => null);
    // Ping mode — just touch file metadata to keep it alive
    if (body.ping && body.tempId) {
      try {
        const pingR = await fetch(`${GRAPH}/sites/${process.env.SP_SITE_ID}/drive/items/${body.tempId}?$select=id,name`,
          { headers: { Authorization: `Bearer ${await getToken()}` } });
        return { status: 200, jsonBody: { alive: pingR.ok } };
      } catch(e) { return { status: 200, jsonBody: { alive: false } }; }
    }

    if (!body?.reportData) return { status: 400, jsonBody: { error: 'reportData required' } };

    const { reportData, authorizedBy = '', reviewDate = '' } = body;
    const meta       = reportData.meta     || {};
    const labId      = reportData.labId    || '';
    const params     = reportData.activeParams || reportData.paramRows || [];
    const fhaParams  = reportData.fhaParams   || reportData.fhaRows   || [];
    const needsFHA   = reportData.needsFHA;
    const isRadon    = reportData.isRadon;
    const isArsenicSpec = reportData.isArsenicSpec;
    const today      = new Date().toLocaleDateString('en-US', { timeZone: 'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });

    let token;
    try { token = await getToken(); }
    catch(e) { return { status: 500, jsonBody: { error: 'Auth: ' + e.message } }; }


    // ── Step 1: Download template ─────────────────────────────────────────
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
      context.log('[prepare] Template downloaded:', tmplBuffer.length, 'bytes');
    } catch(e) { return { status: 500, jsonBody: { error: e.message } }; }

    // ── Step 2: Upload as named copy (delete existing first to avoid lock) ──
    const fileName = isRadon ? `${labId} RW Report.xlsx` : `${labId} Report.xlsx`;
    let tempId;
    try {
      // Delete any existing file with this name to avoid 423 Locked errors
      const existCheck = await fetch(
        `${GRAPH}/sites/${siteId}/drive/root:/${folderPath}/${encodeURIComponent(fileName)}?$select=id`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (existCheck.ok) {
        const existData = await existCheck.json();
        if (existData.id) {
          // Retry delete up to 3 times — file may be briefly locked
          for (let attempt = 0; attempt < 3; attempt++) {
            const delR = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${existData.id}`,
              { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } });
            if (delR.ok || delR.status === 404) break;
            context.log(`[prepare] Delete attempt ${attempt + 1} failed (${delR.status}), retrying...`);
            await new Promise(r => setTimeout(r, 1500));
          }
          context.log('[prepare] Deleted existing file:', fileName);
          await new Promise(r => setTimeout(r, 1000)); // wait for deletion to propagate
        }
      }

      const upR = await fetch(
        `${GRAPH}/sites/${siteId}/drive/root:/${folderPath}/${encodeURIComponent(fileName)}:/content`,
        { method: 'PUT', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }, body: tmplBuffer }
      );
      if (!upR.ok) throw new Error(`Upload failed (${upR.status})`);
      tempId = (await upR.json()).id;
      context.log('[prepare] Temp file created:', fileName, tempId);
    } catch(e) { return { status: 500, jsonBody: { error: e.message } }; }

    // ── Step 3: Open workbook session ─────────────────────────────────────
    let sid = null;
    try {
      const sr = await gReq('POST', `/sites/${siteId}/drive/items/${tempId}/workbook/createSession`, token, { persistChanges: true });
      if (sr.ok) { sid = (await sr.json()).id || null; context.log('[prepare] Step 3 OK — session:', sid ? 'open' : 'none'); }
      else { const t = await sr.text(); context.log('[prepare] Step 3 session failed:', sr.status, t.slice(0,200)); }
    } catch(e) { context.log('[prepare] Step 3 error:', e.message); }

    // ── Step 4: Get and clean up sheets ──────────────────────────────────
    let sheets = [], ws;
    try {
      const wr = await gReq('GET', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`, token, undefined, sid);
      sheets   = wr.ok ? (await wr.json()).value || [] : [];
      ws       = name => sheets.find(s => s.name === name);
      context.log('[prepare] Step 4 OK — sheets:', sheets.map(s=>s.name).join(', '));
    } catch(e) { context.log('[prepare] Step 4 error:', e.message); ws = () => undefined; }

    try {
      const toDelete = [];
      if (isRadon) {
        if (ws(TMPL_LAB))   toDelete.push(ws(TMPL_LAB));
        if (ws(TMPL_FHA))   toDelete.push(ws(TMPL_FHA));
      } else {
        if (ws(TMPL_RADON)) toDelete.push(ws(TMPL_RADON));
        if (!needsFHA || !fhaParams.length) { if (ws(TMPL_FHA)) toDelete.push(ws(TMPL_FHA)); }
      }
      for (const sheet of toDelete.filter(Boolean)) {
        await gReq('DELETE', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${sheet.id}`, token, undefined, sid);
      }
      context.log('[prepare] Step 4b OK — deleted', toDelete.filter(Boolean).length, 'sheets');
    } catch(e) { context.log('[prepare] Step 4b error:', e.message); }

    // ── Step 5: Rename sheets ─────────────────────────────────────────────
    try {
      const remainR   = await gReq('GET', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`, token, undefined, sid);
      const remaining = remainR.ok ? (await remainR.json()).value || [] : [];
      for (const sheet of remaining) {
        const newName = sheet.name.replace(/\s*-\s*Template\s*$/i, '').trim();
        if (newName !== sheet.name) {
          await gReq('PATCH', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${sheet.id}`, token, { name: newName }, sid);
        }
      }
      context.log('[prepare] Step 5 OK');
    } catch(e) { context.log('[prepare] Step 5 error:', e.message); }

    // ── Step 6: Fill sheets ───────────────────────────────────────────────
    context.log('[prepare] Starting Step 6 — fill sheets');
    const finalR    = await gReq('GET', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets`, token, undefined, sid);
    const finalSheets = finalR.ok ? (await finalR.json()).value || [] : [];
    const labSheet   = finalSheets.find(s => /^lab report/i.test(s.name));
    const fhaSheet   = finalSheets.find(s => /^fha/i.test(s.name));
    const radonSheet = finalSheets.find(s => /^radon/i.test(s.name));
    const specSheet  = finalSheets.find(s => /arsenic.*spec/i.test(s.name));

    // Header writer helper
    const writeHeaders = async (wsId, cells) => {
      const wsBase = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${wsId}`;
      const hHdr   = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
      const splitDT = dt => { if (!dt) return ['','']; const i = dt.indexOf(' '); return i > 0 ? [dt.slice(0,i), dt.slice(i+1)] : [dt,'']; };
      const [dc, tc] = splitDT(meta.dtCollected || meta.dateDrawn || '');
      const [dr, tr] = splitDT(meta.dtReceived  || meta.dateReceived || '');
      const vals = { labId, dc, tc, dr, tr, today: meta.dateReported || today };
      for (const [addr, key] of cells) {
        if (vals[key]) await fetch(`${wsBase}/range(address='${addr}')`,
          { method:'PATCH', headers:hHdr, body:JSON.stringify({ values:[[vals[key]]] }) }).catch(()=>{});
      }
    };

    // MODULE-LEVEL cellColors accumulator — no block-scope issues
    let cellColors = {};

    if (isRadon && radonSheet) {
      await writeHeaders(radonSheet.id, [['I7','labId'],['I8','dc'],['J8','tc'],['I9','dr'],['J9','tr'],['I10','today']]);
      const c = await fillSheet(siteId, tempId, radonSheet.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context, reportData._comments||'', '');
      Object.assign(cellColors, c);
      if (specSheet) await gReq('DELETE', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${specSheet.id}`, token, null, sid);
    } else if (isArsenicSpec && specSheet) {
      await writeHeaders(specSheet.id, [['H7','labId'],['H8','dc'],['I8','tc'],['H9','dr'],['I9','tr'],['H10','today']]);
      const c = await fillSheet(siteId, tempId, specSheet.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context, reportData._comments||'', 'A24');
      Object.assign(cellColors, c);
      const wsBaseSpec2 = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${specSheet.id}`;
      const wbHdrSpec2  = { Authorization: `Bearer ${token}`, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
      for (const [addr, val] of [['D32', authorizedBy||''],['I32', reviewDate||'']]) {
        if (val) await fetch(`${wsBaseSpec2}/range(address='${addr}')`, {
          method: 'PATCH', headers: wbHdrSpec2, body: JSON.stringify({ values: [[val]] })
        }).catch(()=>{});
      }
    } else if (labSheet) {
      if (specSheet) await gReq('DELETE', `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${specSheet.id}`, token, null, sid);
      await writeHeaders(labSheet.id, [['H7','labId'],['H8','dc'],['I8','tc'],['H9','dr'],['I9','tr'],['H10','today']]);
      // Write authorized by and review date
      const wbHdrLab = { Authorization:`Bearer ${token}`, 'workbook-session-id':sid, 'Content-Type':'application/json' };
      const wsBaseLab = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${labSheet.id}`;
      for (const [addr, val] of [['D57', authorizedBy||''],['I57', reviewDate||'']]) {
        if (val) await fetch(`${wsBaseLab}/range(address='${addr}')`, { method:'PATCH', headers:wbHdrLab, body:JSON.stringify({ values:[[val]] }) }).catch(()=>{});
      }
      const c = await fillSheet(siteId, tempId, labSheet.id, params, meta, labId, authorizedBy, reviewDate, today, token, sid, context, reportData._comments||'', 'A48');
      Object.assign(cellColors, c);
    }

    if (fhaSheet && needsFHA && fhaParams.length) {
      await writeHeaders(fhaSheet.id, [['H7','labId'],['H8','dc'],['I8','tc'],['H9','dr'],['I9','tr'],['H10','today']]);
      const c = await fillSheet(siteId, tempId, fhaSheet.id, fhaParams, meta, labId, authorizedBy, reviewDate, today, token, sid, context, reportData._comments||'', 'A27');
      Object.assign(cellColors, c);
    }

    // ── Step 7: Close write session ─────────────────────────────────────
    if (sid) {
      await gReq('POST', `/sites/${siteId}/drive/items/${tempId}/workbook/closeSession`, token, {}, sid).catch(()=>{});
    }

    // ── Step 8: Re-open read-only session and read actual cell fill colors ──
    // Single range read per sheet — no per-cell API calls to avoid timeout
    try {
      await new Promise(r => setTimeout(r, 1500));

      const roSr = await gReq('POST', `/sites/${siteId}/drive/items/${tempId}/workbook/createSession`,
        token, { persistChanges: false });
      if (roSr.ok) {
        const roSid = (await roSr.json()).id;
        const roHdr = { Authorization: `Bearer ${token}`, 'workbook-session-id': roSid };

        const sheetsToRead = [labSheet, fhaSheet, radonSheet, specSheet].filter(Boolean);
        for (const sheet of sheetsToRead) {
          const wsBase = `${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${sheet.id}`;

          // Read values + entire used range in one call
          const rr2 = await fetch(`${wsBase}/usedRange?$select=values,columnCount`, { headers: roHdr });
          if (!rr2.ok) continue;
          const { values: rows2, columnCount: nc2 } = await rr2.json();

          let hdrRow2 = -1, colResult2 = -1;
          for (let r = 0; r < rows2.length; r++) {
            const rl = (rows2[r] || []).map(c => normalizeCell(c));
            if (rl.some(c => c.includes('your result') || c === 'result')) {
              hdrRow2    = r;
              colResult2 = rl.findIndex(c => c.includes('your result') || c === 'result');
              break;
            }
          }
          if (hdrRow2 < 0 || colResult2 < 1) continue;
          const colorCol2 = colResult2 - 1;

          // Read ALL cells in the color indicator column in one range call
          const lastRow   = rows2.length;
          const rangeAddr = `${colLetter(colorCol2)}${hdrRow2 + 2}:${colLetter(colorCol2)}${lastRow}`;
          const fillRes   = await fetch(`${wsBase}/range(address='${rangeAddr}')/format/fill`, { headers: roHdr });
          if (!fillRes.ok) continue;

          // Graph returns one fill object for a range only if all cells share the same fill.
          // If not uniform, we need cell-by-cell reads — but try the range first.
          const fillData  = await fillRes.json();
          const rangeColor = fillData.color;

          if (rangeColor && rangeColor !== '#FFFFFF' && rangeColor !== '#ffffff') {
            // All cells same color — apply to all matched params
            for (let r = hdrRow2 + 1; r < rows2.length; r++) {
              const paramName    = String((rows2[r] || [])[0] || '').trim();
              const matchedParam = [...params, ...fhaParams].find(p => {
                const pn = normalizeCell(p.name), rn = normalizeCell(paramName);
                return pn === rn || pn.startsWith(rn) || rn.startsWith(pn);
              });
              if (matchedParam) cellColors[matchedParam.name] = rangeColor.startsWith('#') ? rangeColor : `#${rangeColor}`;
            }
          } else {
            // Mixed colors — read each param cell individually via batch
            const batchReqs = [];
            const batchMap  = [];
            for (let r = hdrRow2 + 1; r < rows2.length; r++) {
              const paramName    = String((rows2[r] || [])[0] || '').trim();
              const matchedParam = [...params, ...fhaParams].find(p => {
                const pn = normalizeCell(p.name), rn = normalizeCell(paramName);
                return pn === rn || pn.startsWith(rn) || rn.startsWith(pn);
              });
              if (!matchedParam) continue;
              const addr = `${colLetter(colorCol2)}${r + 1}`;
              batchReqs.push({ method: 'GET', url: `/sites/${siteId}/drive/items/${tempId}/workbook/worksheets/${sheet.id}/range(address='${addr}')/format/fill` });
              batchMap.push(matchedParam.name);
            }
            // Batch GET — Graph batch supports GET requests
            if (batchReqs.length) {
              const batchRes = await fetch(`${GRAPH}/$batch`, {
                method: 'POST',
                headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ requests: batchReqs.map((r, i) => ({ id: String(i+1), method: r.method, url: r.url, headers: { 'workbook-session-id': roSid } })) }),
              });
              if (batchRes.ok) {
                const batchData = await batchRes.json();
                (batchData.responses || []).forEach(resp => {
                  const idx   = parseInt(resp.id) - 1;
                  const color = resp.body?.color;
                  if (color && color !== '#FFFFFF' && color !== '#ffffff' && batchMap[idx]) {
                    cellColors[batchMap[idx]] = color.startsWith('#') ? color : `#${color}`;
                  }
                });
              }
            }
          }
        }

        await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}/workbook/closeSession`,
          { method: 'POST', headers: roHdr }).catch(() => {});
        context.log(`[prepare] Read ${Object.keys(cellColors).length} actual cell colors from Excel`);
      }
    } catch(e) {
      context.log('[prepare] Color read-back failed (non-fatal):', e.message);
    }

    context.log(`[prepare] Done — tempId: ${tempId}, colors returned: ${Object.keys(cellColors).length}`);

    return {
      status:   200,
      jsonBody: { success: true, tempId, fileName, cellColors },
    };
    } catch(fatalErr) {
      context.log('[prepare] FATAL:', fatalErr.message, fatalErr.stack);
      return { status: 500, jsonBody: { error: fatalErr.message, stack: fatalErr.stack?.slice(0,500) } };
    }
  },
});
