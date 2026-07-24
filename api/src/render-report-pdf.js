/**
 * render-report-pdf.js — Azure version
 * Builds COA/RW report PDF using pdf-lib.
 * No Google Sheets dependency — renders directly from report data.
 *
 * POST { reportData, authorizedBy, reviewDate }
 * Returns { success, pdfPages: [base64], pageCount }
 */
const { app }    = require('@azure/functions');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

// ── Color helpers ──────────────────────────────────────────────────────────────
const COLORS = {
  green: rgb(0,    0.73, 0.26),
  blue:  rgb(0.13, 0.47, 0.71),
  red:   rgb(0.90, 0.22, 0.21),
  navy:  rgb(0,    0.227,0.361),
  white: rgb(1,    1,    1),
  black: rgb(0,    0,    0),
  light: rgb(0.95, 0.95, 0.95),
  grey:  rgb(0.4,  0.4,  0.4),
  border:rgb(0.8,  0.8,  0.8),
};

// ── Column letter helper ───────────────────────────────────────────────────────
function col(idx) {
  let s = ''; idx++;
  while (idx > 0) { s = String.fromCharCode(64 + (idx - 1) % 26 + 1) + s; idx = Math.floor((idx - 1) / 26); }
  return s;
}

// ── Build a single COA/RW page ─────────────────────────────────────────────────
async function buildPage(doc, fonts, data, paramRows, isRadon, pageType) {
  const page   = doc.addPage([612, 792]); // portrait letter
  const { width, height } = page.getSize();
  const { bold, regular } = fonts;

  let y = height - 30;
  const LEFT  = 36;
  const RIGHT = width - 36;
  const COL_W = width - 72;

  // ── Header ──────────────────────────────────────────────────────────────────
  page.drawRectangle({ x:LEFT, y:y-42, width:COL_W, height:44, color:COLORS.navy });
  page.drawText('Chanalytical Laboratories, Inc.', { x:LEFT+8, y:y-16, size:13, font:bold, color:COLORS.white });
  page.drawText('347 Main St., Unit 1B  ·  Gorham, ME 04038  ·  207-747-1815  ·  Labs@chanalytical.com',
    { x:LEFT+8, y:y-30, size:7, font:regular, color:rgb(0.8,0.9,1) });
  y -= 56;

  // ── Title ───────────────────────────────────────────────────────────────────
  const title = isRadon ? 'Certificate of Analysis — Radon Water' :
                pageType === 'FHA' ? 'Certificate of Analysis — FHA' :
                'Certificate of Analysis';
  page.drawText(title, { x:LEFT, y, size:14, font:bold, color:COLORS.navy });
  y -= 18;

  // ── Sample info grid ─────────────────────────────────────────────────────────
  const meta = data.meta || {};
  const drawInfo = (label, value, x, w, yPos) => {
    page.drawText(label + ':', { x, y:yPos+2, size:7, font:regular, color:COLORS.grey });
    page.drawText(String(value||''), { x:x+2, y:yPos-9, size:9, font:bold, color:COLORS.black });
    page.drawLine({ start:{x,y:yPos-12}, end:{x:x+w,y:yPos-12}, thickness:0.3, color:COLORS.border });
  };

  const halfW = (COL_W - 8) / 2;
  drawInfo('Attention / Client',   meta.customer    || '', LEFT,        halfW, y); y -= 28;
  drawInfo('Lab ID Number',        meta.labId       || '', LEFT,        halfW, y);
  drawInfo('Date Reported',        data.today       || '', LEFT+halfW+8,halfW, y); y -= 28;
  drawInfo('Date/Time Collected',  meta.dtCollected || '', LEFT,        halfW, y);
  drawInfo('Date/Time Received',   meta.dtReceived  || '', LEFT+halfW+8,halfW, y); y -= 28;
  drawInfo('Sample Location',      [meta.location, meta.city, meta.state, meta.zip].filter(Boolean).join(', '),
    LEFT, COL_W, y); y -= 28;
  drawInfo('Authorized By',        data.authorizedBy|| '', LEFT,        halfW, y);
  drawInfo('Review Date',          data.reviewDate  || '', LEFT+halfW+8,halfW, y); y -= 20;

  // ── Legend ──────────────────────────────────────────────────────────────────
  const legendY = y;
  const boxes = isRadon
    ? [['At/below Maine MEG (4,000 pCi/L)', COLORS.green], ['Above Maine MEG', COLORS.blue]]
    : [['Meets EPA Limits', COLORS.green], ['See Notation', COLORS.blue], ['Exceeds EPA Limits', COLORS.red]];
  let lx = LEFT;
  boxes.forEach(([label, color]) => {
    page.drawRectangle({ x:lx, y:legendY-10, width:14, height:10, color });
    page.drawText(label, { x:lx+17, y:legendY-9, size:7, font:regular, color:COLORS.black });
    lx += label.length * 4.5 + 26;
  });
  y -= 20;

  // ── Parameter table ──────────────────────────────────────────────────────────
  const ROW_H     = 14;
  const COL_WIDTHS = [170, 58, 52, 38, 70, 70, 72]; // param, result, epa, unit, method, prep, anal
  const HEADERS    = ['Parameter', 'Your Result', 'EPA Limit', 'Unit', 'Method', 'Prep Date/Time', 'Analysis Date/Time'];
  const col_x      = [LEFT];
  COL_WIDTHS.forEach((w,i) => col_x.push(col_x[i]+w));

  // Table header
  page.drawRectangle({ x:LEFT, y:y-ROW_H, width:COL_W, height:ROW_H, color:COLORS.navy });
  HEADERS.forEach((h,i) => {
    page.drawText(h, { x:col_x[i]+2, y:y-ROW_H+4, size:6.5, font:bold, color:COLORS.white });
  });
  y -= ROW_H;

  // Table rows
  if (isRadon) {
    const radonRes = (data.resultsMap || {})['Radon Water'] || {};
    const rawVal   = parseFloat(radonRes.value) || 0;
    const display  = !radonRes.value ? '' : rawVal < 100 ? '<100' : String(Math.round(rawVal/100)*100);
    const bg       = rawVal > 4000 ? COLORS.blue : COLORS.green;

    page.drawRectangle({ x:LEFT, y:y-ROW_H, width:COL_W, height:ROW_H, color:COLORS.light });
    page.drawLine({ start:{x:LEFT,y:y-ROW_H}, end:{x:RIGHT,y:y-ROW_H}, thickness:0.3, color:COLORS.border });
    page.drawText('Radon Water', { x:col_x[0]+2, y:y-ROW_H+4, size:7, font:regular });
    page.drawRectangle({ x:col_x[1], y:y-ROW_H+1, width:COL_WIDTHS[1]-2, height:ROW_H-2, color:bg });
    page.drawText(display, { x:col_x[1]+2, y:y-ROW_H+4, size:7, font:bold, color:COLORS.white });
    page.drawText('4,000',   { x:col_x[2]+2, y:y-ROW_H+4, size:7, font:regular });
    page.drawText('pCi/l',   { x:col_x[3]+2, y:y-ROW_H+4, size:7, font:regular });
    page.drawText(radonRes.time||'', { x:col_x[6]+2, y:y-ROW_H+4, size:6, font:regular });
    y -= ROW_H;
  } else {
    paramRows.forEach((p, idx) => {
      if (y < 80) return; // page overflow guard
      const bg = idx % 2 === 0 ? COLORS.light : rgb(1,1,1);
      page.drawRectangle({ x:LEFT, y:y-ROW_H, width:COL_W, height:ROW_H, color:bg });
      page.drawLine({ start:{x:LEFT,y:y-ROW_H}, end:{x:RIGHT,y:y-ROW_H}, thickness:0.3, color:COLORS.border });

      page.drawText(p.name  ||'', { x:col_x[0]+2, y:y-ROW_H+4, size:7, font:regular, maxWidth:COL_WIDTHS[0]-4 });
      // Result cell with color
      if (p.value) {
        const resultBg = COLORS[p.color] || null;
        if (resultBg) page.drawRectangle({ x:col_x[1], y:y-ROW_H+1, width:COL_WIDTHS[1]-2, height:ROW_H-2, color:resultBg });
        const textColor = (p.color === 'blue' || p.color === 'red') ? COLORS.white : COLORS.black;
        page.drawText(String(p.value), { x:col_x[1]+2, y:y-ROW_H+4, size:7, font:bold, color:textColor });
      }
      page.drawText(String(p.epa    ||''), { x:col_x[2]+2, y:y-ROW_H+4, size:7, font:regular });
      page.drawText(String(p.unit   ||''), { x:col_x[3]+2, y:y-ROW_H+4, size:7, font:regular });
      page.drawText(String(p.method ||''), { x:col_x[4]+2, y:y-ROW_H+4, size:6, font:regular });
      page.drawText(String(p.prepDT ||''), { x:col_x[5]+2, y:y-ROW_H+4, size:6, font:regular });
      page.drawText(String(p.analDT||p.time||''), { x:col_x[6]+2, y:y-ROW_H+4, size:6, font:regular });
      y -= ROW_H;
    });
  }

  // Column borders
  COL_WIDTHS.forEach((_,i) => {
    if (i === 0) return;
    page.drawLine({ start:{x:col_x[i],y:legendY-20}, end:{x:col_x[i],y:y}, thickness:0.3, color:COLORS.border });
  });

  // ── Comments ──────────────────────────────────────────────────────────────────
  y -= 12;
  if (data._comments) {
    page.drawText('Comments:', { x:LEFT, y, size:8, font:bold });
    y -= 12;
    page.drawText(String(data._comments), { x:LEFT, y, size:8, font:regular, maxWidth:COL_W });
    y -= 16;
  }

  // ── Notations ────────────────────────────────────────────────────────────────
  y -= 8;
  const notations = isRadon ? [
    'Maine\'s current Maximum Exposure Guideline (MEG) for radon in well water is 4,000 pCi/L.',
    'Radon in water can be reduced by aeration or carbon filtration. Work should be done by a mitigation contractor registered with the State of Maine.',
    'Maine Disclose: This lab meets EPA requirements for radon testing. The State of Maine Radon Registration Act requires this laboratory to report test results, zip codes and street addresses.',
  ] : [
    'Notation 1: The Maximum Contaminant Level (MCL) is a health-based guideline set by the Maine Center for Disease Control and Prevention (MECDCP).',
    'Notation 2: The Secondary Maximum Contaminant Level (SMCL) is set by the USEPA through the National Secondary Drinking Water Regulations.',
    'Notation 3: Total coliform bacteria are used as indicator organisms for the presence of pathogens.',
    'This report shall not be reproduced, except in full, without written permission from Chanalytical Laboratories Inc.',
    'If you have any questions regarding your report please call 207-747-1815.',
  ];

  if (y > 60) {
    page.drawLine({ start:{x:LEFT,y}, end:{x:RIGHT,y}, thickness:0.5, color:COLORS.border });
    y -= 10;
    notations.forEach(note => {
      if (y < 36) return;
      page.drawText(note, { x:LEFT, y, size:6.5, font:regular, maxWidth:COL_W, color:COLORS.grey });
      y -= 10;
    });
  }
}

// ── Main handler ───────────────────────────────────────────────────────────────
app.http('render-report-pdf', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { reportData, authorizedBy, reviewDate } = await request.json().catch(() => ({}));
      if (!reportData) return { status: 400, jsonBody: { error: 'reportData required' } };

      reportData.authorizedBy = authorizedBy || '';
      reportData.reviewDate   = reviewDate   || '';

      const doc   = await PDFDocument.create();
      const bold  = await doc.embedFont(StandardFonts.HelveticaBold);
      const regular = await doc.embedFont(StandardFonts.Helvetica);
      const fonts = { bold, regular };

      // Build results map from paramRows
      const resultsMap = reportData.resultsMap || {};
      if (!Object.keys(resultsMap).length) {
        (reportData.paramRows || []).forEach(p => { resultsMap[p.name] = p; });
        (reportData.fhaRows   || []).forEach(p => { if (!resultsMap[p.name]) resultsMap[p.name] = p; });
        reportData.resultsMap = resultsMap;
      }

      // Main COA/RW page
      await buildPage(doc, fonts, reportData, reportData.paramRows || [], reportData.isRadon, 'main');

      // FHA page if needed
      if (reportData.needsFHA && (reportData.fhaRows||[]).length) {
        await buildPage(doc, fonts, reportData, reportData.fhaRows, false, 'FHA');
      }

      const pdfBytes = await doc.save();
      const b64      = Buffer.from(pdfBytes).toString('base64');

      return {
        status:   200,
        jsonBody: { success: true, pdfPages: [b64], pageCount: doc.getPageCount() },
      };

    } catch (err) {
      context.log('[render-report-pdf] Error:', err.message);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
