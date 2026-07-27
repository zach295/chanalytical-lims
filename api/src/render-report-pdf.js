/**
 * render-report-pdf.js — Azure version (Excel template approach)
 * 1. Downloads Report Templates.xlsx from SharePoint
 * 2. Copies the workbook to a temp file
 * 3. Writes sample data into specific cells via Graph API Excel endpoints
 * 4. Exports as PDF via Graph (/content?format=pdf)
 * 5. Deletes the temp workbook
 * 6. Returns base64 PDF
 *
 * Cell map (Lab Report - Template):
 *   B7=customer  B8=billing address  B9=email
 *   H7=Lab ID    H8=date collected   I8=time collected
 *                H9=date received    I9=time received
 *   B11=street   H10=date reported
 *   B12=city/state/zip
 *   D17:D43 = results, I=prep DT, J=analysis DT
 *   D55=authorized by   I56=review date
 *
 * Radon Lab Report - Template: same header, result D18, analysis J18
 * FHA Lab Report - Template: same layout as Lab Report
 */
const { app }    = require('@azure/functions');
const { getToken } = require('../shared/graph');

const GRAPH   = 'https://graph.microsoft.com/v1.0';
const SITE_ID = process.env.SP_SITE_ID;

const TEMPLATE_PATH = process.env.SP_REPORT_TEMPLATES ||
  '/sites/Laboratory/Shared Documents/Documents/Report Templates.xlsx';

const PARAM_ROWS = {
  'Chloride, Total':17,'Fluoride':18,'Nitrite-Nitrogen, Total':19,'Nitrite':19,
  'Nitrate-Nitrogen, Total':20,'Nitrate':20,'Arsenic, Total':21,'Lead, Total':22,
  'Uranium, Total':23,'Copper, Total':24,'Iron, Total':25,'Manganese, Total':26,
  'Sodium, Total':27,'Hardness by calculation':28,'Total Hardness':28,
  'Calcium, Total':29,'Magnesium, Total':30,'Antimony, Total':31,'Antimony':31,
  'Cadmium, Total':32,'Chromium, Total':33,'Chromium':33,'Cobalt':34,
  'pH Electrometric':35,'pH':35,'Alkalinity':36,'Sulfate':37,'Tannins':38,
  'Total Dissolved Solids (TDS)':39,'Bromide':40,'Total Coliform':42,'E. Coli':43,
};

async function gGet(path, token) {
  const res = await fetch(`${GRAPH}${path}`, { headers:{ Authorization:`Bearer ${token}` } });
  if (!res.ok) throw new Error(`GET ${path.slice(0,80)} → ${res.status}`);
  return res.json();
}

async function gPost(path, body, token) {
  const res = await fetch(`${GRAPH}${path}`, {
    method:'POST', headers:{ Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
    body: JSON.stringify(body),
  });
  return res.json().catch(()=>({}));
}

async function gPatch(path, body, token) {
  const res = await fetch(`${GRAPH}${path}`, {
    method:'PATCH', headers:{ Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`PATCH ${path.slice(0,80)} → ${res.status}: ${(await res.text()).slice(0,100)}`);
  return res.json().catch(()=>({}));
}

async function gDelete(path, token) {
  await fetch(`${GRAPH}${path}`, { method:'DELETE', headers:{ Authorization:`Bearer ${token}` } });
}

function toDrivePath(p) {
  const i = p.indexOf('Shared Documents/');
  return i>=0 ? p.slice(i+17) : p.replace(/^\/+/,'');
}

function fmtDate(d) {
  if (!d) return '';
  const s = String(d).trim();
  if (/^\d{2}-\d{2}-\d{2}$/.test(s)) return s.replace(/-/g,'/');
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) { const [y,m,dd]=s.split('-'); return `${m}/${dd}/${y.slice(-2)}`; }
  return s;
}

async function getItemId(drivePath, token) {
  const d = await gGet(`/sites/${SITE_ID}/drive/root:/${drivePath}?$select=id`, token);
  return d.id;
}

async function copyToTemp(srcId, tempName, token) {
  const root = await gGet(`/sites/${SITE_ID}/drive/root?$select=id`, token);
  await gPost(`/sites/${SITE_ID}/drive/items/${srcId}/copy`,
    { parentReference:{ id:root.id }, name:tempName }, token);
  await new Promise(r=>setTimeout(r,3000));
  for (let i=0;i<10;i++) {
    try {
      const f = await gGet(`/sites/${SITE_ID}/drive/root:/${encodeURIComponent(tempName)}?$select=id`, token);
      if (f.id) return f.id;
    } catch {}
    await new Promise(r=>setTimeout(r,2000));
  }
  throw new Error('Temp workbook copy timed out');
}

async function writeCell(fileId, sheet, addr, value, token) {
  const s = encodeURIComponent(sheet);
  await gPatch(`/sites/${SITE_ID}/drive/items/${fileId}/workbook/worksheets/${s}/range(address='${addr}')`,
    { values:[[value??'']] }, token);
}

async function writeCells(fileId, sheet, cells, token) {
  for (let i=0;i<cells.length;i+=4) {
    await Promise.all(cells.slice(i,i+4).map(({address,value}) =>
      writeCell(fileId,sheet,address,value,token).catch(e=>console.warn(`writeCell ${address}:`,e.message))
    ));
  }
}

async function exportPDF(fileId, token) {
  const res = await fetch(`${GRAPH}/sites/${SITE_ID}/drive/items/${fileId}/content?format=pdf`,
    { headers:{ Authorization:`Bearer ${token}` } });
  if (!res.ok) throw new Error(`PDF export → ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

async function buildReport(templateId, sheetName, meta, paramList, resultsMap, authorizedBy, reviewDate, today, token) {
  const tempName = `_rpt_${Date.now()}_${Math.random().toString(36).slice(2)}.xlsx`;
  const tempId   = await copyToTemp(templateId, tempName, token);
  try {
    const cells = [
      { address:'B7',  value: meta.customer    ||'' },
      { address:'B8',  value: [meta.location,meta.city,meta.state,meta.zip].filter(Boolean).join(', ') },
      { address:'B9',  value: meta.email       ||'' },
      { address:'H7',  value: meta.labId       ||'' },
      { address:'B11', value: meta.location    ||'' },
      { address:'B12', value: [meta.city,meta.state,meta.zip].filter(Boolean).join(', ') },
      { address:'H10', value: fmtDate(today)       },
      { address:'D55', value: authorizedBy     ||'' },
      { address:'I56', value: fmtDate(reviewDate||today) },
    ];
    if (meta.dateDrawn)    { cells.push({address:'H8',value:fmtDate(meta.dateDrawn)}); cells.push({address:'I8',value:meta.timeDrawn||''}); }
    if (meta.dateReceived) { cells.push({address:'H9',value:fmtDate(meta.dateReceived)}); cells.push({address:'I9',value:meta.timeReceived||''}); }

    for (const p of paramList) {
      const row = PARAM_ROWS[p.name];
      if (!row) continue;
      const res = resultsMap[p.name]||p;
      if (res.value!==undefined&&res.value!=='') cells.push({address:`D${row}`,value:String(res.value)});
      if (res.prepDT)            cells.push({address:`I${row}`,value:String(res.prepDT)});
      if (res.analDT||res.time)  cells.push({address:`J${row}`,value:String(res.analDT||res.time)});
    }

    await writeCells(tempId, sheetName, cells, token);
    const pdf = await exportPDF(tempId, token);
    return pdf;
  } finally {
    await gDelete(`/sites/${SITE_ID}/drive/items/${tempId}`, token).catch(()=>{});
  }
}

app.http('render-report-pdf', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { reportData, authorizedBy, reviewDate } = await request.json().catch(()=>({}));
      if (!reportData) return { status:400, jsonBody:{ error:'reportData required' } };

      const token      = await getToken();
      const meta       = reportData.meta       || {};
      const paramRows  = reportData.paramRows  || [];
      const fhaRows    = reportData.fhaRows    || [];
      const resultsMap = reportData.resultsMap || {};
      const isRadon    = reportData.isRadon    || false;
      const needsFHA   = reportData.needsFHA   || false;
      const today      = reportData.today      || new Date().toLocaleDateString('en-US',{month:'2-digit',day:'2-digit',year:'2-digit'});

      const templateDrivePath = toDrivePath(TEMPLATE_PATH);
      const templateId        = await getItemId(templateDrivePath, token);
      context.log(`[render] Template ID: ${templateId}`);

      const pdfPages = [];

      if (isRadon) {
        // Radon report — single row D18/J18
        const tempName = `_rpt_${Date.now()}.xlsx`;
        const tempId   = await copyToTemp(templateId, tempName, token);
        try {
          const radonRes = resultsMap['Radon Water']||{};
          const rawVal   = parseFloat(radonRes.value)||0;
          const display  = !radonRes.value?'':rawVal<100?'<100':String(Math.round(rawVal/100)*100);
          const cells = [
            {address:'B7', value:meta.customer||''},{address:'B8',value:[meta.location,meta.city,meta.state,meta.zip].filter(Boolean).join(', ')},
            {address:'B9',value:meta.email||''},{address:'H7',value:meta.labId||''},
            {address:'B11',value:meta.location||''},{address:'B12',value:[meta.city,meta.state,meta.zip].filter(Boolean).join(', ')},
            {address:'H10',value:fmtDate(today)},{address:'D55',value:authorizedBy||''},{address:'I56',value:fmtDate(reviewDate||today)},
            {address:'D18',value:display},{address:'J18',value:radonRes.analDT||radonRes.time||''},
          ];
          if (meta.dateDrawn)    { cells.push({address:'H8',value:fmtDate(meta.dateDrawn)});    cells.push({address:'I8',value:meta.timeDrawn||''}); }
          if (meta.dateReceived) { cells.push({address:'H9',value:fmtDate(meta.dateReceived)}); cells.push({address:'I9',value:meta.timeReceived||''}); }
          await writeCells(tempId,'Radon Lab Report - Template',cells,token);
          const pdf = await exportPDF(tempId,token);
          pdfPages.push(pdf.toString('base64'));
          context.log(`[render] Radon PDF: ${pdf.length} bytes`);
        } finally {
          await gDelete(`/sites/${SITE_ID}/drive/items/${tempId}`,token).catch(()=>{});
        }
      } else {
        // COA report
        const pdf = await buildReport(templateId,'Lab Report - Template',meta,paramRows,resultsMap,authorizedBy,reviewDate,today,token);
        pdfPages.push(pdf.toString('base64'));
        context.log(`[render] COA PDF: ${pdf.length} bytes`);

        // FHA page if needed
        if (needsFHA && fhaRows.length) {
          try {
            const fhaPdf = await buildReport(templateId,'FHA Lab Report - Template',meta,fhaRows,resultsMap,authorizedBy,reviewDate,today,token);
            pdfPages.push(fhaPdf.toString('base64'));
            context.log(`[render] FHA PDF: ${fhaPdf.length} bytes`);
          } catch(e) { context.log('[render] FHA failed (non-fatal):',e.message); }
        }
      }

      return { status:200, jsonBody:{ success:true, pdfPages, pageCount:pdfPages.length } };

    } catch (err) {
      context.log('[render-report-pdf] Error:', err.message);
      return { status:500, jsonBody:{ error:err.message } };
    }
  },
});
