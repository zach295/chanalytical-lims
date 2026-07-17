/**
 * approve-scan — assigns Lab ID, writes to Archived Intake, deletes from Review Queue
 */
const { app } = require('@azure/functions');
const { createItem, updateItem, deleteItem, findItem, listItems, LISTS } = require('../shared/graph');

// ── ET Time helpers ────────────────────────────────────────────────────────────
const TZ = 'America/New_York';
function etParts(d) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: TZ, year:'numeric', month:'2-digit', day:'2-digit',
    hour:'2-digit', minute:'2-digit', second:'2-digit', hour12: false,
  }).formatToParts(d || new Date());
  const get = t => parts.find(p => p.type === t)?.value || '';
  return { year:get('year'), month:get('month'), day:get('day'),
           hour:get('hour')==='24'?'00':get('hour'), minute:get('minute'), second:get('second') };
}
function todayDisplayET() { const p=etParts(); return `${p.month}-${p.day}-${p.year.slice(-2)}`; }
function nowTimeET() { return (new Date()).toLocaleTimeString('en-US',{timeZone:TZ,hour:'2-digit',minute:'2-digit',hour12:true}); }
function mmddyyET() { const p=etParts(); return `${p.month}${p.day}${p.year.slice(-2)}`; }
function nextWorkdayET() {
  const d = new Date();
  const dow = new Date(d.toLocaleString('en-US',{timeZone:TZ})).getDay();
  const add = dow===5?3:dow===6?2:dow===0?1:1;
  const next = new Date(d.getTime()+add*86400000);
  const p = etParts(next);
  return `${p.month}-${p.day}-${p.year.slice(-2)}`;
}

// ── Suffix / Package maps ──────────────────────────────────────────────────────
const SUFFIX_MAP = {
  'Basic Safety (FHA)':'BS','Basic Safety':'BS','Standard Safety':'SS',
  'Expanded Safety (Mortgage Test)':'EXP','WW - Expanded Safety':'WW',
  'Comprehensive':'COMP','Pro Plus':'PP','Radon Water':'RW',
  'AIO FHA':'AIOFHA','AIO Portability':'PORT',
  'Rejected - Timeout':'REJ','Rejected - Chlorine':'REJ','Rejected - Other':'REJ',
  'Alkalinity':'ALK','Arsenic, Total':'AS','Bacteria':'BAC',
  'Cadmium, Total':'CD','Calcium, Total':'CA','Chloride, Total':'CL',
  'Chromium':'CR','Cobalt':'CO','Copper, Total':'CU','Fluoride':'FL',
  'Hardness, Total':'HRD','Iron, Total':'FE','Lead, Total':'PB',
  'Magnesium, Total':'MG','Manganese, Total':'MN','Nitrate':'NO3',
  'Nitrite':'NO2','pH':'PH','Sodium, Total':'NA','Sulfate':'SO4',
  'Tannins':'TAN','Total Dissolved Solids (TDS)':'TDS','Uranium, Total':'U',
};
const PACKAGE_SET = new Set([
  'Basic Safety (FHA)','Basic Safety','Standard Safety',
  'Expanded Safety (Mortgage Test)','WW - Expanded Safety',
  'Comprehensive','Pro Plus','AIO FHA','AIO Portability',
]);

function normalizeTest(t) {
  const map = {
    'basic safety (fha)':'Basic Safety (FHA)','basic safety':'Basic Safety',
    'standard safety':'Standard Safety','expanded safety (mortgage test)':'Expanded Safety (Mortgage Test)',
    'ww - expanded safety':'WW - Expanded Safety','comprehensive':'Comprehensive',
    'pro plus':'Pro Plus','radon water':'Radon Water','aio fha':'AIO FHA',
    'aio portability':'AIO Portability','portability':'AIO Portability',
    'rejected - timeout':'Rejected - Timeout','rejected - chlorine':'Rejected - Chlorine',
    'rejected - other':'Rejected - Other',
  };
  return map[t.toLowerCase()] || t;
}

function getAbbrev(name) {
  return name.split(/\s+/).filter(w=>w.length>2&&!/^(inc|llc|ltd|corp|co)\.?$/i.test(w))
    .map(w=>w[0].toUpperCase()).join('').slice(0,4) || name.slice(0,3).toUpperCase();
}

app.http('approve-scan', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const {
        fileId, reviewQueueRow, reviewedBy,
        customer, dateDrawn, timeDrawn, receivedDate, receivedTime,
        location, city, state, zip, tests, hasRadon, notes, email,
      } = await request.json();

      if (!fileId || !tests?.length)
        return { status:400, body: JSON.stringify({ error:'fileId and tests required' }) };

      const n = new Date();
      const etDow = new Date(n.toLocaleString('en-US',{timeZone:TZ})).getDay();
      let labDateOffset = 0;
      if (etDow===6) labDateOffset=-1;
      if (etDow===0) labDateOffset=-2;
      const labDateET = new Date(n.getTime()+labDateOffset*86400000);
      const mmddyy = mmddyyET();
      const reportDateStr = nextWorkdayET();
      const tdStr = todayDisplayET();
      const tmStr = nowTimeET();
      const ts = new Date().toISOString();

      // ── Get next sequence number for today ───────────────────────────────────
      const todayItems = await listItems(LISTS.ACCESSION_LOG, {
        filter: `fields/BaseId ge '${mmddyy}-001' and fields/BaseId le '${mmddyy}-999'`,
        top: 500,
      });
      const used = new Set(todayItems.map(r => r.BaseId));
      let seq = 1;
      while (used.has(`${mmddyy}-${String(seq).padStart(3,'0')}`)) seq++;
      const baseId = `${mmddyy}-${String(seq).padStart(3,'0')}`;

      const normalizedTests = tests.map(normalizeTest);
      const selectedPackage = normalizedTests.find(t => PACKAGE_SET.has(t)) || null;
      const selectedElements = normalizedTests.filter(t => !PACKAGE_SET.has(t) && t !== 'Radon Water');
      const hasRadon = hasRadon || normalizedTests.includes('Radon Water');

      if (!selectedPackage && !selectedElements.length)
        return { status:400, body: JSON.stringify({ error:'No valid tests selected' }) };

      const fmt = d => {
        if (!d) return '';
        if (/^\d{4}-\d{2}-\d{2}$/.test(d)) {
          const [y,m,dd] = d.split('-');
          return `${m}-${dd}-${y.slice(-2)}`;
        }
        return d;
      };

      const labItems = [];

      // Non-radon items
      if (selectedPackage || selectedElements.length) {
        const suffixParts = [];
        if (selectedPackage) suffixParts.push(SUFFIX_MAP[selectedPackage] || 'GEN');
        for (const el of selectedElements) suffixParts.push(SUFFIX_MAP[el] || el.substring(0,3).toUpperCase());
        const suffix = suffixParts.join('-');
        const fullId = `${baseId} ${suffix}`;
        const allTestNames = [selectedPackage, ...selectedElements].filter(Boolean).join(', ');
        labItems.push({ baseId, fullId, suffix, isRadon:false, coaTest:allTestNames });

        // Handle rejections
        const isRejection = normalizedTests.some(t => t.startsWith('Rejected'));
        if (isRejection) {
          const rejType = normalizedTests.find(t => t.startsWith('Rejected')) || 'Rejected - Other';
          const rejReason = notes || `${rejType} — approved via Review Queue`;
          await createItem(LISTS.REJECTED, {
            Title: fullId, LabId: fullId, RejectionType: rejType,
            Reason: rejReason, RejectedBy: reviewedBy || 'Lab Staff', Timestamp: ts,
          });
        }
      }

      // Radon item
      if (hasRadon) {
        labItems.push({ baseId, fullId:`${baseId} RW`, suffix:'RW', isRadon:true, coaTest:'Radon Water' });
      }

      // ── Write Accession Log ──────────────────────────────────────────────────
      for (const item of labItems) {
        await createItem(LISTS.ACCESSION_LOG, {
          Title: item.fullId, BaseId: item.baseId, FullId: item.fullId,
          CoaTest: item.coaTest, Suffix: item.suffix, Timestamp: ts,
        });
      }

      // ── Write Archived Intake ────────────────────────────────────────────────
      for (const item of labItems) {
        await createItem(LISTS.ARCHIVED_INTAKE, {
          Title:        item.fullId,
          Timestamp:    ts,
          FullId:       item.fullId,
          CoaTest:      item.coaTest,
          Customer:     customer || '',
          DateDrawn:    fmt(dateDrawn) || '',
          TimeDrawn:    timeDrawn || '',
          ReceivedDate: fmt(receivedDate) || tdStr,
          ReceivedTime: receivedTime || tmStr,
          Location:     location || '',
          City:         city || '',
          State:        state || 'ME',
          Zip:          zip ? String(zip).padStart(5,'0') : '',
          ReviewedBy:   reviewedBy || 'Lab Staff',
          Notes:        notes || '',
          ReportStatus: 'Pending',
        });
      }

      // ── Add new client if not already on list ────────────────────────────────
      if (customer) {
        const existingClient = await findItem(LISTS.CLIENTS, 'ClientName', customer).catch(() => null);
        if (!existingClient) {
          await createItem(LISTS.CLIENTS, {
            Title: customer, ClientName: customer,
            Abbrev: getAbbrev(customer), Email: email || '', Active: 'TRUE',
          });
        }
      }

      // ── Delete from Review Queue ─────────────────────────────────────────────
      if (reviewQueueRow) {
        await deleteItem(LISTS.REVIEW_QUEUE, reviewQueueRow).catch(() => {});
      }

      const labIds = labItems.map(l => l.fullId);
      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success:true, labIds }),
      };
    } catch(e) {
      context.log('[approve-scan] Error:', e.message);
      return { status:500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
