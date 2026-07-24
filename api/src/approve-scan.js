/**
 * approve-scan.js — Azure version (v352)
 * Assigns Lab IDs, writes to SharePoint Lists, moves scan file to Archive.
 * Replaces Google Sheets + Google Drive with SharePoint Lists + Graph API.
 */
const { app } = require('@azure/functions');
const { createItem, updateItem, deleteItem, findItem, listItems, getToken, LISTS } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

// ── ET Time helpers ────────────────────────────────────────────────────────────
const TZ = 'America/New_York';
function etParts(d) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: TZ, year:'numeric', month:'2-digit', day:'2-digit',
    hour:'2-digit', minute:'2-digit', second:'2-digit', hour12:false,
  }).formatToParts(d || new Date());
  const get = t => parts.find(p => p.type === t)?.value || '';
  return { year:get('year'), month:get('month'), day:get('day'),
           hour:get('hour')==='24'?'00':get('hour'), minute:get('minute') };
}
function todayDisplayET(d) { const p=etParts(d); return `${p.month}-${p.day}-${p.year.slice(-2)}`; }
function nowTimeET() { return new Date().toLocaleTimeString('en-US',{timeZone:TZ,hour:'2-digit',minute:'2-digit',hour12:true}); }
function mmddyyET(d) { const p=etParts(d); return `${p.month}${p.day}${p.year.slice(-2)}`; }
function nextWorkdayET(from) {
  const d = from || new Date();
  const dow = new Date(d.toLocaleString('en-US',{timeZone:TZ})).getDay();
  const add = dow===5?3:dow===6?2:1;
  const next = new Date(d.getTime()+add*86400000);
  const p = etParts(next);
  return `${p.month}-${p.day}-${p.year.slice(-2)}`;
}

function fmt(iso) {
  if (!iso) return '';
  try { const [y,m,d]=iso.split('-'); return `${m}-${d}-${y.slice(-2)}`; } catch { return iso; }
}

function to24h(t) {
  if (!t) return '';
  const s = String(t).trim().replace(/^[^\d]*/, '');
  const extracted = s.match(/(\d{1,2}:\d{2}(?:\s*(?:AM|PM))?)$/i)?.[1] || s;
  const plain = extracted.match(/^(\d{1,2}):(\d{2})$/);
  if (plain) {
    const h=parseInt(plain[1]), m=parseInt(plain[2]);
    if (h>=0&&h<=23&&m>=0&&m<=59) return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
  }
  const ampm = extracted.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (ampm) {
    let h=parseInt(ampm[1]); const m=parseInt(ampm[2]);
    const isPM = ampm[3].toUpperCase()==='PM';
    if (isPM&&h<12) h+=12; if (!isPM&&h===12) h=0;
    if (h>=0&&h<=23&&m>=0&&m<=59) return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
  }
  return extracted;
}

// ── Suffix / Package maps ──────────────────────────────────────────────────────
const SUFFIX_MAP = {
  'Basic Safety (FHA)':'BS','Basic Safety':'BS','Standard Safety':'SS',
  'Expanded Safety (Mortgage Test)':'EXP','WW - Expanded Safety':'WW',
  'Comprehensive':'COMP','Pro Plus':'PP','Radon Water':'RW',
  'AIO FHA':'AIOFHA','AIO Portability':'PORT',
  'Rejected - Timeout':'REJ','Rejected - Chlorine':'REJ','Rejected - Other':'REJ',
  'Alkalinity':'ALK','Antimony':'SB','Arsenic, Total':'AS','Arsenic, Speciation':'AS-SPEC',
  'Bacteria':'BAC','Bromide':'BR','Cadmium, Total':'CD','Calcium, Total':'CA',
  'Chloride, Total':'CL','Chromium':'CR','Cobalt':'CO','Copper, Total':'CU',
  'Fluoride':'FL','Iron, Total':'FE','Lead, Total':'PB','Magnesium, Total':'MG',
  'Manganese, Total':'MN','Nitrate':'NO3','Nitrite':'NO2','pH':'PH',
  'Sodium, Total':'NA','Sulfate':'SO4','Tannins':'TAN',
  'Total Dissolved Solids (TDS)':'TDS','Total Hardness':'HRD','Uranium, Total':'U',
};

const PACKAGE_SET = new Set([
  'Basic Safety (FHA)','Basic Safety','Standard Safety',
  'Expanded Safety (Mortgage Test)','WW - Expanded Safety',
  'Comprehensive','Pro Plus','AIO FHA','AIO Portability',
]);

const PACKAGE_COVERAGE = {
  'Basic Safety (FHA)':              ['Nitrite','Nitrate','Lead, Total','Bacteria'],
  'Basic Safety':                    ['Nitrite','Nitrate','Lead, Total','Bacteria'],
  'Standard Safety':                 ['Chloride, Total','Fluoride','Nitrite','Nitrate','Iron, Total','Manganese, Total','Sodium, Total','Total Hardness','Calcium, Total','Magnesium, Total','pH','Bacteria'],
  'Expanded Safety (Mortgage Test)': ['Chloride, Total','Fluoride','Nitrite','Nitrate','Arsenic, Total','Lead, Total','Uranium, Total','Copper, Total','Iron, Total','Manganese, Total','Sodium, Total','Total Hardness','Calcium, Total','Magnesium, Total','pH','Bacteria'],
  'WW - Expanded Safety':            ['Chloride, Total','Fluoride','Nitrite','Nitrate','Arsenic, Total','Lead, Total','Uranium, Total','Copper, Total','Iron, Total','Manganese, Total','Sodium, Total','Total Hardness','Calcium, Total','Magnesium, Total','pH','Alkalinity','Bacteria'],
  'Comprehensive':                   ['Chloride, Total','Fluoride','Nitrite','Nitrate','Arsenic, Total','Lead, Total','Uranium, Total','Copper, Total','Iron, Total','Manganese, Total','Sodium, Total','Total Hardness','Calcium, Total','Magnesium, Total','Antimony','Cadmium, Total','Chromium','pH','Alkalinity','Sulfate','Bacteria'],
  'Pro Plus':                        ['Uranium, Total','Iron, Total','Manganese, Total','pH','Alkalinity','Tannins'],
};

const TEST_NORMALIZE = {
  'basic safety':'Basic Safety (FHA)','basic safety (fha)':'Basic Safety (FHA)',
  'fha':'Basic Safety (FHA)','standard safety':'Standard Safety',
  'expanded safety':'Expanded Safety (Mortgage Test)',
  'expanded safety (mortgage test)':'Expanded Safety (Mortgage Test)',
  'expanded':'Expanded Safety (Mortgage Test)','mortgage test':'Expanded Safety (Mortgage Test)',
  'ww - expanded safety':'WW - Expanded Safety','ww expanded safety':'WW - Expanded Safety',
  'comprehensive':'Comprehensive','pro plus':'Pro Plus',
  'radon water':'Radon Water','radon':'Radon Water',
  'aio fha':'AIO FHA','aio portability':'AIO Portability','portability':'AIO Portability',
  'alkalinity':'Alkalinity','antimony':'Antimony','antimony, total':'Antimony',
  'arsenic':'Arsenic, Total','arsenic, total':'Arsenic, Total','arsenic, speciation':'Arsenic, Speciation',
  'bacteria':'Bacteria','total coliform':'Bacteria','e. coli':'Bacteria','coliform':'Bacteria',
  'bromide':'Bromide','cadmium':'Cadmium, Total','cadmium, total':'Cadmium, Total',
  'calcium':'Calcium, Total','calcium, total':'Calcium, Total',
  'chloride':'Chloride, Total','chloride, total':'Chloride, Total',
  'chromium':'Chromium','chromium, total':'Chromium','cobalt':'Cobalt','cobalt, total':'Cobalt',
  'copper':'Copper, Total','copper, total':'Copper, Total','fluoride':'Fluoride',
  'iron':'Iron, Total','iron, total':'Iron, Total','lead':'Lead, Total','lead, total':'Lead, Total',
  'magnesium':'Magnesium, Total','magnesium, total':'Magnesium, Total',
  'manganese':'Manganese, Total','manganese, total':'Manganese, Total',
  'nitrate':'Nitrate','nitrate-nitrogen':'Nitrate','nitrate, total':'Nitrate',
  'nitrite':'Nitrite','nitrite-nitrogen':'Nitrite','nitrite, total':'Nitrite',
  'ph':'pH','ph electrometric':'pH','sodium':'Sodium, Total','sodium, total':'Sodium, Total',
  'sulfate':'Sulfate','tannins':'Tannins',
  'tds':'Total Dissolved Solids (TDS)','total dissolved solids':'Total Dissolved Solids (TDS)',
  'hardness':'Total Hardness','total hardness':'Total Hardness','hardness by calculation':'Total Hardness',
  'uranium':'Uranium, Total','uranium, total':'Uranium, Total',
};

function normalizeTest(t) { return TEST_NORMALIZE[t.toLowerCase().trim()] || t; }

// ── Hard-coded alias fallbacks ─────────────────────────────────────────────────
const HARD_ALIASES = {
  'maine radon water treatment':             'Maine Radon & Environmental, LLC',
  'ward water':                              'Critical Plumbing Inc. a/k/a Ward Water',
  'critical plumbing':                       'Critical Plumbing Inc. a/k/a Ward Water',
  'critical plumbing inc.a/k/a ward water': 'Critical Plumbing Inc. a/k/a Ward Water',
  'critical plumbing inc a/k/a ward water': 'Critical Plumbing Inc. a/k/a Ward Water',
  'all in one':                              'All In One Home Inspections, LLC',
  'all in one home inspections':             'All In One Home Inspections, LLC',
  'lusser team':                             'Downeast Home Inspections, LLC',
  'lussier team':                            'Downeast Home Inspections, LLC',
  'pillar to post':                          'Downeast Home Inspections, LLC',
};

function matchClient(name, clients) {
  if (!name || !clients.length) return null;
  const nl = name.toLowerCase().trim();
  const hardMatch = HARD_ALIASES[nl];
  if (hardMatch) {
    const found = clients.find(c => c.name.toLowerCase() === hardMatch.toLowerCase());
    if (found) return found;
  }
  let m = clients.find(c => c.name.toLowerCase().trim() === nl); if (m) return m;
  m = clients.find(c => (c.aliases||'').split(',').map(a=>a.trim().toLowerCase())
    .some(a => a.length >= 3 && (nl.includes(a) || a === nl))); if (m) return m;
  m = clients.find(c => { const cl=c.name.toLowerCase().trim(); return cl.includes(nl)||nl.includes(cl); }); if (m) return m;
  const STOP = new Set(['water','home','inspection','inspections','inc','llc','ltd','corp','and','the','of','for','plumbing','systems','services','environmental','radon']);
  m = clients.find(c => {
    const sigWords = c.name.toLowerCase().split(/[\s,./&]+/).filter(w=>w.length>=4&&!STOP.has(w));
    return sigWords.length>0 && sigWords.filter(w=>nl.includes(w)).length>=Math.min(2,sigWords.length);
  });
  return m || null;
}

function getAbbrev(name) {
  if (!name) return 'UNK';
  const CLIENT_ABBREV = {
    'ward water':'WW','critical plumbing':'WW','maine radon':'MR','downeast':'DEH',
    'madden':'MHI','yankee':'YHI','a-z water':'AZW','fontus':'FON','main choice':'MCI',
    'defender':'DHI','elliott':'EIS','campbell':'CPI','onpoint':'OPH','peter mason':'PM',
    'advanced':'ADV','nova':'NOV','evergreen':'EVG','fpi':'FPI','chancorp':'FPI',
    'super inspector':'SUP','all in one':'AIO',
  };
  const low = name.toLowerCase();
  for (const [k,v] of Object.entries(CLIENT_ABBREV)) { if (low.includes(k)) return v; }
  return name.split(/\s+/).map(w=>w[0]?.toUpperCase()||'').join('').slice(0,5)||'UNK';
}

// ── Get client info from SharePoint Clients list ───────────────────────────────
async function getClientInfo(token, customerName) {
  try {
    const siteId = process.env.SP_SITE_ID;
    const res = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields($select=Title,ClientCode,Abbrev,Email,Aliases,Phone)&$top=500`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return { formalName: customerName||'', clientCode:'', email:'', abbrev:'' };
    const data = await res.json();
    const clients = (data.value||[]).map(item => ({
      name:    item.fields?.Title      || '',
      code:    item.fields?.ClientCode || '',
      email:   item.fields?.Email      || '',
      abbrev:  item.fields?.Abbrev     || '',
      aliases: item.fields?.Aliases    || '',
    }));
    const match = matchClient(customerName||'', clients);
    return {
      formalName:  match ? match.name  : customerName || '',
      clientCode:  match ? match.code  : '',
      email:       match ? match.email : '',
      abbrev:      match ? match.abbrev: '',
    };
  } catch { return { formalName: customerName||'', clientCode:'', email:'', abbrev:'' }; }
}

// ── Move SP file to Archive folder ─────────────────────────────────────────────
async function moveSpFile(itemId, destFolderPath, token) {
  const siteId = process.env.SP_SITE_ID;
  const marker = 'Shared Documents/';
  const idx    = destFolderPath.indexOf(marker);
  const rel    = idx >= 0 ? destFolderPath.slice(idx + marker.length) : destFolderPath.replace(/^\/+/,'');
  const drivePath = rel.split('/').map(s=>encodeURIComponent(s)).join('/');
  try {
    const folderRes = await fetch(
      `${GRAPH}/sites/${siteId}/drive/root:/${drivePath}`,
      { headers: { Authorization:`Bearer ${token}` } }
    );
    if (!folderRes.ok) { console.warn(`[moveSpFile] Cannot resolve ${destFolderPath}`); return; }
    const destId = (await folderRes.json()).id;
    await fetch(`${GRAPH}/sites/${siteId}/drive/items/${itemId}`, {
      method:'PATCH',
      headers:{ Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
      body: JSON.stringify({ parentReference:{ id:destId } }),
    });
  } catch(e) { console.warn('[moveSpFile]', e.message); }
}

// ── Load dynamic test types from SP list ───────────────────────────────────────
async function loadDynamicTestTypes(token) {
  let dynamicSuffixMap  = { ...SUFFIX_MAP };
  let dynamicPackageSet = new Set([...PACKAGE_SET]);
  try {
    const siteId = process.env.SP_SITE_ID;
    const res = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Test Types/items?$expand=fields($select=Title,Suffix,Category,Active)&$top=200`,
      { headers: { Authorization:`Bearer ${token}` } }
    );
    if (res.ok) {
      const data = await res.json();
      for (const item of (data.value||[])) {
        const f = item.fields||{};
        if (f.Active === false || f.Active === 'FALSE') continue;
        if (f.Title && f.Suffix) dynamicSuffixMap[f.Title] = f.Suffix;
        if (f.Title) dynamicPackageSet.add(f.Title);
        if (f.Title) TEST_NORMALIZE[f.Title.toLowerCase()] = f.Title;
      }
    }
  } catch(e) { console.warn('[approve] Could not load dynamic test types:', e.message); }
  return { dynamicSuffixMap, dynamicPackageSet };
}

// ── Main handler ───────────────────────────────────────────────────────────────
app.http('approve-scan', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const {
        fileId, reviewQueueRow, reviewedBy,
        customer, isPublicOverride, dateDrawn, timeDrawn, receivedDate, receivedTime,
        location, city, state, zip, tests, hasRadon, wqReject, rwReject, notes, email,
      } = await request.json();

      if (!fileId || !tests?.length)
        return { status:400, jsonBody:{ error:'fileId and tests required' } };

      const token = await getToken();

      // Load dynamic test types from SP
      const { dynamicSuffixMap, dynamicPackageSet } = await loadDynamicTestTypes(token);

      // Normalize test names
      const normalizedTests = tests.map(normalizeTest);
      const radonRequested  = hasRadon || normalizedTests.includes('Radon Water') || wqReject || rwReject;

      const PARTIAL_REJECT_FLAGS = new Set(['WQ - Reject','RW - Reject']);
      const rejectionTests   = normalizedTests.filter(t => t.toLowerCase().startsWith('rejected'));
      const nonRejectedTests = normalizedTests.filter(t =>
        !t.toLowerCase().startsWith('rejected') &&
        t !== 'Radon Water' &&
        !PARTIAL_REJECT_FLAGS.has(t)
      );
      const selectedPackage  = nonRejectedTests.find(t => dynamicPackageSet.has(t)) || null;
      const selectedElements = nonRejectedTests.filter(t => !dynamicPackageSet.has(t));

      // Validate: block if element already covered by package
      if (selectedPackage) {
        const covered   = PACKAGE_COVERAGE[selectedPackage] || [];
        const redundant = selectedElements.filter(e => covered.includes(e));
        if (redundant.length > 0) {
          return { status:400, jsonBody:{
            error: `The following tests are already included in ${selectedPackage} and cannot be added separately: ${redundant.join(', ')}. Please remove them before approving.`,
            redundant,
          }};
        }
      }

      // Weekend Lab ID date: Sat/Sun roll back to Friday (ET)
      const n = new Date();
      const etDow = new Date(n.toLocaleString('en-US',{timeZone:TZ})).getDay();
      let labDateOffset = 0;
      if (etDow===6) labDateOffset=-1;
      if (etDow===0) labDateOffset=-2;
      const labDateET    = new Date(n.getTime()+labDateOffset*86400000);
      const mmddyy       = mmddyyET(labDateET);
      const reportDateStr = nextWorkdayET(n);

      // Get next sequence number from Accession Log
      const todayItems = await listItems(LISTS.ACCESSION_LOG, {
        filter: `fields/BaseId ge '${mmddyy}-001' and fields/BaseId le '${mmddyy}-999'`,
        top: 500,
      }).catch(() => []);
      const used = new Set(todayItems.map(r => r.BaseId).filter(Boolean));
      let seq = 1;
      while (used.has(`${mmddyy}-${String(seq).padStart(3,'0')}`)) seq++;

      const labItems = [];
      const tdStr = todayDisplayET();
      const tmStr = nowTimeET();
      const ts    = new Date().toISOString();

      // Non-radon / non-rejection lab items
      if (selectedPackage || selectedElements.length > 0) {
        const suffixParts = [];
        if (selectedPackage) suffixParts.push(dynamicSuffixMap[selectedPackage]||'GEN');
        for (const el of selectedElements) suffixParts.push(dynamicSuffixMap[el]||el.substring(0,3).toUpperCase());
        const suffix   = suffixParts.join(', ');
        const baseId   = `${mmddyy}-${String(seq).padStart(3,'0')}`;
        const wqSuffix = wqReject ? 'REJ' : suffix;
        const fullId   = `${baseId} ${wqSuffix}`;
        const allTestNames = [selectedPackage?selectedPackage:null,...selectedElements].filter(Boolean).join(' | ');
        labItems.push({ baseId, fullId, suffix:wqSuffix, isRadon:false, coaTest:allTestNames, isRejected:wqReject });
        seq++;
      }

      // Rejection items
      if (rejectionTests.length > 0) {
        const rejType = rejectionTests[0];
        const baseId  = `${mmddyy}-${String(seq).padStart(3,'0')}`;
        labItems.push({ baseId, fullId:`${baseId} REJ`, suffix:'REJ', isRadon:false, coaTest:rejType, isRejection:true, rejType });
        seq++;
      }

      // Radon item
      if (radonRequested) {
        const baseId   = `${mmddyy}-${String(seq).padStart(3,'0')}`;
        const rwSuffix = rwReject ? 'REJ' : 'RW';
        labItems.push({ baseId, fullId:`${baseId} ${rwSuffix}`, suffix:rwSuffix, isRadon:!rwReject, coaTest:'Radon Water', isRejected:rwReject });
        seq++;
      }

      if (!labItems.length)
        return { status:400, jsonBody:{ error:'No tests selected' } };

      // Get client info from SP Clients list
      const clientInfo   = await getClientInfo(token, customer);
      const clientCode   = clientInfo.clientCode;
      const formalName   = clientInfo.formalName || customer;
      const isPublicClient = formalName.startsWith('Public-');
      const abbrev       = clientInfo.abbrev || (isPublicClient ? 'PUBLIC' : getAbbrev(formalName));

      // ── Write Accession Log ──────────────────────────────────────────────────
      for (const item of labItems) {
        await createItem(LISTS.ACCESSION_LOG, {
          Title:     item.fullId,
          BaseId:    item.baseId,
          FullId:    item.fullId,
          CoaTest:   item.coaTest,
          Suffix:    item.suffix,
          Customer:  customer || '',
          ClientCode: clientCode || '',
          ReportDate: reportDateStr,
          Timestamp: ts,
          ReportStatus: 'Pending',
        }).catch(e => context.log('[AccessionLog]', e.message));
      }

      // ── Write Archived Intake ────────────────────────────────────────────────
      for (const item of labItems) {
        await createItem(LISTS.ARCHIVED_INTAKE, {
          Title:        item.fullId,
          Timestamp:    ts,
          FullId:       item.fullId,
          CoaTest:      item.coaTest,
          ClientName:   formalName || customer || '',
          ClientCode:   clientCode || '',
          DateDrawn:    fmt(dateDrawn) || '',
          TimeDrawn:    to24h(timeDrawn) || '',
          ReceivedDate: fmt(receivedDate) || tdStr,
          ReceivedTime: to24h(receivedTime) || tmStr,
          Address:      location || '',
          City:         city    || '',
          State:        state   || 'ME',
          Zip:          zip     ? String(zip).padStart(5,'0') : '',
          ReviewedBy:   reviewedBy || 'Lab Staff',
          Notes:        notes   || '',
          ReportStatus: 'Pending',
          ReportDate:   reportDateStr,
        }).catch(e => context.log('[ArchivedIntake]', e.message));
      }

      // ── Write Rejected items ─────────────────────────────────────────────────
      const rejectionItems = labItems.filter(l => l.isRejection || l.isRejected);
      for (const item of rejectionItems) {
        const rType   = item.rejType || (item.isRejected && !item.isRadon ? 'WQ - Reject' : item.isRejected ? 'RW - Reject' : 'Rejected');
        const rReason = notes || `${rType} — approved via Review Queue`;
        await createItem(LISTS.REJECTED, {
          Title:         item.fullId,
          LabId:         item.fullId,
          RejectionType: rType,
          Reason:        rReason,
          RejectedBy:    reviewedBy || 'Lab Staff',
          Timestamp:     ts,
        }).catch(e => context.log('[Rejected]', e.message));
      }

      // ── Auto-add new client if not in Clients list ───────────────────────────
      if (customer) {
        try {
          const BUSINESS_WORDS = /\b(inc|llc|ltd|corp|co\b|inspection|inspections|water|environmental|radon|plumbing|realty|real estate|services|systems|labs|laboratory|laboratories|associates|group|enterprise|properties)\b/i;
          const looksLikeBusiness = customer.includes('/')||customer.includes('&')||BUSINESS_WORDS.test(customer);
          const isPublicName = !!isPublicOverride || (!looksLikeBusiness && !!customer.trim());
          let formalClientName = customer;
          if (isPublicName) {
            const parts = customer.trim().split(/\s+/);
            if (parts.length >= 2) {
              const last = parts[parts.length-1];
              const first = parts.slice(0,parts.length-1).join(' ');
              formalClientName = `Public-${last}, ${first}`;
            } else {
              formalClientName = `Public-${customer}`;
            }
          }
          const existing = await findItem(LISTS.CLIENTS, 'Title', formalClientName).catch(()=>null);
          if (!existing) {
            await createItem(LISTS.CLIENTS, {
              Title:      formalClientName,
              ClientName: formalClientName,
              Email:      email || '',
              Abbrev:     isPublicName ? 'PUBLIC' : getAbbrev(formalClientName),
              Active:     'Yes',
            }).catch(e => context.log('[AddClient]', e.message));
          }
        } catch(e) { context.log('[AddClient] Failed:', e.message); }
      }

      // ── Delete from Review Queue ─────────────────────────────────────────────
      if (reviewQueueRow) {
        await deleteItem(LISTS.REVIEW_QUEUE, reviewQueueRow).catch(() => {});
      }

      // ── Move scan file to Archive in SharePoint ───────────────────────────────
      const SCAN_ARCHIVE = process.env.SP_SCAN_ARCHIVE || '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Archived';
      if (fileId) {
        await moveSpFile(fileId, SCAN_ARCHIVE, token).catch(e => context.log('[Archive]', e.message));
      }

      // ── Call control-sheet function to add Lab IDs ────────────────────────────
      const allBaseIds = [...new Set(labItems.map(l => l.baseId))];
      const allFullIds = labItems.map(l => l.fullId);
      try {
        const csRes = await fetch(
          `${process.env.WEBSITE_HOSTNAME ? 'https://'+process.env.WEBSITE_HOSTNAME : ''}/api/control-sheet`,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action:'addLabIds', labIds:allFullIds }),
          }
        );
        if (!csRes.ok) context.log('[CS] control-sheet returned', csRes.status);
      } catch(e) { context.log('[CS] control-sheet call failed:', e.message); }

      return {
        status: 200,
        jsonBody: {
          success:    true,
          labIds:     allFullIds,
          archiveNote: fileId ? 'File moved to Archive' : 'No file to archive',
        },
      };

    } catch(e) {
      context.log('[approve-scan] Error:', e.message);
      return { status:500, jsonBody:{ error:e.message } };
    }
  }
});
