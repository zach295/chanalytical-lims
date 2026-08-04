/**
 * approve-scan.js — Azure version (v352)
 * Assigns Lab IDs, writes to SharePoint Lists, moves scan file to Archive.
 * Replaces Google Sheets + Google Drive with SharePoint Lists + Graph API.
 */
const { app } = require('@azure/functions');
const { createItem, updateItem, deleteItem, findItem, listItems, getToken, LISTS } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

// Module-level cache for list IDs (avoids repeated lookups per approval)
const _listIdCache = {};
async function getListId(displayName, token) {
  if (_listIdCache[displayName]) return _listIdCache[displayName];
  const siteId = process.env.SP_SITE_ID;
  const res  = await fetch(`${GRAPH}/sites/${siteId}/lists?$select=id,displayName&$top=50`,
    { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) return null;
  const id = ((await res.json()).value||[]).find(l=>l.displayName===displayName)?.id||null;
  if (id) _listIdCache[displayName] = id;
  return id;
}

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

function nextBusinessDay(dateStr) {
  // Returns the next business day (Mon-Fri) after the given date as MM/DD/YYYY
  if (!dateStr) return '';
  try {
    // Parse MM-DD-YY or MM/DD/YYYY or YYYY-MM-DD
    let d;
    const mmddyy = dateStr.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
    if (mmddyy) {
      const yr = mmddyy[3].length === 2 ? 2000 + parseInt(mmddyy[3]) : parseInt(mmddyy[3]);
      d = new Date(yr, parseInt(mmddyy[1]) - 1, parseInt(mmddyy[2]));
    } else {
      d = new Date(dateStr);
    }
    if (isNaN(d.getTime())) return '';
    d.setDate(d.getDate() + 1);
    // Skip weekends
    while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1);
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
  } catch { return ''; }
}

function fmtExcel(dateStr) {
  // Convert any date format to MM/DD/YYYY for Excel
  if (!dateStr) return '';
  // Handle MM-DD-YY format
  const mmddyy = dateStr.match(/^(\d{1,2})-(\d{1,2})-(\d{2})$/);
  if (mmddyy) return `${mmddyy[1].padStart(2,'0')}/${mmddyy[2].padStart(2,'0')}/20${mmddyy[3]}`;
  // Handle ISO YYYY-MM-DD
  const iso = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return `${iso[2]}/${iso[3]}/${iso[1]}`;
  return dateStr;
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
  'arsenic':'Arsenic, Total','arsenic speciation':'Arsenic, Speciation','arsenic, speciation':'Arsenic, Speciation','arsenic spec':'Arsenic, Speciation','speciation':'Arsenic, Speciation','arsenic, total':'Arsenic, Total','arsenic, speciation':'Arsenic, Speciation',
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

// ── Excel serial date/time converters ─────────────────────────────────────────
// Archived Intake stores dates as Excel serial numbers and times as day fractions
function toExcelSerial(dateStr) {
  if (!dateStr) return null;
  let isoStr;
  // MM-DD-YY format
  const mmddyy = String(dateStr).match(/^(\d{2})-(\d{2})-(\d{2})$/);
  if (mmddyy) isoStr = `20${mmddyy[3]}-${mmddyy[1]}-${mmddyy[2]}T00:00:00Z`;
  // YYYY-MM-DD format
  else if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) isoStr = `${dateStr}T00:00:00Z`;
  else return null;
  const d = new Date(isoStr);
  if (isNaN(d.getTime())) return null;
  const excelEpoch = Date.UTC(1899, 11, 30);
  return Math.round((d.getTime() - excelEpoch) / 86400000);
}
function toExcelTime(timeStr) {
  if (!timeStr) return null;
  const m = String(timeStr).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const h = parseInt(m[1]), min = parseInt(m[2]);
  if (h < 0 || h > 23 || min < 0 || min > 59) return null;
  return (h * 60 + min) / 1440;
}

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

// Format public client name: "Public-Chandler, Zach" → "Zach Chandler"
function formatCustomerName(name) {
  if (!name) return '';
  if (!name.startsWith('Public-')) return name;
  const inner = name.slice('Public-'.length).trim();
  const commaIdx = inner.indexOf(',');
  if (commaIdx > 0) {
    const last  = inner.slice(0, commaIdx).trim();
    const first = inner.slice(commaIdx + 1).trim();
    return first ? `${first} ${last}` : last;
  }
  return inner;
}

// ── Get client info from SharePoint Clients list ───────────────────────────────
async function getClientInfo(token, customerName) {
  try {
    const siteId = process.env.SP_SITE_ID;
    const res = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields($select=ClientName,ClientCode,Abbrev,Email,Phone,Active,Aliases,Notes,BillingAddress,PricingCategory,BillingFrequency)&$top=500`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return { formalName: customerName||'', clientCode:'', email:'', reportEmail:'', billingEmail:'', phone:'', dbaName:'', abbrev:'', billingAddress:'', pricingCategory:'', billingFrequency:'' };
    const data = await res.json();
    const clients = (data.value||[]).map(item => ({
      name:            item.fields?.ClientName      || item.fields?.Title || '',
      code:            item.fields?.ClientCode      || '',
      reportEmail:     item.fields?.Aliases         || '',  // Report Email Address
      billingEmail:    item.fields?.Notes           || '',  // Billing Email Address
      email:           item.fields?.Aliases         || '',  // compatibility alias
      phone:           item.fields?.Active          || '',  // Phone #
      dbaName:         item.fields?.Phone           || '',  // DBA Name
      abbrev:          item.fields?.Abbrev          || '',
      billingAddress:  item.fields?.BillingAddress  || '',
      pricingCategory: item.fields?.PricingCategory || '',
      billingFrequency: item.fields?.BillingFrequency || '',
    }));
    const match = matchClient(customerName||'', clients);
    return {
      formalName:       match ? match.name           : customerName || '',
      clientCode:       match ? match.code           : '',
      email:            match ? match.email          : '',
      reportEmail:      match ? match.reportEmail    : '',
      billingEmail:     match ? match.billingEmail   : '',
      phone:            match ? match.phone          : '',
      dbaName:          match ? match.dbaName        : '',
      abbrev:           match ? match.abbrev         : '',
      billingAddress:   match ? match.billingAddress : '',
      pricingCategory:  match ? match.pricingCategory: '',
      billingFrequency: match ? match.billingFrequency: '',
    };
  } catch { return { formalName: customerName||'', clientCode:'', email:'', reportEmail:'', billingEmail:'', phone:'', dbaName:'', abbrev:'', billingAddress:'', pricingCategory:'', billingFrequency:'' }; }
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

// ── Ensure nested folder path exists, create if missing ─────────────────────────
async function ensureFolderPath(basePath, subFolders, token) {
  const siteId = process.env.SP_SITE_ID;
  let currentPath = basePath;
  for (const folderName of subFolders) {
    // Resolve parent to get its drive item ID
    const marker = 'Shared Documents/';
    const idx = currentPath.indexOf(marker);
    const rel = idx >= 0 ? currentPath.slice(idx + marker.length) : currentPath.replace(/^\/+/, '');
    const dp  = rel.split('/').map(s => encodeURIComponent(s)).join('/');
    const pRes = await fetch(
      `${GRAPH}/sites/${siteId}/drive/root:/${dp}?$select=id`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!pRes.ok) { console.warn('[ensureFolder] Parent not found:', currentPath); return currentPath; }
    const parentId = (await pRes.json()).id;
    // Create subfolder — if 409 Conflict it already exists, which is fine
    await fetch(`${GRAPH}/sites/${siteId}/drive/items/${parentId}/children`, {
      method:  'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body:    JSON.stringify({ name: folderName, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
    });
    currentPath = `${currentPath}/${folderName}`;
  }
  return currentPath;
}

// ── Load dynamic test types from SP list ───────────────────────────────────────
async function loadDynamicTestTypes(token) {
  let dynamicSuffixMap  = { ...SUFFIX_MAP };
  let dynamicPackageSet = new Set([...PACKAGE_SET]);
  let dynamicCoverage   = { ...PACKAGE_COVERAGE }; // what elements each test includes
  let dynamicPricing    = {};                       // pricing per test type
  try {
    const siteId = process.env.SP_SITE_ID;
    const res = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Current%20Pricing-V1/items?$expand=fields($select=Title,Suffix,Description,WQPrice,InspectorPrice,PublicPrice,Active)&$top=200`,
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
        // Description column: element list (newline, pipe, or <br> separated)
        const descVal = f.Description || f.Elements || '';
        if (f.Title && descVal) {
          dynamicCoverage[f.Title] = descVal
            .replace(/<[^>]+>/gi, '|')
            .split(/[\r\n|]+/)
            .map(s => s.trim().replace(/^,+|,+$/g, '').trim())
            .filter(Boolean);
        }
        // Pricing
        if (f.Title) {
          dynamicPricing[f.Title] = {
            wq:        f.WQPrice        || null,
            inspector: f.InspectorPrice || null,
            public:    f.PublicPrice    || null,
          };
        }
      }
    }
  } catch(e) { console.warn('[approve] Could not load dynamic test types:', e.message); }
  return { dynamicSuffixMap, dynamicPackageSet, dynamicCoverage, dynamicPricing };
}

// ── Main handler ───────────────────────────────────────────────────────────────

// ── Write to Radon Control Sheet on approval ──────────────────────────────────
async function writeRadonControlSheet(siteId, token, labId, dateDrawn, timeDrawn, context) {
  const MONTHS = ['January','February','March','April','May','June',
                  'July','August','September','October','November','December'];
  const baseId    = labId.split(' ')[0].trim();
  const datePrefix = baseId.slice(0, 6); // MMDDYY
  const mm  = datePrefix.slice(0, 2);
  const yy  = datePrefix.slice(4, 6);
  const year = '20' + yy;
  const monthName = MONTHS[parseInt(mm, 10) - 1];

  const authHdr = { Authorization: `Bearer ${token}` };
  if (context) context.log(`[RCS] Starting for ${labId} | siteId=${siteId} | month=${monthName} ${year}`);

  // ── 1. Ensure monthly Radon folder exists ────────────────────────────────────
  const controlFolder = process.env.SP_CONTROL_FOLDER ||
    '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
  const marker  = 'Shared Documents/';
  const idx     = controlFolder.indexOf(marker);
  const relPath = idx >= 0 ? controlFolder.slice(idx + marker.length) : controlFolder.replace(/^\/+/, '');
  const monthFolder = `${monthName} Radon ${year}`;

  const monthFolderPath = `${relPath}/${monthFolder}`;
  const encMonthPath    = monthFolderPath.split('/').map(encodeURIComponent).join('/');

  // Check if monthly folder exists, create if not
  const folderCheckRes = await fetch(
    `${GRAPH}/sites/${siteId}/drive/root:/${encMonthPath}?$select=id`,
    { headers: authHdr }
  );
  if (!folderCheckRes.ok) {
    // Create the folder
    const parentEncPath = relPath.split('/').map(encodeURIComponent).join('/');
    const parentRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${parentEncPath}?$select=id`, { headers: authHdr });
    if (!parentRes.ok) throw new Error(`Could not find Test C folder`);
    const parentId = (await parentRes.json()).id;
    await fetch(`${GRAPH}/sites/${siteId}/drive/items/${parentId}/children`, {
      method: 'POST',
      headers: { ...authHdr, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: monthFolder, folder: {}, '@microsoft.graph.conflictBehavior': 'replace' }),
    });
    context.log(`[RCS] Created folder: ${monthFolder}`);
  }

  // ── 2. Check if daily RCS file exists, copy master if not ─────────────────────
  const rcsName    = `RCS_${datePrefix}.xlsx`;
  const rcsPath    = `${monthFolderPath}/${rcsName}`;
  const encRcsPath = rcsPath.split('/').map(encodeURIComponent).join('/');

  const rcsCheckRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encRcsPath}?$select=id`, { headers: authHdr });
  let rcsFileId;

  if (!rcsCheckRes.ok) {
    // Copy master radon control sheet
    // Strip full SP path prefix if present — only need relative path after Shared Documents/
    const masterRaw  = process.env.SP_RADON_TEMPLATE ||
      'Documents/Control Sheets/Master Radon Control Sheet.xlsx';
    const masterMarker = 'Shared Documents/';
    const masterIdx    = masterRaw.indexOf(masterMarker);
    const masterPath   = masterIdx >= 0 ? masterRaw.slice(masterIdx + masterMarker.length) : masterRaw.replace(/^\/+/,'');
    const encMasterPath = masterPath.split('/').map(encodeURIComponent).join('/');
    if (context) context.log(`[RCS] Looking for master at: ${masterPath}`);
    const masterRes     = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encMasterPath}?$select=id`, { headers: authHdr });
    if (!masterRes.ok) {
      const errText = await masterRes.text().catch(()=>'');
      throw new Error(`Master Radon Control Sheet not found (${masterRes.status}): ${errText.slice(0,100)}`);
    }
    const masterId = (await masterRes.json()).id;

    // Get the monthly folder ID for the copy destination
    const monthRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encMonthPath}?$select=id`, { headers: authHdr });
    if (!monthRes.ok) throw new Error(`Monthly radon folder not found`);
    const monthFolderId = (await monthRes.json()).id;

    // Copy master → new RCS file
    const copyRes = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${masterId}/copy`, {
      method: 'POST',
      headers: { ...authHdr, 'Content-Type': 'application/json' },
      body: JSON.stringify({ parentReference: { id: monthFolderId }, name: rcsName }),
    });
    if (!copyRes.ok) throw new Error(`Failed to copy master radon sheet`);

    // Wait briefly for copy to complete
    await new Promise(r => setTimeout(r, 3000));
    const newRcsRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encRcsPath}?$select=id`, { headers: authHdr });
    if (!newRcsRes.ok) throw new Error(`Copied RCS file not found after copy`);
    rcsFileId = (await newRcsRes.json()).id;
    context.log(`[RCS] Created ${rcsName} from master`);
  } else {
    rcsFileId = (await rcsCheckRes.json()).id;
  }

  // ── 3. Open workbook session and find first empty row ─────────────────────────
  const wbBase = `${GRAPH}/sites/${siteId}/drive/items/${rcsFileId}/workbook`;

  const sesRes = await fetch(`${wbBase}/createSession`, {
    method: 'POST',
    headers: { ...authHdr, 'Content-Type': 'application/json' },
    body: JSON.stringify({ persistChanges: true }),
  });
  const { id: sid } = await sesRes.json();
  const wbHdr = { ...authHdr, 'workbook-session-id': sid, 'Content-Type': 'application/json' };

  try {
    // Get first worksheet
    const sheetsRes = await fetch(`${wbBase}/worksheets`, { headers: wbHdr });
    const wsId      = ((await sheetsRes.json()).value || [])[0]?.id;
    if (!wsId) throw new Error('No worksheets in RCS file');

    // Read column A directly to find first empty row
    const colARes  = await fetch(`${wbBase}/worksheets/${wsId}/range(address='A1:A200')?$select=values`, { headers: wbHdr });
    const colAVals = (await colARes.json()).values || [];
    let targetRow  = 2; // default: row 2 (after header)
    for (let i = 1; i < colAVals.length; i++) {
      if (!String(colAVals[i][0] || '').trim()) { targetRow = i + 1; break; }
      targetRow = i + 2;
    }
    if (context) context.log(`[RCS] colA rows=${colAVals.length} targetRow=${targetRow}`);

    // Build today's date in ET for Date Tested
    const todayET = new Date().toLocaleDateString('en-US', {
      timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: 'numeric'
    }); // MM/DD/YYYY

    // Col A: Lab Barcode #, E: Date Drawn, F: Time Drawn, G: Date Tested
    const writeRes = await fetch(`${wbBase}/worksheets/${wsId}/range(address='A${targetRow}:G${targetRow}')`, {
      method: 'PATCH',
      headers: wbHdr,
      body: JSON.stringify({
        values: [[labId, '', '', '', dateDrawn || '', timeDrawn || '', todayET]],
      }),
    });
    if (!writeRes.ok) {
      const errText = await writeRes.text().catch(()=>'');
      throw new Error(`RCS write failed (${writeRes.status}): ${errText.slice(0,200)}`);
    }

    context.log(`[RCS] Wrote ${labId} to ${rcsName} row ${targetRow}`);
    return { success: true, file: rcsName, row: targetRow };
  } finally {
    await fetch(`${wbBase}/closeSession`, { method: 'POST', headers: wbHdr }).catch(() => {});
  }
}


// ── Write to Reports to be Billed.xlsx ────────────────────────────────────────
async function writeReportsToBilled(siteId, token, params, context) {
  // params: { labId, suffix, customer, clientCode, pricingCategory,
  //           dateDrawn, timeDrawn, receivedDate, receivedTime,
  //           location, city, state, zip, testName, today }
  const GRAPH   = 'https://graph.microsoft.com/v1.0';
  const authHdr = { Authorization: `Bearer ${token}` };

  try {
    // 1. Find Reports to be Billed.xlsx in Lab Scans folder
    const labScansPath = 'Documents/Lab Scans/Reports to be Billed.xlsx';
    const encPath = labScansPath.split('/').map(encodeURIComponent).join('/');
    const fileRes = await fetch(
      `${GRAPH}/sites/${siteId}/drive/root:/${encPath}?$select=id`,
      { headers: authHdr }
    );
    if (!fileRes.ok) throw new Error(`Reports to be Billed.xlsx not found (${fileRes.status})`);
    const { id: fileId } = await fileRes.json();

    // 2. Look up Rate:
    //    Step A → find client in Clients list → get PricingCategory
    //    Step B → find test row in Current Pricing V1 → use matching column
    let rate = '';
    try {
      // Step A: Get client's PricingCategory from Clients list
      let pricingCategory = params.pricingCategory || '';
      if (!pricingCategory && params.customer) {
        const clientRes = await fetch(
          `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields($select=Title,ClientName,PricingCategory)&$top=500`,
          { headers: authHdr }
        );
        if (clientRes.ok) {
          const clientData = await clientRes.json();
          const customerLow = params.customer.toLowerCase().trim();
          const clientMatch = (clientData.value || []).find(i => {
            const name = (i.fields?.ClientName || i.fields?.Title || '').toLowerCase().trim();
            return name === customerLow || customerLow.includes(name) || name.includes(customerLow);
          });
          pricingCategory = clientMatch?.fields?.PricingCategory || '';
          if (context) context.log(`[RTB] Client "${params.customer}" → PricingCategory="${pricingCategory}"`);
        }
      }

      // Map PricingCategory to the Current Pricing V1 column name
      const catLow = pricingCategory.toLowerCase();
      let priceCol;
      if (catLow.includes('inspector')) {
        priceCol = 'InspectorPrice';
      } else if (catLow.includes('wq') || catLow.includes('water quality')) {
        priceCol = 'WQPrice';
      } else {
        priceCol = 'PublicPrice'; // default for public/unknown
      }
      if (context) context.log(`[RTB] Using price column: ${priceCol}`);

      // Step B: Find the test row in Current Pricing V1
      // Try both possible list names
      let pricingRes = await fetch(
        `${GRAPH}/sites/${siteId}/lists/Current%20Pricing-V1/items?$expand=fields($select=Title,Suffix,WQPrice,InspectorPrice,PublicPrice)&$top=200`,
        { headers: authHdr }
      );
      if (!pricingRes.ok) {
        // Try alternate name
        pricingRes = await fetch(
          `${GRAPH}/sites/${siteId}/lists/Current%20Pricing%20V1/items?$expand=fields($select=Title,Suffix,WQPrice,InspectorPrice,PublicPrice)&$top=200`,
          { headers: authHdr }
        );
        if (context) context.log(`[RTB] Tried alternate list name, status: ${pricingRes.status}`);
      }
      if (pricingRes.ok) {
        const pricingData = await pricingRes.json();
        const testNameLow = (params.testName || '').toLowerCase().trim();
        const suffixLow   = (params.suffix   || '').toLowerCase().trim();
        const match = (pricingData.value || []).find(item => {
          const f     = item.fields || {};
          const title = (f.Title  || '').toLowerCase().trim();
          const suf   = (f.Suffix || '').toLowerCase().trim();
          return title === testNameLow || suf === suffixLow;
        });
        // Log what we searched and what's in the list
        const listSample = (pricingData.value||[]).slice(0,6)
          .map(i=>`"${i.fields?.Title}"/"${i.fields?.Suffix}"`).join(', ');
        if (context) context.log(`[RTB] Searching title="${testNameLow}" suffix="${suffixLow}" | List: ${listSample}`);
        if (match) {
          rate = String(match.fields?.[priceCol] || '');
          if (context) context.log(`[RTB] Match: "${match.fields?.Title}" ${priceCol}=${rate} | fields: ${JSON.stringify(match.fields)}`);
        } else {
          if (context) context.log(`[RTB] No match for testName="${params.testName}" suffix="${params.suffix}"`);
        }
      }
    } catch(e) { if (context) context.log('[RTB] Rate lookup error:', e.message); }

    const qty  = 1;
    const amt  = rate ? (qty * parseFloat(rate)).toFixed(2) : '';

    // 3. Open workbook session
    const wbBase = `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook`;
    const sesRes = await fetch(`${wbBase}/createSession`, {
      method: 'POST',
      headers: { ...authHdr, 'Content-Type': 'application/json' },
      body: JSON.stringify({ persistChanges: true }),
    });
    if (!sesRes.ok) throw new Error(`Session open failed (${sesRes.status})`);
    const { id: sid } = await sesRes.json();
    const wbHdr = { ...authHdr, 'workbook-session-id': sid, 'Content-Type': 'application/json' };

    try {
      // 4. Get first worksheet
      const sheetsRes = await fetch(`${wbBase}/worksheets`, { headers: wbHdr });
      const wsId = ((await sheetsRes.json()).value || [])[0]?.id;
      if (!wsId) throw new Error('No worksheets found');

      // 5. Find first empty row — read column A directly (more reliable than usedRange)
      const colARes  = await fetch(
        `${wbBase}/worksheets/${wsId}/range(address='A1:A500')?$select=values`,
        { headers: wbHdr }
      );
      const colAData = await colARes.json();
      const colAVals = colAData.values || [];
      // Find first empty cell in column A after header (row 1 = index 0)
      let nextRow = 2; // default: row 2
      for (let i = 1; i < colAVals.length; i++) {
        if (!String(colAVals[i][0] || '').trim()) {
          nextRow = i + 1; // convert to 1-based row number
          break;
        }
        nextRow = i + 2; // all rows filled — append after last
      }

      // 6. Write in TWO ranges to skip column R (Amt formula — do not overwrite)
      // A:Q = Date Rec'd through Rate
      const rowAQ = [
        fmtExcel(params.receivedDate) || '',  // A Date Rec'd
        params.receivedTime || '',              // B Time Rec'd
        fmtExcel(params.dateDrawn) || '',       // C Date Drawn
        params.timeDrawn    || '',              // D Time Drawn
        params.customer     || '',              // E Customer
        params.clientCode   || '',              // F Client Code
        nextBusinessDay(params.receivedDate || params.dateDrawn), // G Report Date
        (params.labId || '').split(' ')[0].trim() || '',  // H Lab #
        params.location     || '',              // I Location
        params.city         || '',              // J City/Town
        params.state        || '',              // K State
        params.zip ? String(params.zip).replace(/\D/g,'').padStart(5,'0') : '',  // L Zip
        params.testName     || '',              // M Item/Service
        params.suffix       || '',              // N Test Type SKU
        '',                                     // O RW Results (blank)
        qty,                                    // P Qty
        rate,                                   // Q Rate
      ];
      // S:X = QB through Date Pd (skip R=Amt which has a formula)
      const rowSX = ['', '', '', '', '', ''];   // S-X blank

      await fetch(`${wbBase}/worksheets/${wsId}/range(address='A${nextRow}:Q${nextRow}')`, {
        method: 'PATCH', headers: wbHdr,
        body: JSON.stringify({ values: [rowAQ] }),
      });
      await fetch(`${wbBase}/worksheets/${wsId}/range(address='S${nextRow}:X${nextRow}')`, {
        method: 'PATCH', headers: wbHdr,
        body: JSON.stringify({ values: [rowSX] }),
      });

      context.log(`[RTB] Wrote ${params.labId} to row ${nextRow}`);
      return { success: true, row: nextRow };
    } finally {
      await fetch(`${wbBase}/closeSession`, { method: 'POST', headers: wbHdr }).catch(() => {});
    }
  } catch(e) {
    context.log('[RTB] Error:', e.message);
    return { success: false, error: e.message };
  }
}

app.http('approve-scan', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const {
        fileId, reviewQueueRow, reviewedBy,
        customer, isPublicOverride, dateDrawn, timeDrawn, receivedDate, receivedTime,
        location, city, state, zip, tests, hasRadon, wqReject, rwReject, notes, email,
        phone, billingAddress,
      } = await request.json();

      if (!fileId || !tests?.length)
        return { status:400, jsonBody:{ error:'fileId and tests required' } };

      const token = await getToken();

      // Parallelize independent startup reads
      const [{ dynamicSuffixMap, dynamicPackageSet, dynamicCoverage, dynamicPricing }, clientInfo] = await Promise.all([
        loadDynamicTestTypes(token),
        getClientInfo(token, customer),
      ]);

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
        const covered   = dynamicCoverage[selectedPackage] || PACKAGE_COVERAGE[selectedPackage] || [];
        const redundant = selectedElements.filter(e => covered.includes(e));
        if (redundant.length > 0) {
          return { status:400, jsonBody:{
            error: `The following tests are already included in ${selectedPackage} and cannot be added separately: ${redundant.join(', ')}. Please remove them before approving.`,
            redundant,
          }};
        }
      }

      // Sat/Sun use their own date and get their own lab ID sequence and control sheet
      const n = new Date();
      const mmddyy       = mmddyyET(n);
      const reportDateStr = nextWorkdayET(n);

      // Get next sequence number from Accession Log
      // field_1 = baseId in the Excel-imported Accession Log list
      const todayItems = await listItems(LISTS.ACCESSION_LOG, { top: 500 }).catch(() => []);
      const used = new Set(
        todayItems
          .map(r => r.field_1 || '')
          .filter(id => id.startsWith(mmddyy + '-'))
      );
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
      // clientInfo loaded in parallel above
      const clientCode   = clientInfo.clientCode;
      const formalName   = clientInfo.formalName || customer;
      const isPublicClient = formalName.startsWith('Public-');
      // Only use PUBLIC abbrev for genuine public (residential) clients
      // Business clients (with LLC, Inc, Inspections etc.) should never get PUBLIC
      const BWORDS = /\b(inc|llc|ltd|corp|co\b|inspection|inspections|water|environmental|radon|plumbing|realty|services|systems|labs|laboratory|laboratories|associates|group|enterprise|properties|testing|analysis|real estate)\b/i;
      const nameIsBusinessLike = BWORDS.test(formalName) || formalName.includes('/') || formalName.includes('&');
      const usePublic = (isPublicOverride || isPublicClient) && !nameIsBusinessLike;
      const abbrev    = usePublic ? 'PUBLIC' : (clientInfo.abbrev || getAbbrev(formalName));

      // ── Write Accession Log ──────────────────────────────────────────────────
      // Field mapping (Excel-imported): Title=timestamp, field_1=baseId,
      // field_2=fullId, field_3=coaTest, field_4=suffix
      for (const item of labItems) {
        await createItem(LISTS.ACCESSION_LOG, {
          Title:   ts,
          field_1: item.baseId,
          field_2: item.fullId,
          field_3: item.coaTest,
          field_4: item.suffix,
        }).catch(e => context.log('[AccessionLog]', e.message));
      }

      // ── Write Archived Intake ────────────────────────────────────────────────
      // Use direct Graph API to avoid LISTS constant dependency issues
      const _siteId = process.env.SP_SITE_ID;
      const _token  = await getToken();
      const archivedIntakeListId = await (async () => {
        const r = await fetch(`${GRAPH}/sites/${_siteId}/lists?$select=id,displayName`,
          { headers: { Authorization: `Bearer ${_token}` } });
        const d = await r.json();
        return (d.value||[]).find(l=>l.displayName==='Archived Intake')?.id || null;
      })();
      if (!archivedIntakeListId) context.log('[ArchivedIntake] WARNING: list not found!');

      for (const item of labItems) {
        const intakeFields = {
          Title:    ts,
          field_1:  item.fullId,
          field_2:  item.coaTest,
          field_3:  formalName || customer || '',
          field_4:  fmt(dateDrawn) || '',
          field_5:  to24h(timeDrawn) || '',
          field_6:  fmt(receivedDate) || tdStr,
          field_7:  to24h(receivedTime) || tmStr,
          field_8:  location   || '',
          field_9:  city       || '',
          field_10: state      || 'ME',
          field_11: zip        ? String(zip).padStart(5,'0') : '',
          field_12: reviewedBy || 'Lab Staff',
          field_14: item.isRejection ? 'Rejected' : 'Pending',
        };
        if (notes && notes.trim()) intakeFields.field_13 = notes;
        // For rejections, also store rejection type in notes if no other note
        if (item.isRejection && !notes?.trim()) {
          intakeFields.field_13 = item.rejType || item.coaTest || 'Rejected';
        }
        let intakeResult = null;
        let intakeError  = null;
        if (archivedIntakeListId) {
          const ir = await fetch(
            `${GRAPH}/sites/${_siteId}/lists/${archivedIntakeListId}/items`,
            { method: 'POST', headers: { Authorization: `Bearer ${_token}`, 'Content-Type': 'application/json' },
              body: JSON.stringify({ fields: intakeFields }) }
          );
          if (ir.ok) {
            intakeResult = await ir.json();
            context.log('[ArchivedIntake] ✓ wrote', item.fullId, 'id=', intakeResult.id);
          } else {
            intakeError = `${ir.status}: ${await ir.text().catch(()=>'')}`;
            context.log('[ArchivedIntake] FAILED', intakeError);
          }
        }
      }

      // ── Write Rejected items ─────────────────────────────────────────────────
      const rejectionItems = labItems.filter(l => l.isRejection || l.isRejected);
      for (const item of rejectionItems) {
        const rType   = item.rejType || (item.isRejected && !item.isRadon ? 'WQ - Reject' : item.isRejected ? 'RW - Reject' : 'Rejected');
        const rReason = notes || `${rType} — approved via Review Queue`;
        // Direct Graph API write to bypass LISTS constant issues
        try {
          const rejListRes = await fetch(
            `${GRAPH}/sites/${_siteId}/lists?$select=id,displayName`,
            { headers: { Authorization: `Bearer ${_token}` } }
          );
          const rejListId = ((await rejListRes.json()).value || [])
            .find(l => l.displayName === 'Rejected')?.id;
          if (rejListId) {
            const rr = await fetch(
              `${GRAPH}/sites/${_siteId}/lists/${rejListId}/items`,
              { method: 'POST',
                headers: { Authorization: `Bearer ${_token}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ fields: {
                  Title:   ts,
                  field_1: item.fullId,
                  field_2: rType,
                  field_3: rReason,
                  field_4: reviewedBy || 'Lab Staff',
                }})
              }
            );
            if (rr.ok) context.log('[Rejected] ✓ wrote', item.fullId);
            else context.log('[Rejected] FAILED', rr.status, await rr.text().catch(()=>''));
          } else {
            context.log('[Rejected] list not found');
          }
        } catch(e) { context.log('[Rejected] error:', e.message); }
        // Also write to Activity Log
        const actNow2  = new Date();
        const logDate2 = actNow2.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
        const logTime2 = actNow2.toLocaleTimeString('en-US', { timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
        await createItem('Activity Log', {
          Title: `${logDate2} ${item.fullId}`,
          Client: item.fullId, ActivityType: rType,
          Notes: rReason, By: reviewedBy || 'Lab Staff',
          LogDate: logDate2, LogTime: logTime2, Quantity: 0,
        }).catch(e => context.log('[ActivityLog/Rejected]', e.message));
      }

      // ── Write Results Cache ──────────────────────────────────────────────────
      for (const item of labItems) {
        await createItem('Results Cache', {
          LabID: item.baseId,
        }).catch(e => context.log('[ResultsCache]', e.message));
      }

      // ── Write RCS and RTB in parallel ────────────────────────────────────────
      const rtbPromises = labItems
        .map(item => writeReportsToBilled(_siteId, _token, {
          labId:           item.fullId,
          suffix:          item.suffix      || '',
          customer:        formalName       || customer || '',
          clientCode:      clientCode       || '',
          pricingCategory: clientInfo.pricingCategory || '',
          dateDrawn:       fmt(dateDrawn)   || '',
          timeDrawn:       to24h(timeDrawn) || '',
          receivedDate:    fmt(receivedDate)|| tdStr,
          receivedTime:    to24h(receivedTime)||tmStr,
          location:        location         || '',
          city:            city             || '',
          state:           state            || 'ME',
          zip:             zip ? String(zip).padStart(5,'0') : '',
          testName:        item.coaTest     || '',
        }, context).catch(e => ({ success:false, error:e.message })));
      const rtbResults = await Promise.all(rtbPromises);
      rtbResults.forEach((r,i) => context.log('[RTB]', labItems[i]?.fullId, r.success ? `row ${r.row}` : r.error));

      // ── Auto-add new client ONLY if not already in Clients list ─────────────
      // STRICT RULE: if clientInfo found a match, skip creation entirely
      const clientAlreadyExists = !!(clientInfo && clientInfo.clientName && clientInfo.clientName.trim());
      if (customer && !clientAlreadyExists) {
        try {
          const BUSINESS_WORDS = /\b(inc|llc|ltd|corp|co\b|inspection|inspections|water|environmental|radon|plumbing|realty|real estate|services|systems|labs|laboratory|laboratories|associates|group|enterprise|properties|testing|analysis|analytic|analytics|engineering|consultants|consulting|solutions|management|partners|professionals|experts|company|businesses|industries|contractors|construction|renovations|hvac|electric|electrical|mechanical|roofing|flooring|painting|appliance|home|remodeling|restoration|development)\b/i;
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
              ClientName:       formalClientName,
              Aliases:          clientInfo.reportEmail || email || '',         // Report Email Address
              Notes:            clientInfo.billingEmail || email || '',        // Billing Email Address
              Active:           clientInfo.phone || phone || '',               // Phone #
              BillingAddress:   clientInfo.billingAddress || billingAddress || [location, city, state, zip].filter(Boolean).join(', '), // Billing Address
              Phone:            clientInfo.dbaName || '',                // DBA Name
              BillingAddress:   clientInfo.billingAddress || '',
              ClientCode:       '',
              Abbrev:           isPublicName ? 'PUBLIC' : getAbbrev(formalClientName),
              BillingFrequency: isPublicName ? 'Pre-Pay' : '',
              PricingCategory:  isPublicName ? 'Public Pricing' : '',
              StartDate:        new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' }),
              Status:           'Active',
              RadonLic_x0023_:  '',
            }).catch(e => context.log('[AddClient]', e.message));
          }
        } catch(e) { context.log('[AddClient] Failed:', e.message); }
      }

      // ── Delete from Review Queue ─────────────────────────────────────────────
      if (reviewQueueRow) {
        await deleteItem(LISTS.REVIEW_QUEUE, reviewQueueRow).catch(() => {});
      }

      // ── Move and rename scan file to Archive (organized by Month/Day) ──────────
      const SCAN_ARCHIVE = process.env.SP_SCAN_ARCHIVE || '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Archived';
      if (fileId) {
        // Build month/day subfolder path: Archived/August 2026/3/
        const archiveDate  = new Date();
        const monthFolder  = archiveDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric', timeZone: 'America/New_York' });
        // Day folder formatted as MM-DD-YY e.g. 08-04-26
        const dayFolder    = archiveDate.toLocaleDateString('en-US', {
          month: '2-digit', day: '2-digit', year: '2-digit', timeZone: 'America/New_York'
        }).replace(/\//g, '-');
        const archiveDest  = await ensureFolderPath(SCAN_ARCHIVE, [monthFolder, dayFolder], token)
          .catch(e => { context.log('[Archive folder]', e.message); return SCAN_ARCHIVE; });
        context.log(`[Archive] Destination: ${archiveDest}`);
        // Move to month/day subfolder
        await moveSpFile(fileId, archiveDest, token).catch(e => context.log('[Archive move]', e.message));

        // Rename PDF to [LabID]_[ClientAbbrev]_[Address].pdf
        // For radon: [LabID] RW_[ClientAbbrev]_[Address].pdf
        try {
          const sanitize = s => String(s||'').replace(/[\/\\:*?"<>|]/g,'').trim().slice(0,60);
          const addrPart   = sanitize(location || '');
          const abbrevPart = sanitize(abbrev || getAbbrev(formalName || customer || '') || 'UNK');
          // Format: [baseId]_[abbrev]_[address].pdf (spaces kept in address)
          // For radon: [baseId] RW_[abbrev]_[address].pdf
          const siteId     = process.env.SP_SITE_ID;
          const nonRadon   = labItems.find(l => !l.isRadon && !l.isRejection);
          const radonItem  = labItems.find(l => l.isRadon && !l.isRejected);
          const primaryItem = nonRadon || labItems[0];
          const prefix = primaryItem.isRadon ? `${primaryItem.baseId} RW` : primaryItem.baseId;
          const newName = `${prefix}_${abbrevPart}_${addrPart}.pdf`;

          // Rename original PDF
          await fetch(`${GRAPH}/sites/${siteId}/drive/items/${fileId}`, {
            method:  'PATCH',
            headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
            body:    JSON.stringify({ name: newName }),
          });
          context.log(`[Archive] Renamed to ${newName}`);

          // If both radon and non-radon — copy PDF for the radon item too
          if (radonItem && nonRadon) {
            try {
              const rwName = `${radonItem.baseId} RW_${abbrevPart}_${addrPart}.pdf`;
              const marker = 'Shared Documents/';
              const idx    = archiveDest.indexOf(marker);
              const rel    = idx >= 0 ? archiveDest.slice(idx + marker.length) : archiveDest.replace(/^\/+/,'');
              const dp     = rel.split('/').map(s => encodeURIComponent(s)).join('/');
              const fRes   = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${dp}?$select=id`,
                { headers: { Authorization: `Bearer ${token}` } });
              if (fRes.ok) {
                const destId = (await fRes.json()).id;
                await fetch(`${GRAPH}/sites/${siteId}/drive/items/${fileId}/copy`, {
                  method:  'POST',
                  headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
                  body:    JSON.stringify({ parentReference: { id: destId }, name: rwName }),
                });
                context.log(`[Archive] Copied RW PDF: ${rwName}`);
              }
            } catch(copyErr) { context.log('[Archive RW copy]', copyErr.message); }
          }
        } catch(e) { context.log('[Archive rename]', e.message); }
      }

      // ── Write Lab IDs to Control Sheet (direct Graph API) ─────────────────────
      const allFullIds = labItems.map(l => l.fullId);
      try {
        // Get today's control sheet file: C_MMDDYY.xlsx in SP_CONTROL_FOLDER
        const now        = new Date();
        const etDate     = now.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
        const csPrefix   = etDate.replace(/\//g,'');  // MMDDYY
        const csFileName = `C_${csPrefix}.xlsx`;
        const csFolder   = process.env.SP_CONTROL_FOLDER || '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
        const csMk       = 'Shared Documents/';
        const csIdx      = csFolder.indexOf(csMk);
        const csRel      = csIdx >= 0 ? csFolder.slice(csIdx + csMk.length) : csFolder.replace(/^\/+/,'');
        const csEncPath  = `${csRel}/${csFileName}`.split('/').map(encodeURIComponent).join('/');

        const csFileRes  = await fetch(`${GRAPH}/sites/${_siteId}/drive/root:/${csEncPath}?$select=id`, { headers: { Authorization:`Bearer ${_token}` } });
        if (csFileRes.ok) {
          const csFileId = (await csFileRes.json()).id;
          const wbBase2  = `${GRAPH}/sites/${_siteId}/drive/items/${csFileId}/workbook`;
          const ses2Res  = await fetch(`${wbBase2}/createSession`, {
            method:'POST', headers:{ Authorization:`Bearer ${_token}`, 'Content-Type':'application/json' },
            body: JSON.stringify({ persistChanges:true }),
          });
          const { id: sid2 } = await ses2Res.json();
          const wbHdr2 = { Authorization:`Bearer ${_token}`, 'workbook-session-id':sid2, 'Content-Type':'application/json' };
          try {
            const sheetsRes2 = await fetch(`${wbBase2}/worksheets`, { headers:wbHdr2 });
            const wsId2      = ((await sheetsRes2.json()).value||[])[0]?.id;
            if (wsId2) {
              // Scan column A for existing IDs and first empty row
              const colARes2 = await fetch(`${wbBase2}/worksheets/${wsId2}/range(address='A1:A500')?$select=values`, { headers:wbHdr2 });
              const colAVals2 = (await colARes2.json()).values || [];
              const existing = new Set(colAVals2.map(r => String(r[0]||'').trim()).filter(Boolean));
              let nextRow2 = 2;
              for (let i = 1; i < colAVals2.length; i++) {
                if (!String(colAVals2[i][0]||'').trim()) { nextRow2 = i+1; break; }
                nextRow2 = i+2;
              }
              const newIds = allFullIds.filter(id => !existing.has(id.trim()));
              if (newIds.length) {
                const endRow2 = nextRow2 + newIds.length - 1;
                await fetch(`${wbBase2}/worksheets/${wsId2}/range(address='A${nextRow2}:A${endRow2}')`, {
                  method:'PATCH', headers:wbHdr2,
                  body: JSON.stringify({ values: newIds.map(id=>[id]) }),
                });
                context.log(`[CS] Wrote ${newIds.length} IDs to ${csFileName} starting row ${nextRow2}`);
              }
            }
          } finally {
            await fetch(`${wbBase2}/closeSession`, { method:'POST', headers:wbHdr2 }).catch(()=>{});
          }
        } else {
          context.log(`[CS] ${csFileName} not found (${csFileRes.status}) — create it first`);
        }
      } catch(e) { context.log('[CS] Control sheet write failed:', e.message); }

      // ── Write to Radon Control Sheet if Radon Water approved ─────────────────
      const radonLabItem = labItems.find(l => l.isRadon && !l.isRejected);
      context.log('[RCS] radonLabItem:', radonLabItem ? radonLabItem.fullId : 'none');
      context.log('[RCS] _siteId:', _siteId, 'token length:', _token?.length);
      let rcsStatus = 'skipped';
      if (radonLabItem) {
        try {
          const rcsResult = await writeRadonControlSheet(
            _siteId, _token, radonLabItem.fullId,
            dateDrawn || '',
            to24h(timeDrawn) || timeDrawn || '',
            context
          );
          rcsStatus = JSON.stringify(rcsResult);
          context.log(`[RCS] ${rcsStatus}`);
        } catch(e) {
          rcsStatus = 'ERROR: ' + e.message;
          context.log('[RCS] Error:', e.message);
        }
      }

      return {
        status: 200,
        jsonBody: {
          success:    true,
          labIds:     allFullIds,
          formalName,
          archiveNote: fileId ? 'File moved to Archive' : 'No file to archive',
          rcsStatus,
          rtbDebug:   rtbResults?.[0] || null,

        },
      };

    } catch(e) {
      context.log('[approve-scan] Error:', e.message);
      return { status:500, jsonBody:{ error:e.message } };
    }
  }
});
