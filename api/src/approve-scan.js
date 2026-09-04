/**
 * approve-scan.js — Azure version (v352)
 * Assigns Lab IDs, writes to SharePoint Lists, moves scan file to Archive.
 * Replaces Google Sheets + Google Drive with SharePoint Lists + Graph API.
 */
const { app } = require('@azure/functions');
const { createItem, updateItem, deleteItem, findItem, listItems, getToken, LISTS } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';
const SHEETS_ID = '15403E6ZaZFQuKNTtgJcb6-2jmlLnk02eq04BGPATqTw';
const SHEETS_TAB = 'Form Responses';

async function getSheetsToken() {
  const sa = JSON.parse(process.env.GMAIL_SERVICE_ACCOUNT || '{}');
  if (!sa.private_key) throw new Error('GMAIL_SERVICE_ACCOUNT missing');
  const crypto = require('crypto');
  const now    = Math.floor(Date.now() / 1000);
  const header  = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify({
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now, exp: now + 3600,
  })).toString('base64url');
  const sign    = crypto.createSign('RSA-SHA256');
  sign.update(`${header}.${payload}`);
  const sig = sign.sign(sa.private_key, 'base64url');
  const jwt = `${header}.${payload}.${sig}`;
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
  });
  const data = await res.json();
  if (!data.access_token) throw new Error('Sheets token failed: ' + JSON.stringify(data));
  return data.access_token;
}

async function writeToGoogleSheet(rows, context) {
  const token = await getSheetsToken();
  const range = encodeURIComponent(`${SHEETS_TAB}!A:N`);
  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SHEETS_ID}/values/${range}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: rows }),
    }
  );
  if (!res.ok) {
    const err = await res.text().catch(()=>'');
    throw new Error(`Google Sheets write failed (${res.status}): ${err.slice(0,300)}`);
  }
  const data = await res.json().catch(()=>({}));
  context.log('[Sheets] Wrote', rows.length, 'row(s) to Google Sheet', data?.updates?.updatedRange || '');
  return { success:true, updatedRange:data?.updates?.updatedRange || '', rows:rows.length };
}

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
      clientName:      item.fields?.ClientName      || item.fields?.Title || '',
      name:            item.fields?.ClientName      || item.fields?.Title || '', // alias
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
      formalName:       match ? (match.clientName || match.name) : customerName || '',
      clientName:       match ? (match.clientName || match.name) : '',
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

// ── Write to Reports to be Billed (SharePoint List) ──────────────────────────
async function writeReportsToBilled(siteId, token, params, context, cache = {}) {
  const GRAPH   = 'https://graph.microsoft.com/v1.0';
  const authHdr = { Authorization: `Bearer ${token}` };
  try {
    const listId = await getListId('Reports to be Billed', token);
    if (!listId) throw new Error('"Reports to be Billed" list not found');
    const qty = 1;
    // Look up rate from Current Pricing-V1 based on client pricing category
    let rate = 0;
    try {
      let pricingCategory = params.pricingCategory || '';
      if (!pricingCategory && params.customer) {
        if (!cache.clientPricingRows) {
          const cRes = await fetch(
            `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields($select=ClientName,PricingCategory)&$top=500`,
            { headers: authHdr });
          cache.clientPricingRows = cRes.ok ? ((await cRes.json()).value || []) : [];
        }
        const cLow  = params.customer.toLowerCase().trim();
        const found = cache.clientPricingRows.find(i => (i.fields?.ClientName||'').toLowerCase().trim() === cLow);
        pricingCategory = found?.fields?.PricingCategory || '';
      }
      const catLow = pricingCategory.toLowerCase();
      const priceCol = catLow.includes('inspector') ? 'InspectorPricing'
                     : catLow.includes('wq')        ? 'WQPricing'
                     : 'PublicPricing';
      if (!cache.pricingRows) {
        const pRes = await fetch(
          `${GRAPH}/sites/${siteId}/lists/Current%20Pricing-V1/items?$expand=fields&$top=200`,
          { headers: authHdr });
        cache.pricingRows = pRes.ok ? ((await pRes.json()).value || []) : [];
      }
      if (cache.pricingRows.length) {
        const pData  = { value: cache.pricingRows };
        // A single Lab ID can contain multiple separately ordered elements, e.g.
        // "Iron, Total | Manganese, Total | Total Hardness". Price EACH component
        // and sum them instead of accidentally matching only one element.
        const testParts   = String(params.testName || '').split(/\s*\|\s*/).map(v => v.trim()).filter(Boolean);
        const suffixParts = String(params.suffix || '').split(/\s*,\s*/).map(v => v.trim()).filter(Boolean);
        const pricingRows = pData.value || [];
        let totalRate = 0;
        const priced = [];
        for (let idx = 0; idx < Math.max(testParts.length, 1); idx++) {
          const testName = testParts[idx] || String(params.testName || '').trim();
          const suffix   = suffixParts[idx] || (testParts.length === 1 ? String(params.suffix || '').trim() : '');
          const tLow = testName.toLowerCase();
          const sLow = suffix.toLowerCase();
          // Prefer exact service/SKU matches. Only use contains matching as a final fallback.
          let match = pricingRows.find(i => (i.fields?.Service || '').toLowerCase().trim() === tLow);
          if (!match && sLow) match = pricingRows.find(i => (i.fields?.CoreAbbr_x002f_Symbol || '').toLowerCase().trim() === sLow);
          if (!match) match = pricingRows.find(i => {
            const svc = (i.fields?.Service || '').toLowerCase().trim();
            return svc && tLow && (svc.includes(tLow) || tLow.includes(svc));
          });
          const componentRate = match ? (parseFloat(String(match.fields?.[priceCol] || '').replace(/[$,]/g, '')) || 0) : 0;
          totalRate += componentRate;
          priced.push(`${testName}=$${componentRate.toFixed(2)}`);
        }
        rate = parseFloat(totalRate.toFixed(2));
        if (context) context.log(`[RTB] Pricing ${params.labId || ''}: ${priced.join(' + ')} = $${rate.toFixed(2)}`);
      }
    } catch(e) { if(context) context.log('[RTB] Rate lookup failed:', e.message); }

    const amt = rate ? parseFloat((qty * rate).toFixed(2)) : null;
    // First: get actual internal field names from the list
    if (!cache.colMap) {
      const colsRes = await fetch(
        `${GRAPH}/sites/${siteId}/lists/${listId}/columns?$select=name,displayName&$top=50`,
        { headers: authHdr }
      );
      cache.colMap = {};
      if (colsRes.ok) {
        const colData = await colsRes.json();
        (colData.value || []).forEach(c => { cache.colMap[c.displayName] = c.name; });
        if (context) context.log('[RTB] Cached column map:', JSON.stringify(cache.colMap));
      }
    }
    const colMap = cache.colMap || {};

    const labNum = (params.labId || '').split(' ')[0].trim();
    const fields = {
      Customer:         params.customer    || '',
      State:            params.state       || '',
      Zip:              params.zip ? String(params.zip).replace(/[^0-9]/g,'').padStart(5,'0') : '',
      Qty:              qty,
      Rate:             rate || null,
      Amt:              amt,
      QB:               false,
      Pd:               false,
    };
    // Add columns with special chars using bracket notation
    fields["Date Rec'd"]         = fmtExcel(params.receivedDate) || '';
    fields["Time Rec'd"]         = params.receivedTime || '';
    fields["Date Drawn"]         = fmtExcel(params.dateDrawn) || '';
    fields["Time Drawn"]         = params.timeDrawn || '';
    fields["Client Code"]        = params.clientCode || '';
    // Report Date left blank — filled in when report is actually sent
    fields["Location"]           = params.location || '';
    fields["City/Town"]          = params.city || '';
    fields["Item/Service"]       = params.testName || '';
    fields["Test Type SKU"]      = params.suffix || '';
    fields["RW Results"]         = '';
    fields["Statement/Inv Date"] = '';

    // Read-only / system fields to skip
    const SKIP_FIELDS = new Set([
      'LinkTitleNoMenu','LinkTitle','FileLeafRef','FileDirRef','FileRef',
      '_UIVersionString','Attachments','Edit','ItemChildCount','FolderChildCount',
      '_ComplianceFlags','_ComplianceTag','ContentType','id','Modified','Created',
      'AuthorLookupId','EditorLookupId','@odata.etag',
    ]);

    // Remap display names → internal names using colMap, skip system/read-only fields
    const mappedFields = {};
    for (const [displayName, value] of Object.entries(fields)) {
      const internalName = colMap[displayName] || displayName;
      if (!SKIP_FIELDS.has(internalName)) {
        mappedFields[internalName] = value;
      }
    }
    mappedFields['Title'] = labNum; // always write Title directly
    if (context) context.log('[RTB] Writing keys:', Object.keys(mappedFields).join(','));

    const res = await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items`,
      { method: 'POST', headers: { ...authHdr, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields: mappedFields }) });
    if (!res.ok) {
      const err = await res.text().catch(()=>'');
      return { success: false, error: `List write failed (${res.status}): ${err.slice(0,200)}` };
    }
    const written = await res.json();
    if (context) context.log(`[RTB] Wrote ${mappedFields.Title || mappedFields.title} rate=${rate} id=${written.id}`);
    return { success: true, id: written.id, rate: String(rate) };
  } catch(e) {
    if (context) context.log('[RTB] Error:', e.message);
    return { success: false, error: e.message };
  }
}

// ── Write to Radon Control Sheet on approval ──────────────────────────────────
async function writeRadonControlSheet(siteId, token, labId, dateDrawn, timeDrawn, receivedDate, context) {
  const MONTHS = ['January','February','March','April','May','June',
                  'July','August','September','October','November','December'];
  const baseId     = labId.split(' ')[0].trim();
  const datePrefix = baseId.slice(0, 6); // MMDDYY
  const newLabId   = labId; // full lab ID including suffix
  const mm   = datePrefix.slice(0, 2);
  const yy   = datePrefix.slice(4, 6);
  const year = '20' + yy;
  const monthName = MONTHS[parseInt(mm, 10) - 1];
  const authHdr = { Authorization: `Bearer ${token}` };
  if (context) context.log(`[RCS] Starting for ${newLabId} | month=${monthName} ${year}`);
  const controlFolder = process.env.SP_CONTROL_FOLDER ||
    '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
  const marker  = 'Shared Documents/';
  const idx     = controlFolder.indexOf(marker);
  const relPath = idx >= 0 ? controlFolder.slice(idx + marker.length) : controlFolder.replace(/^\/+/, '');
  const monthFolder     = `${monthName} Radon ${year}`;
  const monthFolderPath = `${relPath}/${monthFolder}`;
  const encMonthPath    = monthFolderPath.split('/').map(encodeURIComponent).join('/');
  const folderCheckRes  = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encMonthPath}?$select=id`, { headers: authHdr });
  if (!folderCheckRes.ok) {
    const parentEncPath = relPath.split('/').map(encodeURIComponent).join('/');
    const parentRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${parentEncPath}?$select=id`, { headers: authHdr });
    if (!parentRes.ok) throw new Error('Could not find Test C folder');
    const parentId = (await parentRes.json()).id;
    await fetch(`${GRAPH}/sites/${siteId}/drive/items/${parentId}/children`, {
      method: 'POST', headers: { ...authHdr, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: monthFolder, folder: {}, '@microsoft.graph.conflictBehavior': 'replace' }),
    });
    if (context) context.log(`[RCS] Created folder: ${monthFolder}`);
  }
  const rcsName    = `RCS_${datePrefix}.xlsx`;
  const rcsPath    = `${monthFolderPath}/${rcsName}`;
  const encRcsPath = rcsPath.split('/').map(encodeURIComponent).join('/');
  const rcsCheckRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encRcsPath}?$select=id`, { headers: authHdr });
  let rcsFileId;
  if (!rcsCheckRes.ok) {
    const masterRaw  = process.env.SP_RADON_TEMPLATE || 'Documents/Control Sheets/Master Radon Control Sheet.xlsx';
    const masterMarker = 'Shared Documents/';
    const masterIdx    = masterRaw.indexOf(masterMarker);
    const masterPath   = masterIdx >= 0 ? masterRaw.slice(masterIdx + masterMarker.length) : masterRaw.replace(/^\/+/,'');
    const encMasterPath = masterPath.split('/').map(encodeURIComponent).join('/');
    if (context) context.log(`[RCS] Looking for master at: ${masterPath}`);
    const masterRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encMasterPath}?$select=id`, { headers: authHdr });
    if (!masterRes.ok) {
      const errText = await masterRes.text().catch(()=>'');
      throw new Error(`Master Radon Control Sheet not found (${masterRes.status}): ${errText.slice(0,100)}`);
    }
    const masterId = (await masterRes.json()).id;
    const monthRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encMonthPath}?$select=id`, { headers: authHdr });
    if (!monthRes.ok) throw new Error('Monthly radon folder not found');
    const monthFolderId = (await monthRes.json()).id;
    const copyRes = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${masterId}/copy`, {
      method: 'POST', headers: { ...authHdr, 'Content-Type': 'application/json' },
      body: JSON.stringify({ parentReference: { id: monthFolderId }, name: rcsName }),
    });
    if (!copyRes.ok) throw new Error('Failed to copy master radon sheet');
    await new Promise(r => setTimeout(r, 3000));
    const newRcsRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encRcsPath}?$select=id`, { headers: authHdr });
    if (!newRcsRes.ok) throw new Error('Copied RCS file not found after copy');
    rcsFileId = (await newRcsRes.json()).id;
    if (context) context.log(`[RCS] Created ${rcsName} from master`);
  } else {
    rcsFileId = (await rcsCheckRes.json()).id;
  }
  const wbBase = `${GRAPH}/sites/${siteId}/drive/items/${rcsFileId}/workbook`;
  const sesRes = await fetch(`${wbBase}/createSession`, {
    method: 'POST', headers: { ...authHdr, 'Content-Type': 'application/json' },
    body: JSON.stringify({ persistChanges: true }),
  });
  const { id: sid } = await sesRes.json();
  const wbHdr = { ...authHdr, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
  try {
    const sheetsRes = await fetch(`${wbBase}/worksheets`, { headers: wbHdr });
    const wsId = ((await sheetsRes.json()).value || [])[0]?.id;
    if (!wsId) throw new Error('No worksheets in RCS file');
    const colARes  = await fetch(`${wbBase}/worksheets/${wsId}/range(address='A1:A200')?$select=values`, { headers: wbHdr });
    const colAVals = (await colARes.json()).values || [];
    let targetRow  = 2;
    for (let i = 1; i < colAVals.length; i++) {
      if (!String(colAVals[i][0] || '').trim()) { targetRow = i + 1; break; }
      targetRow = i + 2;
    }
    if (context) context.log(`[RCS] colA rows=${colAVals.length} targetRow=${targetRow}`);
    const todayET = new Date().toLocaleDateString('en-US', {
      timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: 'numeric'
    });
    // Use the reviewed/corrected dates from the approval card. Previously column G
    // always used today's approval date, so correcting Date Received before approval
    // never reached the Radon Control Sheet.
    const drawnFmt    = fmtExcel(dateDrawn) || dateDrawn || '';
    const receivedFmt = fmtExcel(receivedDate) || receivedDate || todayET;
    const writeRes = await fetch(`${wbBase}/worksheets/${wsId}/range(address='A${targetRow}:G${targetRow}')`, {
      method: 'PATCH', headers: wbHdr,
      body: JSON.stringify({ values: [[newLabId, '', '', '', drawnFmt, timeDrawn || '', receivedFmt]] }),
    });
    if (!writeRes.ok) {
      const errText = await writeRes.text().catch(()=>'');
      throw new Error(`RCS write failed (${writeRes.status}): ${errText.slice(0,200)}`);
    }
    if (context) context.log(`[RCS] Wrote ${newLabId} to ${rcsName} row ${targetRow}`);
    return { success: true, file: rcsName, row: targetRow };
  } finally {
    await fetch(`${wbBase}/closeSession`, { method: 'POST', headers: wbHdr }).catch(() => {});
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
      // If the client was found in the Clients list and has an abbreviation, always use it
      // (overrides PUBLIC — a listed client is a known business, not a residential public)
      const abbrev    = clientInfo.abbrev || (usePublic ? 'PUBLIC' : getAbbrev(formalName));

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
      // Run RTB writes sequentially to prevent race condition on row finding
      const rtbResults = [];
      const rtbCache = {}; // pricing/client/column metadata loaded once per approval
      for (const item of labItems) {
        // Match COA behavior: one Reports-to-be-Billed row per separately ordered
        // test/element while keeping the same base Lab ID and sample metadata.
        const rtbTests = String(item.coaTest || '')
          .split(/\s*\|\s*/)
          .map(v => v.trim())
          .filter(Boolean);
        const rtbSuffixes = String(item.suffix || '')
          .split(/\s*,\s*/)
          .map(v => v.trim())
          .filter(Boolean);
        const testsForBilling = rtbTests.length ? rtbTests : [item.coaTest || ''];

        for (let idx = 0; idx < testsForBilling.length; idx++) {
          const rowTest = testsForBilling[idx];
          const rowSuffix = rtbSuffixes[idx] || (testsForBilling.length === 1 ? (item.suffix || '') : (dynamicSuffixMap[rowTest] || ''));
          const rtbResult = await writeReportsToBilled(_siteId, _token, {
            labId:           item.baseId,
            suffix:          rowSuffix,
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
            testName:        rowTest,
          }, context, rtbCache).catch(e => ({ success:false, error:e.message }));
          rtbResults.push(rtbResult);
          context.log('[RTB]', item.baseId, rowTest, rtbResult.success ? `id ${rtbResult.id}` : rtbResult.error);
        }
      }

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
              BillingAddress:   billingAddress || clientInfo.billingAddress || [location, city, state, zip].filter(Boolean).join(', '),
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

      // ── Delete from Review Queue — direct Graph API, no shared/graph wrapper ──
      try {
        const rqSiteId  = process.env.SP_SITE_ID;
        const rqAuthHdr = { Authorization: `Bearer ${token}` };

        // Find list ID for Review Queue
        const listsRes = await fetch(`${GRAPH}/sites/${rqSiteId}/lists?$select=id,displayName`, { headers: rqAuthHdr });
        const rqListId = ((await listsRes.json()).value || []).find(l => l.displayName === 'Review Queue')?.id;

        if (rqListId) {
          // Find the row by FileID directly via Graph
          const rqItemsRes = await fetch(
            `${GRAPH}/sites/${rqSiteId}/lists/${rqListId}/items?$expand=fields($select=id,Title,FileID,FileId)&$top=500`,
            { headers: rqAuthHdr }
          );
          const rqItems = (await rqItemsRes.json()).value || [];
          const rqRow   = rqItems.find(i =>
            (i.fields?.FileID || i.fields?.FileId || '') === fileId ||
            String(i.id) === String(reviewQueueRow)
          );

          if (rqRow) {
            // PATCH Title to 'Approved' — direct Graph call, no updateItem wrapper
            const markRes = await fetch(
              `${GRAPH}/sites/${rqSiteId}/lists/${rqListId}/items/${rqRow.id}/fields`,
              { method: 'PATCH', headers: { ...rqAuthHdr, 'Content-Type': 'application/json' },
                body: JSON.stringify({ Title: 'Approved' }) }
            );
            context.log(`[ReviewQueue] Mark Approved: ${markRes.status} for item ${rqRow.id}`);

            // DELETE the row
            const delRes = await fetch(
              `${GRAPH}/sites/${rqSiteId}/lists/${rqListId}/items/${rqRow.id}`,
              { method: 'DELETE', headers: rqAuthHdr }
            );
            context.log(`[ReviewQueue] Delete: ${delRes.status} for item ${rqRow.id}`);
          } else {
            context.log(`[ReviewQueue] No row found for fileId=${fileId} reviewQueueRow=${reviewQueueRow}`);
          }
        }
      } catch(e) {
        context.log('[ReviewQueue] Error:', e.message);
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

        // Move to month/day subfolder — VERIFY success, log hard error if fails
        try {
          await moveSpFile(fileId, archiveDest, token);
          // Verify the file is now in Archive
          const siteId2 = process.env.SP_SITE_ID;
          const verifyR = await fetch(`${GRAPH}/sites/${siteId2}/drive/items/${fileId}?$select=id,name,parentReference`,
            { headers: { Authorization: `Bearer ${token}` } }).catch(() => null);
          if (verifyR?.ok) {
            const meta = await verifyR.json();
            const path = (meta.parentReference?.path || '').toLowerCase();
            if (!path.includes('archive') && !path.includes('archived')) {
              context.log(`[Archive] ⚠️ MOVE VERIFY FAILED — file ${fileId} (${meta.name}) NOT in Archive after move. Path: ${path}. Manual recovery needed.`);
            } else {
              context.log(`[Archive] ✓ Move verified — file in Archive: ${path}`);
            }
          }
        } catch(moveErr) {
          context.log(`[Archive] ⚠️ MOVE FAILED for fileId=${fileId} — file may still be in Review. Error: ${moveErr.message}. Manual recovery needed.`);
        }

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
      let csWarning = '';
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
        // Store daily control sheets in the same month folders used by import-control.
        const monthName  = now.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'long' });
        const year4      = now.toLocaleDateString('en-US', { timeZone:'America/New_York', year:'numeric' });
        const csMonthRel = `${csRel}/${monthName} ${year4}`;
        const csEncPath  = `${csMonthRel}/${csFileName}`.split('/').map(encodeURIComponent).join('/');

        const csFileRes  = await fetch(`${GRAPH}/sites/${_siteId}/drive/root:/${csEncPath}?$select=id`, { headers: { Authorization:`Bearer ${_token}` } });
        let csFileId;
        if (csFileRes.ok) {
          csFileId = (await csFileRes.json()).id;
          context.log(`[CS] Found existing ${csFileName}`);
        } else {
          // Auto-create today's control sheet from template.
          // SP_CONTROL_TEMPLATE may be a full SharePoint server-relative path, so normalize it.
          let tmplId = null;
          const tmplCandidates = [];
          const configuredTemplate = String(process.env.SP_CONTROL_TEMPLATE || '').trim();
          const normalizeDrivePath = raw => {
            const value = String(raw || '').trim();
            if (!value) return '';
            const markerIdx = value.indexOf(csMk);
            return markerIdx >= 0 ? value.slice(markerIdx + csMk.length) : value.replace(/^\/+/, '');
          };
          if (configuredTemplate) tmplCandidates.push(normalizeDrivePath(configuredTemplate));
          tmplCandidates.push(`${csRel}/C_Template.xlsx`);
          tmplCandidates.push(`${csRel}/Control Sheet Template.xlsx`);

          for (const tmplPath of [...new Set(tmplCandidates.filter(Boolean))]) {
            const tmplEnc = tmplPath.split('/').map(encodeURIComponent).join('/');
            const tmplRes = await fetch(`${GRAPH}/sites/${_siteId}/drive/root:/${tmplEnc}?$select=id,name`,
              { headers: { Authorization:`Bearer ${_token}` } });
            if (tmplRes.ok) {
              tmplId = (await tmplRes.json()).id;
              context.log(`[CS] Template found at ${tmplPath}`);
              break;
            }
            context.log(`[CS] Template candidate not found: ${tmplPath} (${tmplRes.status})`);
          }

          if (!tmplId) {
            const parentEnc = csRel.split('/').map(encodeURIComponent).join('/');
            const childrenRes = await fetch(`${GRAPH}/sites/${_siteId}/drive/root:/${parentEnc}:/children?$select=id,name,file&$top=200`,
              { headers: { Authorization:`Bearer ${_token}` } });
            if (childrenRes.ok) {
              const children = (await childrenRes.json()).value || [];
              const found = children.find(f => /template/i.test(f.name || '') && /\.xlsx?$/i.test(f.name || ''));
              if (found) {
                tmplId = found.id;
                context.log(`[CS] Auto-discovered template: ${found.name}`);
              } else {
                context.log(`[CS] No Excel template file found in ${csRel}`);
              }
            } else {
              context.log(`[CS] Could not list template folder ${csRel}: ${childrenRes.status}`);
            }
          }

          if (tmplId) {
            const csMonthEnc = csMonthRel.split('/').map(encodeURIComponent).join('/');
            let folderR = await fetch(`${GRAPH}/sites/${_siteId}/drive/root:/${csMonthEnc}?$select=id`,
              { headers: { Authorization:`Bearer ${_token}` } });
            if (!folderR.ok) {
              const parentEnc = csRel.split('/').map(encodeURIComponent).join('/');
              const parentR = await fetch(`${GRAPH}/sites/${_siteId}/drive/root:/${parentEnc}?$select=id`,
                { headers: { Authorization:`Bearer ${_token}` } });
              if (parentR.ok) {
                const parentId = (await parentR.json()).id;
                const mkR = await fetch(`${GRAPH}/sites/${_siteId}/drive/items/${parentId}/children`, {
                  method:'POST',
                  headers:{ Authorization:`Bearer ${_token}`, 'Content-Type':'application/json' },
                  body:JSON.stringify({ name:`${monthName} ${year4}`, folder:{}, '@microsoft.graph.conflictBehavior':'fail' }),
                });
                if (mkR.ok || mkR.status === 409) {
                  folderR = await fetch(`${GRAPH}/sites/${_siteId}/drive/root:/${csMonthEnc}?$select=id`,
                    { headers: { Authorization:`Bearer ${_token}` } });
                } else {
                  context.log(`[CS] Month folder create failed: ${mkR.status}`);
                }
              } else {
                context.log(`[CS] Parent control folder not found: ${parentR.status}`);
              }
            }

            if (folderR.ok) {
              const folderId = (await folderR.json()).id;
              const copyRes = await fetch(`${GRAPH}/sites/${_siteId}/drive/items/${tmplId}/copy`, {
                method:'POST',
                headers:{ Authorization:`Bearer ${_token}`, 'Content-Type':'application/json' },
                body:JSON.stringify({ parentReference:{ id:folderId }, name:csFileName }),
              });
              context.log(`[CS] Template copy request: ${copyRes.status}`);
              if (copyRes.ok || copyRes.status === 202) {
                for (let attempt = 0; attempt < 12 && !csFileId; attempt++) {
                  await new Promise(r => setTimeout(r, 1500));
                  const newR = await fetch(`${GRAPH}/sites/${_siteId}/drive/root:/${csEncPath}?$select=id`,
                    { headers: { Authorization:`Bearer ${_token}` } });
                  if (newR.ok) csFileId = (await newR.json()).id;
                }
                if (csFileId) context.log(`[CS] Created ${csFileName} from template in ${monthName} ${year4}`);
                else context.log(`[CS] Copy accepted but ${csFileName} did not appear within 18 seconds`);
              }
            }
          }
          if (!csFileId) {
            context.log(`[CS] Could not create ${csFileName} — template not found or copy failed`);
            csWarning = `Control sheet ${csFileName} not found and could not be created. Add lab IDs manually.`;
          }
        }
        if (csFileId) {
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
              context.log(`[CS] existing=${existing.size} toWrite=${newIds.length} nextRow=${nextRow2}`);
              if (newIds.length) {
                const endRow2 = nextRow2 + newIds.length - 1;
                const patchRes2 = await fetch(`${wbBase2}/worksheets/${wsId2}/range(address='A${nextRow2}:A${endRow2}')`, {
                  method:'PATCH', headers:wbHdr2,
                  body: JSON.stringify({ values: newIds.map(id=>[id]) }),
                });
                if (!patchRes2.ok) {
                  context.log(`[CS] PATCH failed ${patchRes2.status}: ${await patchRes2.text().catch(()=>'')}`);
                } else {
                  context.log(`[CS] ✓ Wrote ${newIds.join(',')} to ${csFileName} rows ${nextRow2}-${endRow2}`);
                }
              } else {
                context.log(`[CS] All IDs already present`);
              }
            }
          } finally {
            await fetch(`${wbBase2}/closeSession`, { method:'POST', headers:wbHdr2 }).catch(()=>{});
          }
        } else if (!csWarning) {
          csWarning = `Control sheet ${csFileName} not found — lab IDs not written. Create it and add manually.`;
        }
      } catch(e) { context.log('[CS] Control sheet write failed:', e.message); csWarning = `Control sheet write failed: ${e.message}`; }

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
            receivedDate || '',
            context
          );
          rcsStatus = JSON.stringify(rcsResult);
          context.log(`[RCS] ${rcsStatus}`);
        } catch(e) {
          rcsStatus = 'ERROR: ' + e.message;
          context.log('[RCS] Error:', e.message);
        }
      }

      // ── Write to Google Sheet ────────────────────────────────────────────────
      let coaSheetWarning = '';
      let coaSheetStatus = 'skipped';
      try {
        const now = new Date();
        const etNow = new Date(now.toLocaleString('en-US', { timeZone: 'America/New_York' }));
        const pad = n => String(n).padStart(2, '0');
        const reportDate = (() => {
          const d = new Date(etNow); d.setDate(d.getDate() + 1);
          if (d.getDay() === 6) d.setDate(d.getDate() + 2);
          if (d.getDay() === 0) d.setDate(d.getDate() + 1);
          return `${pad(d.getMonth()+1)}/${pad(d.getDate())}/${String(d.getFullYear()).slice(-2)}`;
        })();
        const dateRec  = receivedDate || `${pad(etNow.getMonth()+1)}/${pad(etNow.getDate())}/${String(etNow.getFullYear()).slice(-2)}`;
        const timeRec  = receivedTime || `${pad(etNow.getHours())}:${pad(etNow.getMinutes())}`;
        const coaDisplayTest = testName => {
          const name = String(testName || '').trim();
          if (name === 'Basic Safety (FHA)') return 'Basic Safety';
          if (name === 'Expanded Safety (Mortgage Test)') return 'Expanded Safety (Mortgage Test)';
          return name;
        };
        const sheetRows = labItems
          .filter(l => !l.isRejected)
          .flatMap(l => {
            // A single Lab ID may contain a package plus one or more separately ordered
            // elements. COA/Form Responses needs one row per test/element while keeping
            // the same sample metadata and Lab ID on every row.
            const rowTests = String(l.coaTest || tests?.join(' | ') || '')
              .split(/\s*\|\s*/)
              .map(coaDisplayTest)
              .filter(Boolean);
            const testsForRows = rowTests.length ? rowTests : [''];
            return testsForRows.map(coaTestName => [
              dateRec,
              timeRec,
              dateDrawn  || '',
              timeDrawn  || '',
              formalName || customer || '',
              clientCode || '',
              reportDate,
              l.baseId,
              location   || '',
              city       || '',
              state      || 'ME',
              zip        || '',
              coaTestName,
              1,
            ]);
          });
        if (sheetRows.length) {
          // Await the write before returning from the Azure Function. Fire-and-forget work
          // can be terminated as soon as the request completes, which caused intermittent/missed COA rows.
          const sheetResult = await Promise.race([
            writeToGoogleSheet(sheetRows, context),
            new Promise((_, rej) => setTimeout(() => rej(new Error('Google Sheets write timed out after 12 seconds')), 12000)),
          ]);
          coaSheetStatus = `written:${sheetResult.rows}`;
        }
      } catch(e) {
        coaSheetStatus = 'failed';
        coaSheetWarning = e.message || 'Google Sheets write failed';
        context.log('[Sheets] COA write failed:', coaSheetWarning);
      }

      // ── Write Approval to Activity Log ────────────────────────────────────────
      try {
        const _now3    = new Date();
        const _ld      = _now3.toLocaleDateString('en-US',{ timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
        const _lt      = _now3.toLocaleTimeString('en-US',{ timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
        const _labIds  = labItems.map(i => i.fullId).join(', ');
        const _tests   = labItems.map(i => i.coaTest).join(', ');
        const _details = [
          `Lab IDs: ${_labIds}`,
          `Test types: ${_tests}`,
          `Customer: ${formalName || customer || '—'}`,
          `Written to: Archived Intake | Accession Log | Reports to be Billed | Results Cache`,
          `COA scan archived | Review Queue row deleted`,
          `COA Google Sheet: ${coaSheetStatus}${coaSheetWarning ? ' — ' + coaSheetWarning : ''}`,
        ].join('\n');
        await createItem('Activity Log', {
          Title: `${_ld} ${labItems[0]?.baseId||''}`,
          Client: labItems[0]?.baseId||'',
          ActivityType: 'Sample Approved',
          Notes: _details,
          By: reviewedBy||'Lab Staff',
          LogDate: _ld, LogTime: _lt, Quantity: labItems.length,
        }).catch(()=>{});
      } catch(e) {}

      return {
        status: 200,
        jsonBody: {
          success:    true,
          labIds:     allFullIds,
          testNames:  labItems.map(l => l.coaTest),
          formalName,
          archiveNote: fileId ? 'File moved to Archive' : 'No file to archive',
          csWarning:  csWarning || undefined,
          coaSheetWarning: coaSheetWarning || undefined,
          coaSheetStatus,
        },
      };

    } catch(e) {
      context.log('[approve-scan] Error:', e.message);
      return { status:500, jsonBody:{ error:e.message } };
    }
  }
});
