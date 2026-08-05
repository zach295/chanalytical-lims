/**
 * generate-report.js — Azure version v2
 * Pulls ALL data from SharePoint Lists (Archived Intake + Results Cache)
 * instead of re-reading Excel files.
 */
const { app }      = require('@azure/functions');
const { getToken, listItems, LISTS } = require('../shared/graph');
const GRAPH = 'https://graph.microsoft.com/v1.0';

// ── Parameter config ──────────────────────────────────────────────────────────
const PARAM_CONFIG = [
  { name:'Chloride, Total',              rl:2.00,    epa:250,       unit:'mg/L', method:'SM4500Cl',         source:'gallery', cacheField:'field_6',  cacheDT:'field_7' },
  { name:'Fluoride, Total',              rl:0.20,    epa:4,         unit:'mg/L', method:'SM4500F',          source:'gallery', cacheField:'field_8',  cacheDT:'field_9' },
  { name:'Nitrite-Nitrogen, Total',      rl:0.20,    epa:1,         unit:'mg/L', method:'EPA 354.1',        source:'gallery', cacheField:'field_10', cacheDT:'field_11' },
  { name:'Nitrate-Nitrogen, Total',      rl:1.00,    epa:10,        unit:'mg/L', method:'SM4500NO3',        source:'gallery', cacheField:'field_12', cacheDT:'field_13' },
  { name:'Arsenic, III',                 rl:1.00,    epa:10,        unit:'ug/L', method:'EPA 200.8',        source:'spec_iii', cacheField:'ArsenicIII' },
  { name:'Arsenic, V',                   rl:1.00,    epa:10,        unit:'ug/L', method:'EPA 200.8',        source:'spec_v',   cacheField:null },
  { name:'Arsenic, Total',               rl:1.00,    epa:10,        unit:'ug/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Arsenic_x0028_As75_x0029_' },
  { name:'Lead, Total',                  rl:1.00,    epa:15,        unit:'ug/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Lead_x0028_Pb208_x0029_' },
  { name:'Uranium, Total',               rl:1.00,    epa:30,        unit:'ug/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Uranium_x0028_U238_x0029_' },
  { name:'Copper, Total',                rl:0.001,   epa:1.3,       unit:'mg/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Copper_x0028_Cu63_x0029_' },
  { name:'Iron, Total',                  rl:0.05,    epa:0.3,       unit:'mg/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Iron_x0028_Fe54_x0029_' },
  { name:'Manganese, Total',             rl:0.001,   epa:0.05,      unit:'mg/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Manganese_x0028_Mn55_x0029_' },
  { name:'Sodium, Total',                rl:1.00,    epa:null,      unit:'mg/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Sodium_x0028_Na23_x0029_' },
  { name:'Hardness by calculation',      rl:0.91,    epa:null,      unit:'mg/L', method:'',                 source:'calc' },
  { name:'Calcium, Total',               rl:0.2,     epa:null,      unit:'mg/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Calcium_x0028_Ca43_x0029_' },
  { name:'Magnesium, Total',             rl:0.1,     epa:null,      unit:'mg/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Magnesium_x0028_Mg24_x0029_' },
  { name:'Antimony, Total',              rl:0.0005,  epa:0.006,     unit:'mg/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Antimony_x0028_Sb121_x0029_' },
  { name:'Cadmium, Total',               rl:0.002,   epa:0.005,     unit:'mg/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Cadmium_x0028_Cd111_x0029_' },
  { name:'Chromium, Total',              rl:0.002,   epa:0.1,       unit:'mg/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Chromium_x0028_Cr52_x0029_' },
  { name:'Cobalt',                       rl:null,    epa:null,      unit:'mg/L', method:'EPA 200.8',        source:'icpms',   cacheField:'Cobalt_x0028_Co59_x0029_' },
  { name:'pH Electrometric',             rl:null,    epa:'6.5-8.5', unit:'',     method:'SM4500H+B',        source:'ph',      decimals:2 },
  { name:'Alkalinity',                   rl:40.00,   epa:null,      unit:'mg/L', method:'',                 source:'gallery', cacheField:'field_14', cacheDT:'field_15' },
  { name:'Sulfate',                      rl:40.00,   epa:250,       unit:'mg/L', method:'SM4500-SO4',       source:'gallery', cacheField:'field_16', cacheDT:'field_17' },
  { name:'Tannins',                      rl:null,    epa:null,      unit:'',     method:'Hach Method 8193', source:'gallery', cacheField:'field_18', cacheDT:'field_19' },
  { name:'Total Dissolved Solids (TDS)', rl:null,    epa:null,      unit:'ppm',  method:'SM4500C1E',        source:'gallery', cacheField:'field_20', cacheDT:'field_21' },
  { name:'Bromide',                      rl:null,    epa:null,      unit:'mg/L', method:'HI 93716',         source:'gallery', cacheField:'field_22', cacheDT:'field_23' },
  { name:'Total Coliform',               rl:null,    epa:1,         unit:'MPN',  method:'SM9223 B',         source:'bac',     cacheField:'field_2' },
  { name:'E. Coli',                      rl:null,    epa:1,         unit:'MPN',  method:'SM9223 B',         source:'bac',     cacheField:'field_3' },
];

// Maps service names (from approve-scan) to PARAM_CONFIG names where they differ
const PARAM_SERVICE_ALIASES = {
  'pH':                    ['pH Electrometric'],
  'Nitrate':               ['Nitrate-Nitrogen, Total'],
  'Nitrite':               ['Nitrite-Nitrogen, Total'],
  'Fluoride':              ['Fluoride, Total'],
  'Hardness':              ['Hardness by calculation'],
  'Total Hardness':        ['Hardness by calculation'],
  'Bacteria':              ['Total Coliform', 'E. Coli'],
  'Total Coliform':        ['Total Coliform'],
  'E. Coli':               ['E. Coli'],
  'Chloride':              ['Chloride, Total'],
  'Arsenic':               ['Arsenic, Total'],
  'Arsenic III':          ['Arsenic, III'],
  'Arsenic V':            ['Arsenic, V'],
  'Lead':                  ['Lead, Total'],
  'Uranium':               ['Uranium, Total'],
  'Copper':                ['Copper, Total'],
  'Iron':                  ['Iron, Total'],
  'Manganese':             ['Manganese, Total'],
  'Sodium':                ['Sodium, Total'],
  'Calcium':               ['Calcium, Total'],
  'Magnesium':             ['Magnesium, Total'],
  'Antimony':              ['Antimony, Total'],
  'Cadmium':               ['Cadmium, Total'],
  'Chromium':              ['Chromium, Total'],
  'Cobalt':                ['Cobalt'],
};

// Hardcoded fallback — reports always work even when SP list is unreachable
const PACKAGE_COVERAGE_FALLBACK = {
  'Basic Safety (FHA)':              ['Nitrite-Nitrogen, Total','Nitrate-Nitrogen, Total','Lead, Total','Total Coliform','E. Coli'],
  'Basic Safety':                    ['Nitrite-Nitrogen, Total','Nitrate-Nitrogen, Total','Lead, Total','Total Coliform','E. Coli'],
  'Standard Safety':                 ['Chloride, Total','Fluoride, Total','Nitrite-Nitrogen, Total','Nitrate-Nitrogen, Total','Iron, Total','Manganese, Total','Sodium, Total','Hardness by calculation','Calcium, Total','Magnesium, Total','pH Electrometric','Total Coliform','E. Coli'],
  'Expanded Safety (Mortgage Test)': ['Chloride, Total','Fluoride, Total','Nitrite-Nitrogen, Total','Nitrate-Nitrogen, Total','Arsenic, Total','Lead, Total','Uranium, Total','Copper, Total','Iron, Total','Manganese, Total','Sodium, Total','Hardness by calculation','Calcium, Total','Magnesium, Total','pH Electrometric','Total Coliform','E. Coli'],
  'WW - Expanded Safety':            ['Chloride, Total','Fluoride, Total','Nitrite-Nitrogen, Total','Nitrate-Nitrogen, Total','Arsenic, Total','Lead, Total','Uranium, Total','Copper, Total','Iron, Total','Manganese, Total','Sodium, Total','Hardness by calculation','Alkalinity','Calcium, Total','Magnesium, Total','pH Electrometric','Total Coliform','E. Coli'],
  'Comprehensive':                   ['Chloride, Total','Fluoride, Total','Nitrite-Nitrogen, Total','Nitrate-Nitrogen, Total','Arsenic, Total','Lead, Total','Uranium, Total','Copper, Total','Iron, Total','Manganese, Total','Sodium, Total','Hardness by calculation','Antimony, Total','Cadmium, Total','Chromium, Total','pH Electrometric','Alkalinity','Sulfate','Total Coliform','E. Coli'],
  'Pro Plus':                        ['Chloride, Total','Fluoride, Total','Nitrite-Nitrogen, Total','Nitrate-Nitrogen, Total','Arsenic, Total','Lead, Total','Uranium, Total','Copper, Total','Iron, Total','Manganese, Total','Sodium, Total','Hardness by calculation','Alkalinity','Sulfate','Tannins','Total Dissolved Solids (TDS)','Bromide','pH Electrometric','Total Coliform','E. Coli'],
  'Arsenic, Speciation':             ['Arsenic, III', 'Arsenic, V', 'Arsenic, Total'],
  'Bacteria':                        ['Total Coliform','E. Coli'],
  'Radon Water':                     [],
};



const FHA_PARAM_NAMES  = ['Nitrite-Nitrogen, Total','Nitrate-Nitrogen, Total','Lead, Total','Total Coliform','E. Coli'];
const NEEDS_FHA_TYPES  = ['Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive'];

// ── Helpers ───────────────────────────────────────────────────────────────────

// Ensure time values have a colon: "1402" → "14:02", "14:02" → "14:02"
function ensureColon(dt) {
  if (!dt) return dt;
  const s = String(dt).trim();
  // Match "MM/DD/YY 1402" or "MM/DD/YY 1402" (no colon in time)
  return s.replace(/(\d{2}\/\d{2}\/\d{2})\s+(\d{2})(\d{2})$/, '$1 $2:$3')
          .replace(/(\d{1,2}\/\d{1,2}\/\d{2,4})\s+(\d{2})(\d{2})$/, '$1 $2:$3');
}
function formatCustomerName(name) {
  if (!name) return '';
  if (!name.startsWith('Public-')) return name;
  const inner    = name.slice('Public-'.length).trim();
  const commaIdx = inner.indexOf(',');
  if (commaIdx > 0) {
    const last  = inner.slice(0, commaIdx).trim();
    const first = inner.slice(commaIdx + 1).trim();
    return first ? `${first} ${last}` : last;
  }
  return inner;
}

function formatResult(rawVal, rl, decimals) {
  if (rawVal === '' || rawVal === null || rawVal === undefined) return '';
  const n = parseFloat(rawVal);
  if (isNaN(n)) return String(rawVal).trim();
  if (rl !== null && rl !== undefined && n < rl) return `<${rl}`;
  if (decimals !== undefined && decimals !== null) return n.toFixed(decimals);
  return parseFloat(n.toFixed(6)).toString();
}

function resultColor(paramName, displayVal, epa) {
  if (!displayVal && displayVal !== 0) return 'none';
  const s = String(displayVal);
  if (s.startsWith('<')) return 'green';
  if (paramName === 'pH Electrometric') {
    const n = parseFloat(s);
    return isNaN(n) ? 'none' : (n >= 6.5 && n <= 8.5) ? 'green' : 'red';
  }
  if (epa === null || epa === undefined || epa === '') return 'none'; // no EPA limit = no color indicator
  const n = parseFloat(s);
  if (isNaN(n)) return 'blue';
  return n <= parseFloat(epa) ? 'green' : 'red';
}

function combineDT(dateStr, timeStr) {
  if (!dateStr) return '';
  const ds = String(dateStr).trim();
  const ts = String(timeStr || '').trim();
  const dm = ds.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!dm) return ds;
  const m = dm[1].padStart(2,'0'), d = dm[2].padStart(2,'0');
  const y = dm[3].length === 4 ? dm[3].slice(-2) : dm[3].padStart(2,'0');
  let hhmm = '';
  if (ts) {
    const ampm = ts.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
    const h24  = ts.match(/^(\d{1,2}):(\d{2})/);
    if (ampm) {
      let h = parseInt(ampm[1]);
      if (ampm[3].toUpperCase() === 'PM' && h !== 12) h += 12;
      if (ampm[3].toUpperCase() === 'AM' && h === 12) h = 0;
      hhmm = `${String(h).padStart(2,'0')}:${ampm[2]}`;
    } else if (h24) {
      hhmm = `${String(parseInt(h24[1])).padStart(2,'0')}:${h24[2]}`;
    }
  }
  return hhmm ? `${m}/${d}/${y} ${hhmm}` : `${m}/${d}/${y}`;
}

// ── Get sample metadata from Archived Intake (field_X mapping) ────────────────
async function getSampleMeta(baseId) {
  try {
    const items = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
    // Find rows where field_1 (fullId) starts with baseId
    const matches = items.filter(r => (r.field_1 || '').startsWith(baseId));
    if (!matches.length) return null;

    // Merge data from all rows for this baseId
    const merged = {};
    for (const r of matches) {
      if (!merged.customer    && r.field_3)  merged.customer    = r.field_3;
      if (!merged.dateDrawn   && r.field_4)  merged.dateDrawn   = r.field_4;
      if (!merged.timeDrawn   && r.field_5)  merged.timeDrawn   = r.field_5;
      if (!merged.dateReceived && r.field_6) merged.dateReceived = r.field_6;
      if (!merged.timeReceived && r.field_7) merged.timeReceived = r.field_7;
      if (!merged.location    && r.field_8)  merged.location    = r.field_8;
      if (!merged.city        && r.field_9)  merged.city        = r.field_9;
      if (!merged.state       && r.field_10) merged.state       = r.field_10;
      if (!merged.zip         && r.field_11) merged.zip         = r.field_11;
      if (!merged.approvedBy  && r.field_12) merged.approvedBy  = r.field_12;
    }

    // Collect all test types
    const tests = [...new Set(matches.map(r => r.field_2).filter(Boolean))];

    return {
      customer:     merged.customer     || '',
      dateDrawn:    merged.dateDrawn    || '',
      timeDrawn:    merged.timeDrawn    || '',
      dateReceived: merged.dateReceived || '',
      timeReceived: merged.timeReceived || '',
      location:     merged.location     || '',
      city:         merged.city         || '',
      state:        merged.state        || 'ME',
      zip:          merged.zip          || '',
      approvedBy:   merged.approvedBy   || '',
      services:     tests.join(' | '),
    };
  } catch(e) { console.error('getSampleMeta:', e.message); return null; }
}

// ── Get Results Cache row for this base ID ────────────────────────────────────
async function getResultsCache(baseId) {
  try {
    const items = await listItems('Results Cache', { top: 500 });
    return items.find(r => {
      const stored = String(r.LabID || '').split(' ')[0].trim();
      return stored === baseId;
    }) || null;
  } catch(e) { console.error('getResultsCache:', e.message); return null; }
}

// ── Get client email from Clients list ────────────────────────────────────────
async function getClientInfo(customerName, token) {
  const empty = { clientName:'', mainContact:'', reportEmail:'', billingEmail:'', email:'', phone:'', dbaName:'', clientCode:'', abbrev:'', billingAddress:'', billingPreference:'', frequency:'', pricingCategory:'', startDate:'', status:'', radonLic:'' };
  if (!customerName) return empty;
  try {
    const siteId = process.env.SP_SITE_ID;
    const res    = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields($select=ClientName,ClientCode,Abbrev,Email,Phone,Active,Aliases,Notes,BillingAddress,BillingFrequency,Frequency,PricingCategory,StartDate,Status,RadonLic_x0023_)&$top=500`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return empty;
    const data  = await res.json();
    const name  = customerName.toLowerCase().trim();
    const formatted = formatCustomerName(customerName).toLowerCase().trim();
    const match = (data.value || []).find(item => {
      const f  = item.fields || {};
      // ClientName is the primary name field in the new client list structure
      const cn = (f.ClientName || f.Title || '').toLowerCase().trim();
      return cn === name || cn === formatted;
    });
    if (!match) return empty;
    const f = match.fields || {};
    // New Clients list field mapping (internal names differ from display names):
    // Aliases = Report Email, Notes = Billing Email, Active = Phone #,
    // Email = Main Contact (first name), Phone = DBA Name
    return {
      clientName:      f.ClientName      || f.Title   || '',
      mainContact:     f.Email           || '',          // display: "Main Contact"
      reportEmail:     f.Aliases         || '',          // display: "Report Email Address"
      billingEmail:    f.Notes           || '',          // display: "Billing Email Address"
      email:           f.Aliases         || '',          // alias for compatibility
      phone:           f.Active          || '',          // display: "Phone #"
      dbaName:         f.Phone           || '',          // display: "DBA Name"
      clientCode:      f.ClientCode      || '',
      abbrev:          f.Abbrev          || '',
      billingAddress:  f.BillingAddress  || '',
      billingPreference: f.BillingFrequency || '',
      frequency:       f.Frequency       || '',
      pricingCategory: f.PricingCategory || '',
      startDate:       f.StartDate       || '',
      status:          f.Status          || '',
      radonLic:        f.RadonLic_x0023_ || '',
    };
  } catch { return empty; }
}

// ── Main handler ──────────────────────────────────────────────────────────────
app.http('generate-report', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {
    try {
      const body = await request.json().catch(() => ({}));
      const { labId, meta: frontendMeta } = body;
      if (!labId) return { status:400, jsonBody:{ error:'labId required' } };

      const baseId = String(labId).match(/(\d{6}-\d{3})/)?.[1];
      if (!baseId)  return { status:400, jsonBody:{ error:`Invalid Lab ID: ${labId}` } };

      const log = [`Lab ID: ${baseId}`];
      const token = await getToken();

      // ── Fetch data in parallel ──────────────────────────────────────────────
      const [metaRaw, cache, clientRaw] = await Promise.all([
        frontendMeta ? Promise.resolve(null) : getSampleMeta(baseId),
        getResultsCache(baseId),
        getToken().then(t => getClientInfo(
          formatCustomerName(frontendMeta?.customer || ''),
          t
        )).catch(() => empty),
      ]);

      // ── Resolve meta ────────────────────────────────────────────────────────
      const meta = frontendMeta || metaRaw || {};
      log.push(`meta: ${meta.customer || 'unknown'}`);
      if (cache) {
        const filledFields = Object.keys(cache).filter(k => !k.startsWith('_') && !k.startsWith('@') && cache[k]);
        log.push(`cache: found — ${filledFields.length} fields: ${filledFields.slice(0,8).join(', ')}`);
      } else {
        log.push(`cache: NOT FOUND for baseId=${baseId}`);
      }

      // Look up client — try raw stored name first (e.g. "Public-Chandler, Zach")
      // then formatted name (e.g. "Zach Chandler")
      const rawName       = meta.customer || '';
      const formattedName = formatCustomerName(rawName);
      const clientInfo = await getClientInfo(rawName, token)
        .then(async c => {
          if (c.email) return c;
          if (rawName !== formattedName) return getClientInfo(formattedName, token).catch(() => c);
          return c;
        })
        .catch(() => empty);

      // ── Determine test types ────────────────────────────────────────────────
      // meta.services = "Alkalinity | pH" (string from Archived Intake)
      // frontendMeta.tests = ["Alkalinity | pH"] (array from accession-status)
      // Both need splitting on | to get individual service names
      // testOverride from dashboard takes precedence over accession-status data
      const testOverride = (body.testOverride || '').trim();
      const services = testOverride
        ? [testOverride]
        : meta.services
          ? meta.services.split(/[|;]/).map(s=>s.trim()).filter(Boolean)
          : (frontendMeta?.tests || []).flatMap(t =>
              String(t).split(/[|;]/).map(s=>s.trim()).filter(Boolean)
            );

      const isRadon      = services.some(s => /radon/i.test(s));

      // ── Fetch radon result from Results Cache if isRadon ─────────────────
      let radonResult = { display:'', raw:0, color:'green', time:'', date:'' };
      if (isRadon) {
        try {
          const rcAuthHdr = { Authorization: `Bearer ${token}` };
          const rcListSearchRes = await fetch(
            `${GRAPH}/sites/${siteId}/lists?$select=id,displayName`,
            { headers: rcAuthHdr }
          );
          if (rcListSearchRes.ok) {
            const rcListData = (await rcListSearchRes.json()).value || [];
            const rcListId   = rcListData.find(l=>l.displayName==='Results Cache')?.id;
            if (rcListId) {
              const rcItemRes = await fetch(
                `${GRAPH}/sites/${siteId}/lists/${rcListId}/items?$expand=fields($select=LabID,Radon,RadonDate,RadonTime)&$top=2000`,
                { headers: rcAuthHdr }
              );
              if (rcItemRes.ok) {
                const rcItems = (await rcItemRes.json()).value || [];
                const match   = rcItems.find(i => (i.fields?.LabID||'').trim() === baseId);
                if (match?.fields?.Radon) {
                  const val = parseFloat(match.fields.Radon) || 0;
                  radonResult = {
                    display: match.fields.Radon,
                    raw:     val,
                    color:   'green',
                    time:    match.fields.RadonTime || '',
                    date:    match.fields.RadonDate || '',
                  };
                  context.log(`[gen] Radon result for ${baseId}: ${match.fields.Radon} pCi/L`);
                }
              }
            }
          }
        } catch(e) { context.log('[gen] Radon fetch error:', e.message); }
      }
      const isArsenicSpec = services.some(s => /arsenic.*spec/i.test(s) || /spec.*arsenic/i.test(s));
      const needsFHA = services.some(s => NEEDS_FHA_TYPES.includes(s));

      // ── Load which parameters each test type includes from SharePoint ─────
      let testTypeElements = {};
      try {
        const token2  = await getToken();
        const siteId  = process.env.SP_SITE_ID;
        const GRAPH   = 'https://graph.microsoft.com/v1.0';
        const authHdr = { Authorization: `Bearer ${token2}` };

        // Find the list by display name (more reliable than URL encoding)
        const listSearchRes = await fetch(
          `${GRAPH}/sites/${siteId}/lists?$select=id,displayName&$top=100`,
          { headers: authHdr }
        );
        let listId = null;
        if (listSearchRes.ok) {
          const listData = await listSearchRes.json();
          const found = (listData.value || []).find(l =>
            l.displayName === 'Current Pricing-V1' ||
            l.displayName === 'Current Pricing V1' ||
            (l.displayName||'').toLowerCase().includes('current pricing')
          );
          listId = found?.id;
          context.log('[gen] Pricing list found:', found?.displayName, 'id:', listId);
        }

        if (listId) {
          const ttRes = await fetch(
            `${GRAPH}/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`,
            { headers: authHdr }
          );
          if (ttRes.ok) {
            const ttData = await ttRes.json();
            if (ttData.value?.[0]) {
              context.log('[gen] SP field keys:', JSON.stringify(Object.keys(ttData.value[0].fields || {})));
            }
            for (const item of (ttData.value || [])) {
              const f = item.fields || {};
              // Description column (try multiple internal names SP might use)
              const descVal = f.Description || f.Description0 || f.Elements || '';
              const titleVal = f.Title || f.Service || '';
              if (!titleVal || !descVal) continue;
              const rawElems = descVal
                .replace(/<[^>]+>/gi, '|')
                .split(/[\r\n|]+/)
                .map(s => s.trim().replace(/^,+|,+$/g, '').trim())
                .filter(Boolean);
              const resolved = [];
              for (const elem of rawElems) {
                const aliased = PARAM_SERVICE_ALIASES[elem];
                if (aliased) aliased.forEach(a => resolved.push(a));
                else resolved.push(elem);
              }
              if (resolved.length > 0) testTypeElements[titleVal] = resolved;
            }
            context.log('[gen] SP loaded', Object.keys(testTypeElements).length, 'entries');
          } else {
            context.log('[gen] Items fetch failed:', ttRes.status);
          }
        } else {
          context.log('[gen] Pricing list not found — using fallback only');
        }
      } catch(e) {
        context.log('[gen] SP error:', e.message);
      }

      // Build the set of parameter names needed for this report.
      // Priority:
      //   1. PARAM_SERVICE_ALIASES  — single element name mismatches (pH→pH Electrometric)
      //   2. SP list (Current Pricing-V1) — dynamic package coverage
      //   3. PACKAGE_COVERAGE_FALLBACK  — hardcoded fallback (always works)
      //   4. Direct name match — element name == PARAM_CONFIG name
      const needed = new Set();
      for (const svc of services) {
        const aliases = PARAM_SERVICE_ALIASES[svc];
        if (aliases) {
          // Single element alias (pH→pH Electrometric, Fluoride→Fluoride, Total, etc.)
          aliases.forEach(a => needed.add(a));
          context.log(`[gen] svc="${svc}" → alias [${aliases.join(', ')}]`);
        } else if (testTypeElements[svc] && testTypeElements[svc].length > 0) {
          // SP list has this package/service with elements defined
          testTypeElements[svc].forEach(e => needed.add(e));
          context.log(`[gen] svc="${svc}" → SP list (${testTypeElements[svc].length} params)`);
        } else if (PACKAGE_COVERAGE_FALLBACK[svc]) {
          // Hardcoded fallback for known packages
          PACKAGE_COVERAGE_FALLBACK[svc].forEach(e => needed.add(e));
          context.log(`[gen] svc="${svc}" → fallback (${PACKAGE_COVERAGE_FALLBACK[svc].length} params)`);
        } else {
          // Direct name match — element name already matches PARAM_CONFIG name
          needed.add(svc);
          context.log(`[gen] svc="${svc}" → direct`);
        }
      }

      const activeParams = PARAM_CONFIG.filter(p => needed.has(p.name));
      const fhaParams    = PARAM_CONFIG.filter(p => FHA_PARAM_NAMES.includes(p.name));
      context.log(`[gen] services=${JSON.stringify(services)}`);
      context.log(`[gen] needed=${JSON.stringify([...needed])}`);
      context.log(`[gen] activeParams=${activeParams.map(p=>p.name).join(' | ') || 'EMPTY'}`);
      context.log(`[gen] meta.services raw="${meta.services}"`);

      // ── Build param rows from Results Cache ─────────────────────────────────
      const c = cache || {};
      const acidPrepDT   = c.MetalsStartDate_x002f_Time || '';
      const icpmsAcqTime = c.AcquisitionTime || '';

      const buildRow = (p) => {
        let rawVal = '', analDT = '', prepDT = '';

        switch (p.source) {
          case 'gallery':
            rawVal = String(c[p.cacheField] || '');
            analDT = String(c[p.cacheDT]    || '');
            break;
          case 'icpms':
            rawVal = String(c[p.cacheField] || '');
            analDT = icpmsAcqTime;
            prepDT = acidPrepDT;
            break;
          case 'ph':
            rawVal = String(c.Title || ''); // Title = PH in Results Cache
            analDT = String(c.field_1 || '');
            break;
          case 'bac':
            rawVal = String(c[p.cacheField] || '');
            prepDT = String(c.field_4 || ''); // Start Date/Time bacteria
            analDT = String(c.field_5 || ''); // End Date/Time bacteria
            break;
          case 'calc': {
            const ca = parseFloat(c.Calcium_x0028_Ca43_x0029_  || '');
            const mg = parseFloat(c.Magnesium_x0028_Mg24_x0029_ || '');
            if (!isNaN(ca) && !isNaN(mg)) {
              rawVal = (Math.round((ca*2.497 + mg*4.118)*100)/100).toString();
            }
            analDT = icpmsAcqTime;
            prepDT = acidPrepDT;
            break;
          }
          case 'spec_iii':
            // Arsenic III — from ArsenicIII Results Cache field (As3 ICP-MS row)
            // Uses its own acquisition time from the As3 ICP-MS run
            rawVal = String(c.Arsenic3 || c.ArsenicIII || '');
            analDT = String(
              c.Arsenic3AcquisitionTime ||  // "ArsenicIII Acquisition Time"
              c.Arsenic3AcquisitionTime ||                 // fallback no-space version
              icpmsAcqTime || ''
            );
            prepDT = acidPrepDT;
            break;
          case 'spec_v': {
            // Arsenic V = Arsenic Total (TAs) minus Arsenic III (As3) — calculated
            const asTotal = parseFloat(c['Arsenic_x0028_As75_x0029_'] || '');
            const asIII   = parseFloat(c.Arsenic3 || c.ArsenicIII || '');
            if (!isNaN(asTotal) && !isNaN(asIII)) {
              rawVal = String(Math.max(0, Math.round((asTotal - asIII) * 10000) / 10000));
            }
            analDT = ''; // calculated — no instrument date
            prepDT = ''; // calculated — no prep date
            break;
          }
        }

        const display = formatResult(rawVal, p.rl, p.decimals);
        return {
          name:     p.name,
          value:    display,
          rl:       p.rl  !== null && p.rl  !== undefined ? String(p.rl)  : '',
          epa:      p.epa !== null && p.epa !== undefined ? String(p.epa) : '',
          unit:     p.unit,
          method:   p.method,
          prepDT:   ensureColon(prepDT),
          analDT:   ensureColon(analDT),
          time:     ensureColon(analDT),
          color:    resultColor(p.name, display, p.epa),
          source:   p.source,
        };
      };

      const paramRows = activeParams.map(buildRow);
      const fhaRows   = needsFHA ? fhaParams.map(buildRow) : [];

      // ── Dates ───────────────────────────────────────────────────────────────
      const now        = new Date();
      const todayStr   = `${String(now.getMonth()+1).padStart(2,'0')}/${String(now.getDate()).padStart(2,'0')}/${String(now.getFullYear()).slice(-2)}`;
      const dtCollected = meta.dateDrawn    ? combineDT(meta.dateDrawn,    meta.timeDrawn)    : '';
      const dtReceived  = meta.dateReceived ? combineDT(meta.dateReceived, meta.timeReceived) : '';

      return {
        status: 200,
        jsonBody: {
          success:    true,
          labId:      baseId,
          isRadon,
        isArsenicSpec,
          needsFHA,
          reportType: isRadon ? 'RW' : 'COA',
          today:      todayStr,
          log,
          meta: {
            customer:     formatCustomerName(meta.customer || ''),
            // clientName from Clients list takes priority over raw intake customer name
            clientName:   clientInfo.clientName || formatCustomerName(meta.customer || ''),
            email:            clientInfo.email            || '',
            reportEmail:      clientInfo.reportEmail      || '',
            billingEmail:     clientInfo.billingEmail     || '',
            phone:            clientInfo.phone            || '',
            clientCode:       clientInfo.clientCode       || '',
            abbrev:           clientInfo.abbrev           || '',
            billingAddress:   clientInfo.billingAddress   || '',
            billingPreference: clientInfo.billingPreference || '',
            pricingCategory:  clientInfo.pricingCategory  || '',
            dbaName:          clientInfo.dbaName          || '',
            location:     meta.location     || '',
            city:         meta.city         || '',
            state:        meta.state        || 'ME',
            zip:          meta.zip          || '',
            dtCollected,
            dtReceived,
            dateDrawn:    meta.dateDrawn    || '',
            timeDrawn:    meta.timeDrawn    || '',
            dateReceived: meta.dateReceived || '',
            timeReceived: meta.timeReceived || '',
            labId:        baseId,
          },
          services,
          paramRows,
          fhaRows,
          radon: radonResult,
        },
      };

    } catch(err) {
      context.log(`[generate-report] fatal: ${err.message}`);
      return { status:500, jsonBody:{ error: err.message } };
    }
  },
});
