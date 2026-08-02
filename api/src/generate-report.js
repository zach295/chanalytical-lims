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
  if (epa === null || epa === undefined || epa === '') return 'blue';
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
  const empty = { email:'', phone:'', clientCode:'', abbrev:'' };
  if (!customerName) return empty;
  try {
    const siteId = process.env.SP_SITE_ID;
    const res    = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields($select=Title,ClientCode,Abbrev,Email,Aliases,Phone)&$top=500`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return empty;
    const data  = await res.json();
    const name  = customerName.toLowerCase().trim();
    const formatted = formatCustomerName(customerName).toLowerCase().trim();
    const match = (data.value || []).find(item => {
      const f  = item.fields || {};
      const cn = (f.Title || '').toLowerCase().trim();
      const al = (f.Aliases || '').toLowerCase();
      return cn === name || cn === formatted
        || al.split(/[,;]/).map(s => s.trim().toLowerCase()).some(a => a && (a === name || a === formatted));
    });
    if (!match) return empty;
    const f = match.fields || {};
    return { email: f.Email||'', clientCode: f.ClientCode||'', abbrev: f.Abbrev||'', phone: f.Phone||'' };
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
        )).catch(() => ({ email:'', phone:'', clientCode:'', abbrev:'' })),
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
        .catch(() => ({ email:'', phone:'', clientCode:'', abbrev:'' }));

      // ── Determine test types ────────────────────────────────────────────────
      const services = meta.services
        ? meta.services.split(/[|;]/).map(s=>s.trim()).filter(Boolean)
        : (frontendMeta?.tests || []);

      const isRadon  = services.some(s => /radon/i.test(s));
      const needsFHA = services.some(s => NEEDS_FHA_TYPES.includes(s));

      // ── Load which parameters each test type includes from SharePoint ─────
      let testTypeElements = {};
      try {
        const token2 = await getToken();
        const siteId = process.env.SP_SITE_ID;
        const ttRes  = await fetch(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/Current%20Pricing-V1/items?$expand=fields($select=Title,Elements,Active)&$top=200`,
          { headers: { Authorization: `Bearer ${token2}` } }
        );
        if (ttRes.ok) {
          const ttData = await ttRes.json();
          for (const item of (ttData.value || [])) {
            const f = item.fields || {};
            if (f.Active === false || f.Active === 'FALSE') continue;
            if (f.Title && f.Elements) {
              testTypeElements[f.Title] = f.Elements
                .split('|').map(s => s.trim()).filter(Boolean);
            }
          }
        }
      } catch(e) { /* fall through to direct name matching below */ }

      // Build the set of parameter names needed for this report
      const needed = new Set();
      for (const svc of services) {
        if (testTypeElements[svc]) {
          // Package or element test with SP definition — use its element list
          testTypeElements[svc].forEach(e => needed.add(e));
        } else {
          // No SP entry — use alias map first, then direct name match
          const aliases = PARAM_SERVICE_ALIASES[svc];
          if (aliases) {
            aliases.forEach(a => needed.add(a));
          } else {
            needed.add(svc); // direct match (service name == param name)
          }
        }
      }

      const activeParams = PARAM_CONFIG.filter(p => needed.has(p.name));
      const fhaParams    = PARAM_CONFIG.filter(p => FHA_PARAM_NAMES.includes(p.name));
      context.log(`[generate-report] services=${JSON.stringify(services)} needed=${JSON.stringify([...needed])} activeParams=${activeParams.map(p=>p.name).join('|')}`);

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
          needsFHA,
          reportType: isRadon ? 'RW' : 'COA',
          today:      todayStr,
          log,
          meta: {
            customer:     formatCustomerName(meta.customer || ''),
            email:        clientInfo.email       || '',
            phone:        clientInfo.phone        || '',
            clientCode:   clientInfo.clientCode   || '',
            abbrev:       clientInfo.abbrev       || '',
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
          radon: { display:'', raw:0, color:'green', time:'' },
        },
      };

    } catch(err) {
      context.log(`[generate-report] fatal: ${err.message}`);
      return { status:500, jsonBody:{ error: err.message } };
    }
  },
});
