// api/src/scan-folder.js
// Azure version — reads COC PDFs from SharePoint instead of Google Drive.
// Writes scan results to SharePoint "Review Queue" list via graph.js.
// Everything else (Azure Doc Intel OCR, Claude Sonnet extraction, all parsing)
// is identical to the Netlify version.

const { app } = require('@azure/functions');
const {
  getToken,
  spListGet,
  spListCreate,
  spListQuery,
} = require('../shared/graph');

// ── ET time helpers (inline — no external dependency needed) ──────────────────
const TZ = 'America/New_York';
function etParts(date) {
  const d = date || new Date();
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: TZ, year:'numeric', month:'2-digit', day:'2-digit',
    hour:'2-digit', minute:'2-digit', second:'2-digit', hour12:false,
  }).formatToParts(d);
  const get = (type) => parts.find(p => p.type === type)?.value || '';
  const hour = get('hour') === '24' ? '00' : get('hour');
  return { year:get('year'), month:get('month'), day:get('day'), hour, minute:get('minute') };
}
function todayDisplayET() { const p = etParts(); return `${p.month}-${p.day}-${p.year.slice(-2)}`; }
function nowTimeET() { return new Date().toLocaleTimeString('en-US',{timeZone:TZ,hour:'2-digit',minute:'2-digit',hour12:true}); }

// ── MS Graph constants ────────────────────────────────────────────────────────
const GRAPH = 'https://graph.microsoft.com/v1.0';

// ── SharePoint file operations ────────────────────────────────────────────────

// Convert full SP path → Graph drive-relative path (URL-encoded)
function toDrivePath(fullPath) {
  const marker = 'Shared Documents/';
  const idx = fullPath.indexOf(marker);
  const rel = idx >= 0 ? fullPath.slice(idx + marker.length) : fullPath.replace(/^\/+/, '');
  return rel.split('/').map(s => encodeURIComponent(s)).join('/');
}

// List files in a SharePoint folder via Graph
async function listSpFolder(folderPath, token) {
  const siteId    = process.env.SP_SITE_ID;
  const drivePath = toDrivePath(folderPath);
  const res = await fetch(
    `${GRAPH}/sites/${siteId}/drive/root:/${drivePath}:/children?$select=id,name,file,createdDateTime`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`listSpFolder ${folderPath}: ${res.status} ${err}`);
  }
  const data = await res.json();
  return data.value || [];
}

// Download a SharePoint file by item ID → Buffer
async function downloadSpFile(itemId, token) {
  const siteId = process.env.SP_SITE_ID;
  const res = await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${itemId}/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`downloadSpFile ${itemId}: ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

// Move a SharePoint file to a different folder by updating its parent
async function moveSpFile(itemId, destFolderPath, token) {
  const siteId    = process.env.SP_SITE_ID;
  const drivePath = toDrivePath(destFolderPath);

  // Resolve destination folder ID
  const folderRes = await fetch(
    `${GRAPH}/sites/${siteId}/drive/root:/${drivePath}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!folderRes.ok) {
    console.warn(`[moveSpFile] Cannot resolve destination ${destFolderPath}: ${folderRes.status}`);
    return;
  }
  const folderData = await folderRes.json();
  const destId     = folderData.id;

  const patchRes = await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${itemId}`,
    {
      method:  'PATCH',
      headers: { Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
      body:    JSON.stringify({ parentReference: { id: destId } }),
    }
  );
  if (!patchRes.ok) {
    const err = await patchRes.text();
    console.warn(`[moveSpFile] Move failed for ${itemId}: ${patchRes.status} ${err}`);
  }
}

// ── SharePoint List helpers ───────────────────────────────────────────────────

// Load clients from SharePoint "Clients" list
async function loadClients(token) {
  try {
    const siteId = process.env.SP_SITE_ID;
    const res    = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields($select=Title,ClientCode,Abbrev,Email,Aliases)&$top=500`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return [];
    const data = await res.json();
    return (data.value || []).map(item => ({
      clientName: item.fields?.Title      || '',
      clientCode: item.fields?.ClientCode || '',
      abbrev:     item.fields?.Abbrev     || '',
      email:      item.fields?.Email      || '',
      aliases:    item.fields?.Aliases    || '',
    }));
  } catch { return []; }
}

// Look up a barcode ID in the "Archived Intake" SP list
async function lookupBarcode(barcodeId, token) {
  if (!barcodeId || !barcodeId.startsWith('CHA-')) return null;
  try {
    const siteId = process.env.SP_SITE_ID;
    const filter = encodeURIComponent(`fields/BarcodeID eq '${barcodeId}'`);
    const res    = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Archived Intake/items?$filter=${filter}&$expand=fields&$top=1`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return null;
    const data = await res.json();
    if (!data.value?.length) return null;
    const f = data.value[0].fields;
    return {
      barcodeId,
      customer: f.ClientName || '',
      email:    f.Email      || '',
      location: f.Address    || '',
      city:     f.City       || '',
      state:    f.State      || 'ME',
      zip:      f.Zip        || '',
      services: f.Services   || '',
    };
  } catch { return null; }
}

// Get file IDs already in the Review Queue to avoid re-processing
async function getQueuedFileIds(token) {
  try {
    const siteId = process.env.SP_SITE_ID;
    const res    = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Review Queue/items?$expand=fields($select=FileID)&$top=2000`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return new Set();
    const data = await res.json();
    const ids  = new Set();
    (data.value || []).forEach(item => {
      const fid = item.fields?.FileID;
      if (fid) ids.add(String(fid).trim());
    });
    return ids;
  } catch { return new Set(); }
}

// Write one scan result row to the "Review Queue" SP list
async function writeToReviewQueue(fields, token) {
  const siteId = process.env.SP_SITE_ID;
  const res    = await fetch(
    `${GRAPH}/sites/${siteId}/lists/Review Queue/items`,
    {
      method:  'POST',
      headers: { Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
      body:    JSON.stringify({ fields }),
    }
  );
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`writeToReviewQueue: ${res.status} ${err}`);
  }
  return res.json();
}

// ── Client matching ───────────────────────────────────────────────────────────
function matchClient(name, clients) {
  if (!name || !clients.length) return null;
  const s = name.toLowerCase().trim();
  if (s.length < 3) return null;
  return (
    clients.find(c => c.clientName.toLowerCase() === s) ||
    clients.find(c => c.aliases.split(',').map(a => a.trim().toLowerCase())
      .some(a => a.length >= 4 && (s === a || s.includes(a) || a.includes(s)))) ||
    null
  );
}

// ── Test validation ───────────────────────────────────────────────────────────
function validateTests(tests) {
  const pkgs = ['Basic Safety (FHA)','Basic Safety','Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive','Pro Plus'];
  const addons = ['Radon Water'];
  const hasPkg = tests.some(t => pkgs.includes(t));
  const hasInd = tests.some(t => !pkgs.includes(t) && !addons.includes(t));
  return hasPkg && hasInd ? 'Package tests cannot be combined with individual elements' : '';
}

// ── Case normalization ────────────────────────────────────────────────────────
function normalizeCase(s) {
  if (!s || s.length <= 2) return s;
  const letters = s.replace(/[^a-zA-Z]/g, '');
  if (!letters.length) return s;
  const upperRatio = (s.match(/[A-Z]/g) || []).length / letters.length;
  if (upperRatio > 0.8) return s.toLowerCase().replace(/\b([a-z])/g, c => c.toUpperCase());
  return s;
}

// ── Address cleanup ───────────────────────────────────────────────────────────
function cleanAddress(location) {
  if (!location) return location;
  let loc = location.replace(/[.,]+/g, ' ').replace(/\s+/g, ' ').trim();
  loc = loc.replace(/\bwas\b/gi, 'Way').replace(/\s+/g, ' ').trim();
  const abbrevMap = {
    '\\bRoad\\b':'Rd','\\bStreet\\b':'St','\\bAvenue\\b':'Ave','\\bDrive\\b':'Dr',
    '\\bLane\\b':'Ln','\\bTrail\\b':'Trl','\\bBoulevard\\b':'Blvd','\\bCourt\\b':'Ct',
    '\\bPlace\\b':'Pl','\\bCircle\\b':'Cir','\\bHighway\\b':'Hwy','\\bParkway\\b':'Pkwy',
    '\\bRoute\\b':'Rte','\\bTerrace\\b':'Ter','\\bExtension\\b':'Ext',
  };
  for (const [pattern, abbrev] of Object.entries(abbrevMap)) {
    loc = loc.replace(new RegExp(pattern, 'gi'), abbrev);
  }
  return loc;
}

// ── ZIP correction ────────────────────────────────────────────────────────────
function fixZip(zip, state) {
  if (!zip) return '';
  const z = String(zip).replace(/\D/g, '').padStart(5, '0');
  if ((state === 'ME' || state === 'NH') && !z.startsWith('0')) {
    return '0' + z.slice(1);
  }
  return z;
}

// ── Main Azure Function handler ───────────────────────────────────────────────
app.http('scan-folder', {
  methods:   ['GET', 'POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {

    if (!process.env.ANTHROPIC_API_KEY) {
      return { status: 500, jsonBody: { error: 'ANTHROPIC_API_KEY not set' } };
    }

    const SCAN_INCOMING = process.env.SP_SCAN_FOLDER || '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Incoming';
    const SCAN_REVIEW   = process.env.SP_SCAN_REVIEW  || '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Review';
    const SCAN_ARCHIVE  = process.env.SP_SCAN_ARCHIVE || '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Archived';

    try {
      const body          = await request.json().catch(() => ({}));
      const scannedByName = body.scannedBy || 'Lab Staff';

      // Shared MS Graph token (cached in graph.js)
      const token = await getToken();

      // Load clients for alias matching and OCR prompt context
      const clients  = await loadClients(token);
      const aliasCtx = clients.map(c =>
        `- "${c.clientName}"${c.aliases ? ` (aliases: ${c.aliases})` : ''}`
      ).join('\n') +
      '\n⚠️ IMPORTANT: Only match to a client above if their name or alias is EXPLICITLY WRITTEN on this form. Do NOT use your training knowledge to guess which client submitted it. If nothing is written, return "".';

      // List files in INCOMING folder (PDF and image files only)
      const allFiles = await listSpFolder(SCAN_INCOMING, token);
      const files    = allFiles.filter(f =>
        f.file && (
          f.file.mimeType?.includes('image/') ||
          f.file.mimeType === 'application/pdf' ||
          /\.(pdf|jpg|jpeg|png|tiff?)$/i.test(f.name)
        )
      );

      if (!files.length) {
        return { status: 200, jsonBody: { checked: 0, processed: 0, message: 'No files in the INCOMING folder' } };
      }

      // Skip files already in Review Queue
      const queuedIds = await getQueuedFileIds(token);
      const toProcess = files.filter(f => !queuedIds.has(f.id));

      if (!toProcess.length) {
        return { status: 200, jsonBody: { checked: files.length, processed: 0, message: `${files.length} file(s) already in Review Queue` } };
      }

      const results = [];

      for (const file of toProcess) {
        try {
          // Move to REVIEW immediately to prevent duplicate processing
          await moveSpFile(file.id, SCAN_REVIEW, token);

          // Download file as Buffer → base64 for Azure Doc Intel
          const buf  = await downloadSpFile(file.id, token);
          const b64  = buf.toString('base64');
          const isPdf = /\.pdf$/i.test(file.name) || file.file?.mimeType === 'application/pdf';

          // ── PRIMARY: Azure Document Intelligence + Claude Sonnet ─────────────
          let raw       = '';
          let azureText = '';

          const azureEndpoint = process.env.AZURE_DOC_INTEL_ENDPOINT;
          const azureKey      = process.env.AZURE_DOC_INTEL_KEY;

          if (azureEndpoint && azureKey) {
            const endpoint = azureEndpoint.replace(/\/+$/, '');
            context.log(`[scan] Azure endpoint: ${endpoint.slice(0, 50)}`);

            try {
              // Start Azure analysis
              const analyzeUrl = `${endpoint}/documentintelligence/documentModels/prebuilt-layout:analyze?api-version=2024-11-30`;
              const startRes   = await fetch(analyzeUrl, {
                method:  'POST',
                headers: { 'Ocp-Apim-Subscription-Key': azureKey, 'Content-Type': 'application/json' },
                body:    JSON.stringify({ base64Source: b64 }),
              });
              if (!startRes.ok) throw new Error(`Azure start: ${startRes.status} ${await startRes.text()}`);

              const operationUrl = startRes.headers.get('Operation-Location');
              if (!operationUrl) throw new Error('Azure returned no Operation-Location');

              // Poll until complete
              let azureResult;
              await new Promise(r => setTimeout(r, 2000));
              for (let i = 0; i < 12; i++) {
                const pollRes = await fetch(operationUrl, { headers: { 'Ocp-Apim-Subscription-Key': azureKey } });
                azureResult   = await pollRes.json();
                if (azureResult.status === 'succeeded' || azureResult.status === 'failed') break;
                await new Promise(r => setTimeout(r, 1500));
              }
              if (!azureResult || azureResult.status !== 'succeeded') {
                throw new Error(`Azure: ${azureResult?.status || 'timeout'}`);
              }

              // Build structured plain text from Azure output (page 1 only)
              const page1      = azureResult.analyzeResult?.pages?.[0];
              const pageHeight = page1?.height || 792;
              const pageWidth  = page1?.width  || 612;

              const BACK_PAGE_KEYWORDS = [
                'sample collection instructions','dropbox locations','payment information',
                'cardholder','whenever possible, collect sample from a faucet',
                'mastercard','cvv code','there is a 4% tech fee',
              ];
              const isBackPage = text => {
                const t = text.toLowerCase();
                return BACK_PAGE_KEYWORDS.some(k => t.includes(k));
              };

              const paragraphs = (azureResult.analyzeResult?.paragraphs || [])
                .filter(p => p.content && !isBackPage(p.content));

              const topSection = [], middleSection = [], bottomSection = [];
              for (const para of paragraphs) {
                if (!para.content) continue;
                const y       = para.boundingRegions?.[0]?.polygon?.[1] ?? 0;
                const normalY = y / pageHeight;
                const line    = para.content
                  .replace(/:selected:/g,   '[CHECKED]')
                  .replace(/:unselected:/g, '[unchecked]');
                if (normalY < 0.35)      topSection.push(line);
                else if (normalY < 0.75) middleSection.push(line);
                else                     bottomSection.push(line);
              }

              azureText = '';
              if (topSection.length)    azureText += '=== TOP OF FORM (Lab Use Only, Report To, Header) ===\n'                   + topSection.join('\n')    + '\n\n';
              if (middleSection.length) azureText += '=== MIDDLE OF FORM (Well Owner Address, Date/Time Sampled) ===\n'          + middleSection.join('\n') + '\n\n';
              if (bottomSection.length) azureText += '=== BOTTOM OF FORM (Test Type Checkboxes, Individual Elements) ===\n'      + bottomSection.join('\n') + '\n\n';

              const kvPairs = (azureResult.analyzeResult?.keyValuePairs || [])
                .filter(kv => !isBackPage(kv.key?.content || ''));
              if (kvPairs.length) {
                azureText += '=== DETECTED FORM FIELDS ===\n';
                for (const kv of kvPairs) {
                  const k = kv.key?.content || '';
                  const v = kv.value?.content || '';
                  if (k) azureText += `${k}: ${v}\n`;
                }
                azureText += '\n';
              }

              context.log(`[scan] Azure: ${paragraphs.length} paragraphs, ${kvPairs.length} kv pairs`);

              // Claude Sonnet structures Azure's text into JSON
              const extractRes = await fetch('https://api.anthropic.com/v1/messages', {
                method:  'POST',
                headers: {
                  'Content-Type':      'application/json',
                  'x-api-key':         process.env.ANTHROPIC_API_KEY,
                  'anthropic-version': '2023-06-01',
                },
                body: JSON.stringify({
                  model:      'claude-sonnet-4-6',
                  max_tokens: 600,
                  system:     'You are a JSON extraction API. Output ONLY a valid JSON object. No markdown, no explanation.',
                  messages: [{ role: 'user', content:
`Extract structured data from this A-Z Water Systems Chain of Custody form.
Azure Document Intelligence has already read the text. [CHECKED] = checked box. [unchecked] = unchecked.

FORM TEXT:
${azureText}

KNOWN CLIENTS (match customer name exactly if possible — if unsure, return ""):
${aliasCtx}

EXTRACTION RULES:

FIRST — determine form type:
- BUSINESS form: "Report To Be Sent To" has a company name that is NOT Chanalytical Laboratories (e.g. A-Z Water Systems, FPI, Madden Home Inspections). The well owner section has the property address.
- PUBLIC form: "Report To Be Sent To" shows "Chanalytical Laboratories" or is the lab's own address. The form has a "CUSTOMER & PROPERTY INFORMATION" section with the submitter's personal name, address, email, and phone.

- FORM TYPE DETECTION: Check the TOP OF FORM. If "Chanalytical Laboratories" appears as the printed header/logo AND the Report To section appears to be Chanalytical's own address (not a business client), this is a PUBLIC form submitted by an individual. Otherwise it is a BUSINESS form.

- customer:
  • PUBLIC FORM: Use the person's full name written in the "Name:" field of the "CUSTOMER & PROPERTY INFORMATION" section. This will be a person's name (e.g. "Travis Guld") NOT a company — that is correct. Return it as-is. If truly blank → return "".
  • BUSINESS FORM: The "REPORT TO BE SENT TO:" section may contain a checkbox list of multiple company names. ⛔ NEVER use the first name you see — you MUST find which one is CHECKED, MARKED, CIRCLED, or has an X/checkmark next to it. Only the marked company is the customer. If it is a fill-in line (not a checkbox list), copy exactly what is handwritten or typed. ⛔ If the section is BLANK, EMPTY, or nothing is written or marked → return "" — NEVER guess or fill in a company name. ⛔ NEVER substitute or replace what is marked/written with any other company name. ⛔ NEVER invent a company name. If nothing is marked or written → return "". ⛔ Never use form titles or guess.
- location: Street address — for public forms use the Address field in CUSTOMER & PROPERTY INFORMATION (TOP section); for business forms use the Well Owner section (MIDDLE section). Include ALL lines under the "Address:" label until the next label (City:, State:, etc.) even if they look like partial words — e.g. "was" after a street name is likely "Way" misread by OCR, include it. No periods or commas. If a circled T or ⊕ symbol appears → set waterType to "Treated". ⛔ Never use Report To address.
- city/state/zip: from MIDDLE OF FORM (Well Owner) ONLY. Maine zip starts with 04, NH starts with 03.
- email: For PUBLIC forms, extract from "E-mail:" field in CUSTOMER & PROPERTY INFORMATION section. The email may span TWO lines (e.g. "travis.john.gould" on one line and "@gmail.com" on next) — concatenate them into one address with no spaces. If not a public form or email is blank → return "".
- dateDrawn: Date CLIENT collected sample — next to "Date Sampled:" label → YYYY-MM-DD. ⛔ If blank or not written on the form return "" — NEVER guess, infer, or use today's date. ⛔ If date is crossed out/scribed through, return "". 2-digit years: 26=2026, 25=2025.
- timeDrawn: Time CLIENT collected sample — next to "Time Sampled:" label → HH:MM 24-hour. IMPORTANT: look carefully for this — it is often written as "9:59 AM", "9:59 A", "9:59a", "3:50 p", "3:50pm". "a"=AM, "p"=PM. Convert to 24hr (3:50 PM = 15:50, 9:59 AM = 09:59). ⛔ If time is crossed out, ignore it. Digit confusion: "0" vs "1", "5" vs "6". Minutes must be 00-59.
- receivedDate: Date LAB received sample — in the small "Lab Use Only" box (upper right corner). Stamped or written by lab staff. Looks like "JUN 24" or "07/14/26". Year always 2026 → YYYY-MM-DD.
- receivedTime: Time LAB received sample — in same "Lab Use Only" box. Looks like "14:06". DIFFERENT from Time Sampled. Minutes 00-59 only.
- tests: Package names where [CHECKED] appears in the right-side TEST TYPE column. Valid: "Basic Safety (FHA)","Standard Safety","Expanded Safety (Mortgage Test)","WW - Expanded Safety","Comprehensive","Pro Plus"
- hasRadon: true ONLY if [CHECKED] next to "Radon Water" in the TEST TYPE column. ⛔ NOT the "Radon Water Mitigation" note checkbox.
- individualElements: Elements where [CHECKED] appears in the LEFT column of individual tests. Look carefully — these are separate from package tests. Valid names include: Alkalinity, Arsenic Total, Bacteria, Cadmium Total, Calcium Total, Chloride Total, Copper Total, Fluoride, Hardness Total, Iron Total, Lead Total, Magnesium Total, Manganese Total, Nitrate, Nitrite, pH, Sodium Total, Sulfate, Sulfur, Tannins, Total Dissolved Solids (TDS), Uranium Total. Include ALL that have [CHECKED]. Note: "TDS" or "Total Dissolved Solids" on the form = "Total Dissolved Solids (TDS)".
- waterType: "Raw" or "Treated" if mentioned, else ""
- notes: observations, illegible fields. Note "Public submission" for public forms.
- barcodeId: barcode number like 0600326-006 or CHA-YYMMDD-####, else ""
- confidence: 0-100

Return ONLY: {"barcodeId":"","customer":"","email":"","dateDrawn":"","timeDrawn":"","receivedDate":"","receivedTime":"","location":"","city":"","state":"ME","zip":"","tests":[],"individualElements":[],"hasRadon":false,"notes":"","waterType":"","confidence":0}`
                  }],
                }),
              });
              if (!extractRes.ok) throw new Error(`Claude extract: ${extractRes.status}`);
              const extractData = await extractRes.json();
              raw = extractData.content?.find(c => c.type === 'text')?.text || '';
              context.log(`[scan] Claude extracted: ${raw.slice(0, 300)}`);

            } catch (azureErr) {
              context.log(`[scan] Azure hybrid failed: ${azureErr.message}`);
              results.push({ azureError: azureErr.message });
            }
          }

          // ── FALLBACK: Claude Sonnet text-only (if Azure not configured or failed) ──
          if (!raw) {
            const step2Res = await fetch('https://api.anthropic.com/v1/messages', {
              method:  'POST',
              headers: {
                'Content-Type':      'application/json',
                'x-api-key':         process.env.ANTHROPIC_API_KEY,
                'anthropic-version': '2023-06-01',
              },
              body: JSON.stringify({
                model:      'claude-sonnet-4-6',
                max_tokens: 1000,
                system:     'You are a JSON extraction API. Output ONLY a valid JSON object. No markdown, no explanation.',
                messages: [{ role:'user', content:
`Extract from this water testing COC form text. Known business clients: ${aliasCtx}

TWO FORM TYPES EXIST:
- BUSINESS form: Report To section has a company name or checkbox list of companies.
- PUBLIC form: Report To section is blank or says "Chanalytical Laboratories". Has CUSTOMER & PROPERTY INFORMATION with person's name/address/email/phone.

RULES:
- customer: PUBLIC form: person's name from CUSTOMER & PROPERTY INFORMATION. BUSINESS form: checked company in Report To. Return "" if nothing clear.
- location/city/state/zip: BUSINESS=from WELL OWNER section. PUBLIC=from CUSTOMER & PROPERTY INFORMATION.
- dateDrawn: YYYY-MM-DD from "Date Sampled" field. Return "" if blank — NEVER guess.
- timeDrawn: HH:MM 24hr from "Time Sampled". Return "" if blank.
- receivedDate/receivedTime: from "Lab Use Only" box only. Return "" if blank.
- barcodeId: alphanumeric code in Lab Use Only box. Return "" if absent.
- tests: only package tests with [CHECKED] mark.
- individualElements: only individual test rows with [CHECKED] mark.
- hasRadon: true only if Radon Water is [CHECKED].
- confidence: 0-100.

COC TEXT:
${azureText}

Return ONLY: {"barcodeId":"","customer":"","email":"","dateDrawn":"","timeDrawn":"","receivedDate":"","receivedTime":"","location":"","city":"","state":"ME","zip":"","tests":[],"individualElements":[],"hasRadon":false,"notes":"","waterType":"","confidence":0}`
                }],
              }),
            });
            if (!step2Res.ok) throw new Error(`Claude fallback: ${step2Res.status}`);
            raw = (await step2Res.json()).content?.find(c => c.type === 'text')?.text || '';
          }

          // ── Parse JSON from OCR output ────────────────────────────────────────
          let ocr;
          try {
            const s = raw.indexOf('{'), e = raw.lastIndexOf('}');
            if (s < 0 || e < 0) throw new Error('no JSON braces');
            ocr = JSON.parse(raw.slice(s, e + 1));
          } catch {
            throw new Error(`OCR JSON parse failed: ${raw.slice(0, 200)}`);
          }

          if (!ocr.customer && !ocr.location && !ocr.dateDrawn && !ocr.tests?.length) {
            context.log(`[scan] OCR returned empty for ${file.name} — writing to queue for manual review`);
            ocr.confidence = 0;
          }

          // ── Normalize and clean ───────────────────────────────────────────────
          ocr.customer = normalizeCase(ocr.customer);
          ocr.location = normalizeCase(ocr.location);
          ocr.city     = normalizeCase(ocr.city);
          ocr.location = cleanAddress(ocr.location);
          ocr.zip      = fixZip(ocr.zip, ocr.state);

          // ── Hallucination check ───────────────────────────────────────────────
          let validatedCustomer = ocr.customer || '';
          if (validatedCustomer && azureText) {
            const customerWords = validatedCustomer.toLowerCase()
              .replace(/[^a-z0-9\s]/g, '').split(/\s+/)
              .filter(w => w.length >= 4);
            const azureLower  = azureText.toLowerCase();
            const foundInText = customerWords.length === 0 || customerWords.some(w => azureLower.includes(w));
            if (!foundInText) {
              context.log(`[scan] Customer hallucination detected: "${validatedCustomer}" not in Azure text — clearing`);
              validatedCustomer = '';
              ocr.customer      = '';
            }
          }

          // ── Barcode lookup ────────────────────────────────────────────────────
          let barcodeMatch = null;
          let reviewStatus = 'Ready to Review';
          let ocrStatus    = 'Complete';

          if (ocr.barcodeId?.startsWith('CHA-')) {
            barcodeMatch = await lookupBarcode(ocr.barcodeId, token);
            if (barcodeMatch) {
              ocrStatus = 'Barcode Match';
              if (!ocr.customer && barcodeMatch.customer) ocr.customer = barcodeMatch.customer;
              if (!ocr.email    && barcodeMatch.email)    ocr.email    = barcodeMatch.email;
              if (!ocr.location && barcodeMatch.location) ocr.location = barcodeMatch.location;
              if (!ocr.city     && barcodeMatch.city)     ocr.city     = barcodeMatch.city;
              if (!ocr.state    && barcodeMatch.state)    ocr.state    = barcodeMatch.state;
              if (!ocr.zip      && barcodeMatch.zip)      ocr.zip      = barcodeMatch.zip;
              if (!ocr.tests?.length && barcodeMatch.services) {
                ocr.tests = barcodeMatch.services.split(';').map(s => s.trim()).filter(Boolean);
              }
            }
          }

          // ── Client matching ───────────────────────────────────────────────────
          const client     = matchClient(validatedCustomer, clients);
          const clientName = client?.clientName || validatedCustomer || '';

          // ── AIO conversion ────────────────────────────────────────────────────
          const isAIO = clientName.toLowerCase().includes('all in one') || clientName.toLowerCase().includes('aio');
          if (isAIO) {
            ocr.tests = (ocr.tests || []).map(t => {
              if (/^basic safety(\s*\(fha\))?$/i.test(t)) return 'AIO FHA';
              if (/portability/i.test(t))                  return 'AIO Portability';
              return t;
            });
          }

          const pkgTests    = ocr.tests || [];
          const indElements = ocr.individualElements || [];
          const radonTests  = ocr.hasRadon ? ['Radon Water'] : [];
          const tests       = [...pkgTests, ...indElements, ...radonTests];
          const valError    = validateTests(pkgTests);
          if (valError) reviewStatus = 'Validation Error';

          const stamp            = `${todayDisplayET()} ${nowTimeET()}`;
          const waterTypeSuffix  = ocr.waterType ? ` - ${ocr.waterType}` : '';
          const locationWithType = (ocr.location || '') + waterTypeSuffix;

          // ── Write to SharePoint Review Queue list ─────────────────────────────
          // Field names match the SharePoint list columns from setup-lists.js
          await writeToReviewQueue({
            Title:           ocrStatus,                                      // OCR Status
            ReviewStatus:    reviewStatus,
            LabID:           '',                                             // assigned at approval
            ClientName:      client ? client.clientName : (ocr.customer || ''),
            Address:         locationWithType,
            City:            ocr.city        || '',
            State:           ocr.state       || 'ME',
            Zip:             ocr.zip ? String(ocr.zip).padStart(5, '0') : '',
            Email:           client ? client.email : (ocr.email || ''),
            SampleDate:      ocr.dateDrawn   || '',
            SampleTime:      ocr.timeDrawn   || '',
            ReceivedDate:    ocr.receivedDate || '',
            ReceivedTime:    ocr.receivedTime || '',
            TestSelections:  tests.join(', '),
            ValidationErrors: valError       || '',
            OCRConfidence:   ocr.confidence  || 0,
            FileID:          file.id,
            ProcessedDate:   stamp,
            BarcodeID:       ocr.barcodeId   || '',
            ScannedBy:       scannedByName,
            ApprovedBy:      '',
            WaterType:       ocr.waterType   || '',
          }, token);

          results.push({
            fileName:     file.name,
            barcodeId:    ocr.barcodeId || '',
            barcodeMatch: !!barcodeMatch,
            client:       client?.clientName || ocr.customer,
            tests,
            confidence:   ocr.confidence,
          });

          context.log(`[scan] ✓ ${file.name} | ${client?.clientName || ocr.customer} | ${tests.join(',')} | ${ocr.confidence}%`);

        } catch (err) {
          context.log(`[scan] ✗ ${file.name}: ${err.message}`);
          // Move back to INCOMING on failure so it can be retried
          await moveSpFile(file.id, SCAN_INCOMING, token).catch(() => {});
        }
      }

      return {
        status:   200,
        jsonBody: { checked: files.length, processed: results.length, results },
      };

    } catch (err) {
      context.log(`[scan-folder] fatal: ${err.message}`);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
