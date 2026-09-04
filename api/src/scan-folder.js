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
  let lastError = null;

  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      // Resolve destination folder ID on each attempt in case Graph returned stale data.
      const folderRes = await fetch(
        `${GRAPH}/sites/${siteId}/drive/root:/${drivePath}?$select=id,name,parentReference`,
        { headers: { Authorization: `Bearer ${token}` }, cache: 'no-store' }
      );
      if (!folderRes.ok) {
        throw new Error(`cannot resolve destination ${destFolderPath}: ${folderRes.status} ${await folderRes.text()}`);
      }
      const folderData = await folderRes.json();
      const destId = folderData.id;
      if (!destId) throw new Error(`destination ${destFolderPath} returned no folder id`);

      const patchRes = await fetch(
        `${GRAPH}/sites/${siteId}/drive/items/${itemId}`,
        {
          method:  'PATCH',
          headers: { Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
          body:    JSON.stringify({ parentReference: { id: destId } }),
        }
      );
      if (!patchRes.ok) {
        throw new Error(`move failed: ${patchRes.status} ${await patchRes.text()}`);
      }

      // Verify the item really landed in Review before OCR starts. Previously move
      // failures were only logged and processing continued while the file stayed in Incoming.
      for (let verify = 0; verify < 5; verify++) {
        const metaRes = await fetch(
          `${GRAPH}/sites/${siteId}/drive/items/${itemId}?$select=id,name,parentReference`,
          { headers: { Authorization: `Bearer ${token}` }, cache: 'no-store' }
        );
        if (metaRes.ok) {
          const meta = await metaRes.json();
          if (meta.parentReference?.id === destId) return meta;
        }
        await new Promise(r => setTimeout(r, 400));
      }
      throw new Error('move verification failed; item is not in destination folder yet');
    } catch (e) {
      lastError = e;
      console.warn(`[moveSpFile] Attempt ${attempt}/3 failed for ${itemId}: ${e.message}`);
      if (attempt < 3) await new Promise(r => setTimeout(r, 750 * attempt));
    }
  }

  throw new Error(`Could not move scan into Review after 3 attempts: ${lastError?.message || 'unknown error'}`);
}

// ── SharePoint List helpers ───────────────────────────────────────────────────

// Load clients from SharePoint "Clients" list
async function loadClients(token) {
  try {
    const siteId = process.env.SP_SITE_ID;
    const res    = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields($select=Title,ClientName,ClientCode,Abbrev,Email,Aliases,Active,BillingAddress,Notes)&$top=500`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return [];
    const data = await res.json();
    return (data.value || []).map(item => ({
      clientName:     item.fields?.ClientName || item.fields?.Title || '',
      clientCode:     item.fields?.ClientCode     || '',
      abbrev:         item.fields?.Abbrev         || '',
      email:          item.fields?.Aliases        || '', // Aliases = Report Email Address
      reportEmail:    item.fields?.Aliases        || '',
      aliases:        item.fields?.Aliases        || '',
      phone:          item.fields?.Active         || '', // Active = Phone #
      billingAddress: item.fields?.BillingAddress || '',
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
    // Only check files still pending review — not already approved/archived ones
    const res    = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Review Queue/items?$expand=fields($select=FileID,ReviewStatus)&$top=2000`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return new Set();
    const data = await res.json();
    const ids  = new Set();
    (data.value || []).forEach(item => {
      const fid    = item.fields?.FileID;
      const status = (item.fields?.ReviewStatus || '').toLowerCase();
      // Only block if still pending — not if approved/archived/discarded
      const isPending = !status || status.includes('ready') || status.includes('pending') || status.includes('review');
      if (fid && isPending) ids.add(String(fid).trim());
    });
    return ids;
  } catch { return new Set(); }
}

// Write one scan result row to the "Review Queue" SP list
async function writeToReviewQueue(fields, token) {
  const siteId = process.env.SP_SITE_ID;

  const tryWrite = async (f) => {
    const res = await fetch(
      `${GRAPH}/sites/${siteId}/lists/Review Queue/items`,
      {
        method:  'POST',
        headers: { Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
        body:    JSON.stringify({ fields: f }),
      }
    );
    if (!res.ok) {
      const err = await res.text();
      throw new Error(`writeToReviewQueue: ${res.status} ${err}`);
    }
    return res.json();
  };

  try {
    return await tryWrite(fields);
  } catch(e) {
    // Retry with core fields only if a column isn't recognised yet
    if (e.message.includes('400') && e.message.includes('not recognized')) {
      const core = {};
      const coreKeys = ['Title','FileID','BarcodeID','ClientName','Email','Phone',
        'SampleDate','SampleTime','ReceivedDate','ReceivedTime','Address','City',
        'State','Zip','TestSelections','OCRConfidence','ProcessedDate','ReviewStatus',
        'ValidationErrors','WaterType','ScannedBy','IsNewClient','BillingAddress'];
      coreKeys.forEach(k => { if (fields[k] !== undefined) core[k] = fields[k]; });
      return await tryWrite(core);
    }
    throw e;
  }
}

// ── Hard-coded alias fallbacks (v352) ────────────────────────────────────────
const SCAN_HARD_ALIASES = {
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

// ── Client matching ───────────────────────────────────────────────────────────
function matchClient(name, clients) {
  if (!name || !clients.length) return null;
  const normalize = str =>
    String(str || '').toLowerCase().replace(/[.,'"()\-]/g,'').replace(/\s+/g,' ').trim();
  const s = normalize(name);
  if (s.length < 2) return null;

  // Hard-coded alias fallback
  const hardMatch = SCAN_HARD_ALIASES[s];
  if (hardMatch) {
    const found = clients.find(c => normalize(c.clientName) === normalize(hardMatch));
    if (found) return found;
  }

  for (const c of clients) {
    // Build all identifiers for this client: name, abbrev, clientCode, all aliases
    const identifiers = [
      c.clientName,
      c.abbrev,
      c.clientCode,
      ...String(c.aliases || '').split(/[,;|\n]/).map(a => a.trim()).filter(Boolean),
    ].map(normalize).filter(Boolean);

    // 1. Exact match against any identifier
    if (identifiers.some(n => n === s)) return c;

    // 2. Containment match for useful aliases (min 4 chars)
    if (identifiers.some(n => n.length >= 4 && (s.includes(n) || n.includes(s)))) return c;

    // 3. Word-based match against client name
    const cNorm  = normalize(c.clientName);
    const words  = s.split(' ').filter(w => w.length >= 4);
    if (words.length >= 2 && words.every(w => cNorm.includes(w))) return c;
  }
  return null;
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

      // Parallelize independent startup reads
      const [clients, queuedIds] = await Promise.all([
        loadClients(token),
        getQueuedFileIds(token),
      ]);
      const aliasCtx = clients.map(c =>
        `- "${c.clientName}"${c.aliases ? ` (aliases: ${c.aliases})` : ''}${c.billingAddress ? ` [billing: ${c.billingAddress}]` : ''}`
      ).join('\n') +
      '\n⚠️ Use ALL available evidence to identify the customer: the checked/marked name in the Report To section, OR match the billing address on the form against the [billing:] addresses above, OR the phone number. Return "" only if there is truly NO identifying information anywhere on the form.';

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

      // Skip files already in Review Queue (loaded in parallel above)
      const toProcess = files.filter(f => !queuedIds.has(f.id));

      if (!toProcess.length) {
        return { status: 200, jsonBody: { checked: files.length, processed: 0, message: `${files.length} file(s) already in Review Queue` } };
      }

      const results = [];

      for (let _fi = 0; _fi < toProcess.length; _fi++) {
        const file = toProcess[_fi];
        // Add delay between files to avoid Azure DI rate limiting
        if (_fi > 0) await new Promise(r => setTimeout(r, 2000));
        const fileStartedAt = Date.now();
        const timing = {};
        try {
          // Move to REVIEW immediately to prevent duplicate processing
          const moveStartedAt = Date.now();
          await moveSpFile(file.id, SCAN_REVIEW, token);
          timing.moveMs = Date.now() - moveStartedAt;

          // Download file as Buffer → base64 for Azure Doc Intel
          const downloadStartedAt = Date.now();
          const buf  = await downloadSpFile(file.id, token);
          timing.downloadMs = Date.now() - downloadStartedAt;
          const b64  = buf.toString('base64');
          const isPdf = /\.pdf$/i.test(file.name) || file.file?.mimeType === 'application/pdf';

          // ── PRIMARY: Azure Document Intelligence + Claude Haiku ─────────────
          let raw       = '';
          let azureText = '';

          const azureEndpoint = process.env.AZURE_DOC_INTEL_ENDPOINT;
          const azureKey      = process.env.AZURE_DOC_INTEL_KEY;

          // Diagnostic log — written to OCRDebug so failures are visible in the dashboard
          const scanLog = [`File: ${file.name}, size: ${buf.length}b, isPdf: ${isPdf}`];
          context.log(`[scan] STEP 1 — File: ${file.name} size: ${buf.length} bytes isPdf: ${isPdf}`);

          if (azureEndpoint && azureKey) {
            const endpoint = azureEndpoint.replace(/\/+$/, '');
            context.log(`[scan] STEP 2 — Sending to Azure DI`);

            try {
              const azureStartedAt = Date.now();
              // Start Azure analysis
              const analyzeUrl = `${endpoint}/documentintelligence/documentModels/prebuilt-layout:analyze?api-version=2024-11-30`;
              const startRes   = await fetch(analyzeUrl, {
                method:  'POST',
                headers: { 'Ocp-Apim-Subscription-Key': azureKey, 'Content-Type': 'application/json' },
                body:    JSON.stringify({ base64Source: b64 }),
              });
              if (startRes.status === 429) {
                const retryAfter = parseInt(startRes.headers.get('Retry-After') || '30');
                context.log(`[scan] Azure DI rate limited — waiting ${retryAfter}s before retry`);
                await new Promise(r => setTimeout(r, retryAfter * 1000));
                // Retry the analyze call once
                const retryRes = await fetch(analyzeUrl, {
                  method: 'POST',
                  headers: { 'Ocp-Apim-Subscription-Key': azureKey, 'Content-Type': 'application/json' },
                  body: JSON.stringify({ base64Source: b64 }),
                });
                if (!retryRes.ok) throw new Error(`Azure start retry: ${retryRes.status}`);
                Object.defineProperty(startRes, 'ok', { value: true });
                // Replace operationUrl from retry response
                const retryOpUrl = retryRes.headers.get('Operation-Location');
                if (retryOpUrl) Object.defineProperty(startRes.headers, 'get', { value: (k) => k === 'Operation-Location' ? retryOpUrl : null });
              }
              if (!startRes.ok) throw new Error(`Azure start: ${startRes.status} ${await startRes.text()}`);

              const operationUrl = startRes.headers.get('Operation-Location');
              if (!operationUrl) throw new Error('Azure returned no Operation-Location');

              // Poll until complete — scanned image PDFs can take 20-30 seconds
              let azureResult;
              await new Promise(r => setTimeout(r, 2000)); // initial wait
              for (let i = 0; i < 20; i++) {
                const pollRes = await fetch(operationUrl, { headers: { 'Ocp-Apim-Subscription-Key': azureKey } });
                azureResult   = await pollRes.json();
                if (azureResult.status === 'succeeded' || azureResult.status === 'failed') break;
                await new Promise(r => setTimeout(r, i < 5 ? 1000 : 1500)); // ramp up wait time
              }
              if (!azureResult || azureResult.status !== 'succeeded') {
                scanLog.push(`FAIL Azure: ${azureResult?.status || 'timeout'}`);
              context.log(`[scan] STEP 3 FAIL — Azure status: ${azureResult?.status || 'timeout'}`);
                throw new Error(`Azure: ${azureResult?.status || 'timeout'}`);
              }
              timing.azureMs = Date.now() - azureStartedAt;
              scanLog.push(`OK Azure succeeded in ${timing.azureMs}ms`);
              context.log(`[scan] STEP 3 OK — Azure succeeded in ${(timing.azureMs/1000).toFixed(1)}s`);

              // Build structured plain text from Azure output (page 1 only)
              const page1      = azureResult.analyzeResult?.pages?.[0];
              const pageHeight = page1?.height || 792;
              const pageWidth  = page1?.width  || 612;
              const isLandscape = pageWidth > pageHeight;
              context.log(`[scan] Page dimensions: ${pageWidth}x${pageHeight} — ${isLandscape ? 'LANDSCAPE' : 'portrait'}`);

              // For landscape forms, section by X coordinate (left=top, right=bottom of portrait form)
              // For portrait, section by Y coordinate as normal
              const getSectionPos = para => {
                if (isLandscape) {
                  // X position normalized — left side of landscape = top of portrait form
                  return (para.boundingRegions?.[0]?.polygon?.[0] || 0) / pageWidth;
                }
                return (para.boundingRegions?.[0]?.polygon?.[1] || 0) / pageHeight;
              };

              const BACK_PAGE_KEYWORDS = [
                'sample collection instructions','dropbox locations','payment information',
                'cardholder','whenever possible, collect sample from a faucet',
                'mastercard','cvv code','there is a 4% tech fee',
              ];
              const isBackPage = text => {
                const t = text.toLowerCase().trim();
                return BACK_PAGE_KEYWORDS.some(k => t.includes(k));
              };

              const paragraphs = (azureResult.analyzeResult?.paragraphs || [])
                .filter(p => p.content && !isBackPage(p.content));

              const topSection = [], middleSection = [], bottomSection = [];
              for (const para of paragraphs) {
                if (!para.content) continue;
                const normalY = getSectionPos(para);
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

              // Detect upside-down forms — if CUSTOMER section is in BOTTOM, form is reversed
              const customerInBottom = bottomSection.some(l => l.includes('CUSTOMER') || l.includes('REPORT TO BE SENT TO'));
              const customerInTop    = topSection.some(l => l.includes('CUSTOMER') || l.includes('REPORT TO BE SENT TO'));
              if (customerInBottom && !customerInTop) {
                const tmp = [...topSection];
                topSection.length = 0; bottomSection.forEach(l => topSection.push(l));
                bottomSection.length = 0; tmp.forEach(l => bottomSection.push(l));
                azureText = '';
                if (topSection.length)    azureText += '=== TOP OF FORM (Lab Use Only, Report To, Header) ===\n'              + topSection.join('\n')    + '\n\n';
                if (middleSection.length) azureText += '=== MIDDLE OF FORM (Well Owner Address, Date/Time Sampled) ===\n'     + middleSection.join('\n') + '\n\n';
                if (bottomSection.length) azureText += '=== BOTTOM OF FORM (Test Type Checkboxes, Individual Elements) ===\n' + bottomSection.join('\n') + '\n\n';
                context.log('[scan] Upside-down form detected — sections reordered');
              }

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

              scanLog.push(`azureText: ${azureText.length}chars, ${paragraphs.length}paras, ${kvPairs.length}kv`);
              context.log(`[scan] STEP 4 — paragraphs: ${paragraphs.length}, kvPairs: ${kvPairs.length}, azureText length: ${azureText.length}`);
              context.log(`[scan] STEP 4 — azureText preview: ${azureText.slice(0, 200)}`);

              // If paragraphs are thin, fall back to raw analyzeResult.content (always populated)
              // Use a high threshold (500) to catch cases where only partial content was extracted
              if (azureText.length < 500) {
                const rawContent = (azureResult.analyzeResult?.content || '')
                  .replace(/:selected:/g, '[CHECKED]')
                  .replace(/:unselected:/g, '[unchecked]');
                if (rawContent.length > azureText.length) {
                  context.log(`[scan] Paragraphs thin (${azureText.length} chars) — using raw content (${rawContent.length} chars)`);
                  // Re-section the raw content using page-level word positions so Claude
                  // still knows which part is Report To vs Well Owner vs Tests
                  const words      = azureResult.analyzeResult?.pages?.[0]?.words || [];
                  const rawTop = [], rawMid = [], rawBot = [];
                  words.forEach(w => {
                    const pos = isLandscape
                      ? (w.polygon?.[0] || 0) / pageWidth   // X position for landscape
                      : (w.polygon?.[1] || 0) / pageHeight; // Y position for portrait
                    const t = (w.content || '').replace(/:selected:/g,'[CHECKED]').replace(/:unselected:/g,'[unchecked]');
                    if (pos < 0.35)      rawTop.push(t);
                    else if (pos < 0.75) rawMid.push(t);
                    else                 rawBot.push(t);
                  });
                  if (rawTop.length || rawMid.length || rawBot.length) {
                    azureText = '';
                    if (rawTop.length) azureText += '=== TOP OF FORM (Lab Use Only, Report To, Header) ===\n' + rawTop.join(' ') + '\n\n';
                    if (rawMid.length) azureText += '=== MIDDLE OF FORM (Well Owner Address, Date/Time Sampled) ===\n' + rawMid.join(' ') + '\n\n';
                    if (rawBot.length) azureText += '=== BOTTOM OF FORM (Test Type Checkboxes, Individual Elements) ===\n' + rawBot.join(' ') + '\n\n';
                    context.log(`[scan] Rebuilt sections from word positions: top=${rawTop.length} mid=${rawMid.length} bot=${rawBot.length} words`);
                  } else {
                    azureText = rawContent; // no word positions — use flat content as last resort
                  }
                }
              }

              // Supplement with explicit selection marks from page (most reliable checkbox detection)
              const selMarks = azureResult.analyzeResult?.pages?.[0]?.selectionMarks || [];
              const pageWords = azureResult.analyzeResult?.pages?.[0]?.words || [];
              if (selMarks.length > 0) {
                const selText = selMarks.map(mark => {
                  const state  = mark.state === 'selected' ? '[CHECKED]' : '[unchecked]';
                  const markX  = mark.polygon?.[0] || 0;
                  const markY  = mark.polygon?.[1] || 0;
                  const nearby = pageWords
                    .filter(w => Math.abs((w.polygon?.[1]||0) - markY) < 12 && (w.polygon?.[0]||0) > markX - 5)
                    .sort((a,b) => (a.polygon?.[0]||0) - (b.polygon?.[0]||0))
                    .slice(0, 6).map(w => w.content).join(' ');
                  return `${state} ${nearby}`;
                }).join('\n');
                azureText += '\n\n=== SELECTION MARKS (checkboxes) ===\n' + selText;
                context.log(`[scan] Added ${selMarks.length} selection marks to azureText`);
              }

              // Claude Haiku structures Azure's text into JSON; Sonnet is reserved for recovery
              const claudeStartedAt = Date.now();
              const extractRes = await fetch('https://api.anthropic.com/v1/messages', {
                method:  'POST',
                headers: {
                  'Content-Type':      'application/json',
                  'x-api-key':         process.env.ANTHROPIC_API_KEY,
                  'anthropic-version': '2023-06-01',
                },
                body: JSON.stringify({
                  model:      'claude-haiku-4-5',
                  max_tokens: 800,
                  system:     'You are a JSON extraction API. Output ONLY a valid JSON object. No markdown, no explanation.',
                  messages: [{ role: 'user', content:
`Extract from this Chanalytical Laboratories Chain of Custody form.
[CHECKED] = checked checkbox. [unchecked] = unchecked. Azure Document Intelligence already read the text.

FORM TEXT:
${azureText}

RULES:
- Extract only what is present in the OCR text. Do not infer or guess a known client identity; local matching happens after extraction.
- formType: "business" if Report To section has a company/person name other than Chanalytical. "public" if Report To is blank or shows Chanalytical.
- customer: BUSINESS — find the company name next to [CHECKED] in Report To section. If fill-in line, copy what's written. "" if blank/nothing marked. PUBLIC — person's name from Customer & Property Information "Name:" field. "" if blank.
- location: BUSINESS=well owner street address (MIDDLE section). PUBLIC=customer street address (TOP section). Never use Report To address.
- city/state/zip: from Well Owner or Customer section only.
- dateDrawn: "Date Sampled" field → YYYY-MM-DD. "" if blank or crossed out.
- timeDrawn: "Time Sampled" → HH:MM 24hr (convert AM/PM). "" if blank.
- receivedDate/receivedTime: from "Lab Use Only" box ONLY → YYYY-MM-DD / HH:MM.
- barcodeId: alphanumeric code in Lab Use Only box. "" if absent.
- tests: package names with [CHECKED]. Valid: "Basic Safety (FHA)","Standard Safety","Expanded Safety (Mortgage Test)","WW - Expanded Safety","Comprehensive","Pro Plus"
- hasRadon: true ONLY if [CHECKED] next to "Radon Water" test type.
- individualElements: individual element rows with [CHECKED]. "TDS"="Total Dissolved Solids (TDS)".
- email: only if contains @. PUBLIC forms only. "" otherwise.
- phone: from Daytime Phone/Phone/Cell field. "" if blank.
- billingAddress: BUSINESS only — street+city+state+zip from Report To section as one line. "" for public.
- waterType: "Raw" or "Treated" if stated. "" otherwise.
- confidence: 0-100

Return ONLY: {"barcodeId":"","formType":"public","customer":"","email":"","phone":"","billingAddress":"","dateDrawn":"","timeDrawn":"","receivedDate":"","receivedTime":"","location":"","city":"","state":"ME","zip":"","tests":[],"individualElements":[],"hasRadon":false,"notes":"","waterType":"","confidence":0}`
                  }],
                }),
              });
              if (!extractRes.ok) throw new Error(`Claude extract: ${extractRes.status}`);
              const extractData = await extractRes.json();
              raw = extractData.content?.find(c => c.type === 'text')?.text || '';
              timing.haikuMs = Date.now() - claudeStartedAt;
              scanLog.push(`Haiku raw: ${raw.length}chars in ${timing.haikuMs}ms`);
              context.log(`[scan] STEP 5 — Haiku response length: ${raw.length} in ${(timing.haikuMs/1000).toFixed(1)}s`);
              context.log(`[scan] STEP 5 — Claude preview: ${raw.slice(0, 300)}`);

              // Retry if primary extraction returned near-empty result
              if (raw) {
                try {
                  const s = raw.indexOf('{'), e = raw.lastIndexOf('}');
                  if (s >= 0 && e > s) {
                    const testParse = JSON.parse(raw.slice(s, e + 1));
                    const weakExtraction = (Number(testParse.confidence || 0) < 50) &&
                      !testParse.customer && !testParse.location &&
                      !testParse.tests?.length && !testParse.individualElements?.length && !testParse.hasRadon;
                    if (weakExtraction) {
                      context.log('[scan] STEP 5 — Haiku extraction weak, retrying once with Sonnet');
                      const sonnetStartedAt = Date.now();
                      const retryRes = await fetch('https://api.anthropic.com/v1/messages', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json', 'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
                        body: JSON.stringify({
                          model: 'claude-sonnet-4-6',
                          max_tokens: 600,
                          system: 'You are a JSON extraction API. Output ONLY a valid JSON object.',
                          messages: [{ role: 'user', content:
`This is a water testing Chain of Custody form. Extract ONLY these fields:
- customer: the company or person name that is CHECKED or written in the "Report To" section. If individual, their name from the Customer section.
- tests: list of test packages with [CHECKED] next to them.
- hasRadon: true if Radon Water is [CHECKED].
- dateDrawn: date sampled as YYYY-MM-DD.
- location: the well owner or property street address.
- confidence: 0-100

FORM TEXT:
${azureText.slice(0, 3000)}

Return ONLY: {"customer":"","tests":[],"hasRadon":false,"dateDrawn":"","location":"","city":"","state":"ME","zip":"","receivedDate":"","receivedTime":"","confidence":0}`
                          }],
                        }),
                      });
                      if (retryRes.ok) {
                        const retryData = await retryRes.json();
                        const retryRaw = retryData.content?.find(c => c.type === 'text')?.text || '';
                        timing.sonnetRetryMs = Date.now() - sonnetStartedAt;
                        if (retryRaw && retryRaw.includes('{')) raw = retryRaw;
                        context.log(`[scan] STEP 5 Sonnet retry — ${(timing.sonnetRetryMs/1000).toFixed(1)}s — result:`, retryRaw.slice(0, 200));
                        scanLog.push(`retry raw: ${retryRaw.length}chars`);
                      }
                    }
                  }
                } catch(retryErr) { context.log('[scan] Retry parse/check failed:', retryErr.message); }
              }

            } catch (azureErr) {
              context.log(`[scan] Azure hybrid failed: ${azureErr.message}`);
              scanLog.push(`FAIL Azure exception: ${azureErr.message.slice(0, 100)}`);
              results.push({ azureError: azureErr.message });
              azureText = `[Azure DI failed: ${azureErr.message}]`;
            }
          }

          // ── FALLBACK: Claude Sonnet text-only (if Azure not configured or failed) ──
          if (!raw) {
            scanLog.push(`fallback — azureText: ${azureText.length}chars`);
            const step2Res = await fetch('https://api.anthropic.com/v1/messages', {
              method:  'POST',
              headers: {
                'Content-Type':      'application/json',
                'x-api-key':         process.env.ANTHROPIC_API_KEY,
                'anthropic-version': '2023-06-01',
              },
              body: JSON.stringify({
                model:      'claude-sonnet-4-6',
                max_tokens: 800,
                system:     'You are a JSON extraction API. Output ONLY a valid JSON object. No markdown, no explanation.',
                messages: [{ role:'user', content:
`Extract from this Chanalytical Laboratories Chain of Custody form. Known clients: ${aliasCtx}

RULES:
- formType: "business" if Report To has a company/person name other than Chanalytical. "public" if Report To blank or shows Chanalytical.
- customer: BUSINESS=company name next to [CHECKED] in Report To, or what's written on fill-in line. "". PUBLIC=person's name from Customer & Property Information "Name:" field. "" if blank.
- location: BUSINESS=well owner street address. PUBLIC=customer street address. Never Report To address.
- city/state/zip: from Well Owner or Customer section.
- dateDrawn: Date Sampled → YYYY-MM-DD. "" if blank.
- timeDrawn: Time Sampled → HH:MM 24hr. "" if blank.
- receivedDate/receivedTime: Lab Use Only box only.
- tests: packages with [CHECKED]. Valid: "Basic Safety (FHA)","Standard Safety","Expanded Safety (Mortgage Test)","WW - Expanded Safety","Comprehensive","Pro Plus"
- hasRadon: true ONLY if [CHECKED] next to "Radon Water".
- individualElements: individual rows with [CHECKED]. "TDS"="Total Dissolved Solids (TDS)".
- confidence: 0-100

COC TEXT:
${azureText}

Return ONLY: {"barcodeId":"","formType":"public","customer":"","email":"","phone":"","billingAddress":"","dateDrawn":"","timeDrawn":"","receivedDate":"","receivedTime":"","location":"","city":"","state":"ME","zip":"","tests":[],"individualElements":[],"hasRadon":false,"notes":"","waterType":"","confidence":0}`
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
            scanLog.push(`JSON OK: customer="${ocr.customer}", tests=${ocr.tests?.length||0}, conf=${ocr.confidence}`);
            context.log(`[scan] STEP 6 OK — JSON parsed, customer: "${ocr.customer}", tests: ${ocr.tests?.length || 0}, confidence: ${ocr.confidence}`);
          } catch {
            scanLog.push(`JSON FAIL: ${raw.slice(0,100)}`);
            context.log(`[scan] STEP 6 FAIL — JSON parse error, raw: ${raw.slice(0, 200)}`);
            throw new Error(`OCR JSON parse failed: ${raw.slice(0, 200)}`);
          }

          if (!ocr.customer && !ocr.location && !ocr.dateDrawn && !ocr.tests?.length) {
            context.log(`[scan] OCR returned empty for ${file.name} — writing to Review Queue for manual entry`);
            ocr.confidence = 0;
          }

          // ── Normalize and clean ───────────────────────────────────────────────
          ocr.customer = normalizeCase(ocr.customer);
          ocr.location = normalizeCase(ocr.location);
          ocr.city     = normalizeCase(ocr.city);
          ocr.location = cleanAddress(ocr.location);
          ocr.zip      = fixZip(ocr.zip, ocr.state);

          // Validate email - must contain @ otherwise clear it
          // If customer is also blank, rescue the name and use it as customer
          if (ocr.email && !ocr.email.includes('@')) {
            context.log(`[scan] Invalid email cleared: "${ocr.email}"`);
            if (!ocr.customer) {
              ocr.customer = ocr.email; // rescue misplaced name
              context.log(`[scan] Rescued name "${ocr.customer}" from email field`);
            }
            ocr.email = '';
          }

          // Use reportToName as fallback for customer if customer is still empty
          if (!ocr.customer && ocr.reportToName) {
            ocr.customer = ocr.reportToName;
            context.log(`[scan] Used reportToName "${ocr.customer}" as customer`);
          }

          // If billingAddress looks like a company name (no digits = no street number),
          // rescue it as the customer name
          if (!ocr.customer && ocr.billingAddress && !/\d/.test(ocr.billingAddress)) {
            ocr.customer = ocr.billingAddress;
            ocr.billingAddress = '';
            context.log(`[scan] Rescued company name "${ocr.customer}" from billingAddress`);
          }

          // Clear email if it looks like a form label (no @ and contains spaces suggesting a phrase)
          if (ocr.email && !ocr.email.includes('@') && ocr.email.includes(' ')) {
            context.log(`[scan] Cleared label text from email: "${ocr.email}"`);
            ocr.email = '';
          }

          // For public customers with no billing address, build full address
          if (!ocr.billingAddress) {
            const parts = [ocr.location, ocr.city, ocr.state, ocr.zip].filter(Boolean);
            if (parts.length) ocr.billingAddress = parts.join(', ');
          }

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
          // Refresh validatedCustomer AFTER all rescues AND barcode lookup
          validatedCustomer = ocr.customer || validatedCustomer;
          let client = matchClient(validatedCustomer, clients);

          // ── Fallback matching: email → phone → billing address (tried individually) ──
          const normalize = s => String(s||'').toLowerCase().replace(/[^a-z0-9]/g,' ').replace(/\s+/g,' ').trim();

          // Debug: log state going into fallback matching
          context.log(`[scan] Fallback input — email="${ocr.email}" phone="${ocr.phone}" billing="${ocr.billingAddress}" clients=${clients.length}`);
          if (ocr.email) {
            const sample = clients.slice(0,3).map(c => `${c.clientName}|${c.reportEmail||c.email}`).join('; ');
            context.log(`[scan] Sample clients: ${sample}`);
            const yankee = clients.find(c => c.clientName.toLowerCase().includes('yankee'));
            if (yankee) context.log(`[scan] Yankee entry: reportEmail="${yankee.reportEmail}" email="${yankee.email}" phone="${yankee.phone}"`);
          }

          // 1. Email: look up the exact email against Clients list report email
          if (!client && ocr.email && ocr.email.includes('@')) {
            const emailLow = ocr.email.toLowerCase();
            client = clients.find(c => (c.reportEmail||c.email||'').toLowerCase() === emailLow) || null;
            if (client) context.log(`[scan] Matched by email: ${client.clientName}`);
            else context.log(`[scan] Email fallback: no match for "${emailLow}"`);
          }

          // 3. Phone: match last 7 digits against Clients list phone
          if (!client && (ocr.phone || ocr.reportToPhone)) {
            const pd = (ocr.phone||ocr.reportToPhone||'').replace(/\D/g,'');
            if (pd.length >= 7) {
              client = clients.find(c => c.phone && c.phone.replace(/\D/g,'').endsWith(pd.slice(-7))) || null;
              if (client) context.log(`[scan] Matched by phone: ${client.clientName}`);
            }
          }

          // 4. Billing address: match first 3 significant words against Clients list billing address
          if (!client && ocr.billingAddress) {
            const baLow   = ocr.billingAddress.toLowerCase().replace(/[.,]/g,'').trim();
            const baParts = baLow.split(/\s+/).filter(w => w.length >= 3);
            if (baParts.length >= 2) {
              client = clients.find(c => {
                const ca = (c.billingAddress||'').toLowerCase().replace(/[.,]/g,'');
                if (!ca) return false;
                return baParts.slice(0,3).every(w => ca.includes(w));
              }) || null;
              if (client) context.log(`[scan] Matched by billing address: ${client.clientName}`);
            }
          }

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

          // ── Ward Water test conversion ────────────────────────────────────────
          // Ward Water uses WW - Expanded Safety instead of Expanded Safety (Mortgage Test)
          const isWardWater = clientName.toLowerCase().includes('ward water') || clientName.toLowerCase().includes('critical plumbing');
          if (isWardWater) {
            ocr.tests = (ocr.tests || []).map(t => {
              if (/expanded safety/i.test(t)) return 'WW - Expanded Safety';
              return t;
            });
            context.log(`[scan] Ward Water client — converted Expanded Safety to WW - Expanded Safety`);
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
          // Build match debug info
          const yankeeDbg = clients.find(c => c.clientName.toLowerCase().includes('yankee'));
          const matchDebug = {
            extracted:      ocr.customer,
            validated:      validatedCustomer,
            matched:        client?.clientName || 'NO MATCH',
            formType:       ocr.formType || '',
            clientsLoaded:  clients.length,
            yankeeEmail:    yankeeDbg?.reportEmail || yankeeDbg?.email || 'NOT IN LIST',
            yankeePhone:    yankeeDbg?.phone || 'NOT IN LIST',
            ocr_email:      ocr.email,
            ocr_phone:      ocr.phone,
          };
          context.log('[scan] CUSTOMER DEBUG', JSON.stringify(matchDebug));

          const queueStartedAt = Date.now();
          await writeToReviewQueue({
            Title:            reviewStatus,
            LabID:            '',
            ClientName:       client ? client.clientName : (ocr.customer || ''),
            Address:          locationWithType,
            City:             ocr.city         || '',
            State:            ocr.state        || 'ME',
            Zip:              ocr.zip ? String(ocr.zip).padStart(5, '0') : '',
            Email:            client ? (client.reportEmail || client.email) : (ocr.email || ''),
            Phone:            ocr.phone || ocr.reportToPhone || '',
            BillingAddress:   ocr.billingAddress || (ocr.formType === 'business'
              ? [ocr.reportToAddress, ocr.reportToCity, ocr.reportToState, ocr.reportToZip].filter(Boolean).join(', ')
              : ''),
            IsNewClient:      client ? 'No' : 'Yes',
            FormType:         ocr.formType || 'public',
            SampleDate:       ocr.dateDrawn    || '',
            SampleTime:       ocr.timeDrawn    || '',
            ReceivedDate:     ocr.receivedDate || '',
            ReceivedTime:     ocr.receivedTime || '',
            TestSelections:   tests.join(', '),
            ValidationErrors: valError || '',
            OCRConfidence:    ocr.confidence   || 0,
            FileID:           file.id,
            ProcessedDate:    stamp,
            BarcodeID:        ocr.barcodeId    || '',
            ScannedBy:        scannedByName,
            ApprovedBy:       '',
            WaterType:        ocr.waterType    || '',
          }, token);

          timing.reviewQueueMs = Date.now() - queueStartedAt;
          timing.totalMs = Date.now() - fileStartedAt;
          scanLog.push(`TIMING move=${timing.moveMs||0}ms download=${timing.downloadMs||0}ms azure=${timing.azureMs||0}ms haiku=${timing.haikuMs||0}ms sonnetRetry=${timing.sonnetRetryMs||0}ms reviewQueue=${timing.reviewQueueMs||0}ms total=${timing.totalMs}ms`);
          context.log(`[scan] TIMING ${file.name} — move ${(timing.moveMs||0)/1000}s | download ${(timing.downloadMs||0)/1000}s | Azure ${((timing.azureMs||0)/1000).toFixed(1)}s | Haiku ${((timing.haikuMs||0)/1000).toFixed(1)}s | Sonnet retry ${((timing.sonnetRetryMs||0)/1000).toFixed(1)}s | Queue ${((timing.reviewQueueMs||0)/1000).toFixed(1)}s | TOTAL ${(timing.totalMs/1000).toFixed(1)}s`);

          results.push({
            fileId:       file.id,
            fileName:     file.name,
            barcodeId:    ocr.barcodeId || '',
            barcodeMatch: !!barcodeMatch,
            client:       client?.clientName || ocr.customer,
            tests,
            confidence:   ocr.confidence,
            timing,
            ocrExtracted: { phone: ocr.phone, billingAddress: ocr.billingAddress, email: ocr.email, customer: ocr.customer },
            ocrTextSnippet: azureText.slice(0, 800), // first 800 chars for debugging
          });

          context.log(`[scan] ✓ ${file.name} | ${client?.clientName || ocr.customer} | ${tests.join(',')} | ${ocr.confidence}%`);

        } catch (err) {
          context.log(`[scan] ✗ ${file.name}: ${err.message}`);
          results.push({ fileName: file.name, error: err.message });
          // Try to write a minimal error card so the file shows in the Review Queue
          try {
            const errStamp = new Date().toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' })
              + ' ' + new Date().toLocaleTimeString('en-US', { timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
            await writeToReviewQueue({
              Title:            errStamp,
              FileID:           file.id,
              ClientName:       '',
              TestSelections:   '',
              OCRConfidence:    0,
              ProcessedDate:    errStamp,
              ReviewStatus:     'Ready to Review',
              ReceivedDate:     new Date().toISOString().slice(0,10),
              ValidationErrors: `Error: ${err.message}`,
              State:            'ME',
              ScannedBy:        scannedByName,
            }, token);
            context.log(`[scan] Wrote error card for ${file.name}`);
          } catch(e2) { context.log(`[scan] Could not write error card: ${e2.message}`); }
          // Leave file in REVIEW folder
        }
      }

      return {
        status:   200,
        jsonBody: { checked: files.length, processed: results.filter(r=>!r.error).length, results },
      };

    } catch (err) {
      context.log(`[scan-folder] fatal: ${err.message}`);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
