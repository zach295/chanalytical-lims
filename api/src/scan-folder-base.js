const { getDrive, getSheets, SPREADSHEET_ID, FOLDERS, SHEETS } = require('./sheets-auth');
const { todayDisplayET, nowTimeET } = require('./et-time');

// ── HELPERS ───────────────────────────────────────────────────────────────────

async function loadClients(sheets) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SHEETS.CLIENTS}'!A:H`,
    });
    return (res.data.values||[]).slice(1).filter(r=>r[0]).map(r=>({
      clientName: r[0]||'', clientCode: r[1]||'', abbrev: r[2]||'',
      email: r[3]||'', aliases: r[4]||'',
    }));
  } catch { return []; }
}

function matchClient(name, clients) {
  if (!name || !clients.length) return null;
  const s = name.toLowerCase().trim();
  if (s.length < 3) return null;
  return (
    // 1. Exact match
    clients.find(c => c.clientName.toLowerCase() === s) ||
    // 2. Alias match — alias must be meaningful (4+ chars)
    clients.find(c => c.aliases.split(',').map(a => a.trim().toLowerCase())
      .some(a => a.length >= 4 && (s === a || s.includes(a) || a.includes(s)))) ||
    null
    // Note: removed broad partial word match — it caused false positives (e.g. "water" matching A-Z Water Systems)
    // Clients should be matched via exact name or aliases set up in the Clients sheet
  );
}

function validateTests(tests) {
  const pkgs = ['Basic Safety (FHA)','Basic Safety','Standard Safety','Expanded Safety (Mortgage Test)','WW - Expanded Safety','Comprehensive','Pro Plus'];
  const addons = ['Radon Water'];
  const hasPkg = tests.some(t=>pkgs.includes(t));
  const hasInd = tests.some(t=>!pkgs.includes(t)&&!addons.includes(t));
  return hasPkg&&hasInd ? 'Package tests cannot be combined with individual elements' : '';
}

// Convert ALL-CAPS OCR text to Title Case for readability
// Only converts if the whole string is uppercase (e.g. handwriting in all caps)
function normalizeCase(s) {
  if (!s || s.length <= 2) return s;
  // Only normalize if string is predominantly uppercase
  const letters = s.replace(/[^a-zA-Z]/g, '');
  if (!letters.length) return s;
  const upperRatio = (s.match(/[A-Z]/g)||[]).length / letters.length;
  if (upperRatio > 0.8) {
    // Title case: capitalize first letter of each word
    return s.toLowerCase().replace(/\b([a-z])/g, c => c.toUpperCase());
  }
  return s;
}

function todayDisplay() {
  const n = new Date();
  return `${String(n.getMonth()+1).padStart(2,'0')}-${String(n.getDate()).padStart(2,'0')}-${String(n.getFullYear()).slice(-2)}`;
}

function nowTime() {
  return new Date().toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit',hour12:true});
}

async function moveFile(drive, fileId, toFolder) {
  try {
    const m = await drive.files.get({fileId, fields:'parents'});
    await drive.files.update({fileId, addParents:toFolder, removeParents:(m.data.parents||[]).join(','), fields:'id'});
  } catch(e) { console.warn('[move]',e.message); }
}

async function ensureQueueHeader(sheets) {
  try {
    const r = await sheets.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID, range:`'${SHEETS.REVIEW_QUEUE}'!A1`});
    if (r.data.values?.length) return;
  } catch {}
  await sheets.spreadsheets.values.update({
    spreadsheetId:SPREADSHEET_ID, range:`'${SHEETS.REVIEW_QUEUE}'!A1`,
    valueInputOption:'USER_ENTERED',
    requestBody:{values:[['OCR Status','Review Status','Lab ID','Client Name','Address','City','State','Zip','Email','Sample Date','Sample Time','Received Date','Received Time','Test Selections','Validation Errors','OCR Confidence','File ID','Processed Date','Barcode ID','Scanned By','Approved By']]},
  });
}

async function lookupBarcode(sheets, barcodeId) {
  if (!barcodeId||!barcodeId.startsWith('CHA-')) return null;
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${SHEETS.SAMPLE_INTAKE}'!A:O`,
    });
    const rows = (res.data.values||[]).slice(1);
    const row  = rows.find(r => String(r[0]||'').trim() === barcodeId.trim());
    if (!row) return null;
    return {
      barcodeId,
      customer: row[3]||'', email: row[4]||'', phone: row[5]||'',
      location: row[7]||'', city: row[8]||'', state: row[9]||'ME',
      zip: row[10]||'', services: row[11]||'',
    };
  } catch { return null; }
}

// ── HANDLER ───────────────────────────────────────────────────────────────────
exports.handler = async (event) => {
  if (event.httpMethod && !['GET','POST'].includes(event.httpMethod)) {
    return {statusCode:405, body:'Method Not Allowed'};
  }
  if (!process.env.ANTHROPIC_API_KEY) {
    return {statusCode:500, body:JSON.stringify({error:'ANTHROPIC_API_KEY not set'})};
  }

  try {
    const drive  = getDrive();
    const sheets = getSheets();
    const { scannedBy } = (event.body ? JSON.parse(event.body) : {});
    const scannedByName = scannedBy || 'Lab Staff';

    // Load clients for alias matching and OCR prompt context
    const clients  = await loadClients(sheets);
    const aliasCtx = clients.map(c=>`- "${c.clientName}"${c.aliases?` (aliases: ${c.aliases})`:''}`).join('\n') +
      '\n⚠️ IMPORTANT: Only match to a client above if their name or alias is EXPLICITLY WRITTEN on this form. Do NOT use your training knowledge to guess which client submitted it. If nothing is written, return "".';

    // List files in INCOMING folder
    const listRes = await drive.files.list({
      q: `'${FOLDERS.INCOMING}' in parents and (mimeType contains 'image/' or mimeType='application/pdf') and trashed=false`,
      fields: 'files(id,name,mimeType)',
      orderBy: 'createdTime asc',
      pageSize: 10,
    });
    const files = listRes.data.files || [];

    if (!files.length) {
      return {statusCode:200, body:JSON.stringify({checked:0, processed:0, message:'No files in the INCOMING folder'})};
    }

    // Skip files already in the Review Queue
    const queuedIds = new Set();
    try {
      const qr = await sheets.spreadsheets.values.get({spreadsheetId:SPREADSHEET_ID, range:`'${SHEETS.REVIEW_QUEUE}'!Q:Q`});
      (qr.data.values||[]).slice(1).forEach(r=>{if(r[0]) queuedIds.add(String(r[0]).trim());});
    } catch {}

    const toProcess = files.filter(f=>!queuedIds.has(f.id));
    if (!toProcess.length) {
      return {statusCode:200, body:JSON.stringify({checked:files.length, processed:0, message:`${files.length} file(s) already in Review Queue`})};
    }

    await ensureQueueHeader(sheets);
    const results = [];

    for (const file of toProcess) {
      try {
        // Move to REVIEW immediately to prevent duplicate processing
        await moveFile(drive, file.id, FOLDERS.REVIEW);

        // Download file as base64
        const dlRes = await drive.files.get({fileId:file.id, alt:'media'},{responseType:'arraybuffer'});
        const b64   = Buffer.from(dlRes.data).toString('base64');
        const isPdf = file.mimeType === 'application/pdf';

        // ── PRIMARY: Azure Document Intelligence + Claude Sonnet ──────────────
        // Azure reads handwriting and detects checkboxes reliably.
        // Claude Sonnet structures the plain text output into JSON.
        let raw = '';
        const azureEndpoint = process.env.AZURE_DOC_INTEL_ENDPOINT;
        const azureKey      = process.env.AZURE_DOC_INTEL_KEY;

        if (azureEndpoint && azureKey) {
          const endpoint = azureEndpoint.replace(/\/+$/, ''); // strip trailing slash
          console.log(`[scan] Azure endpoint: ${endpoint.slice(0,50)}, key: ${azureKey.slice(0,8)}...`);
          try {
            // Start Azure analysis
            const analyzeUrl = `${endpoint}/documentintelligence/documentModels/prebuilt-layout:analyze?api-version=2024-11-30`;
            console.log(`[scan] Azure URL: ${analyzeUrl}`);
            const startRes = await fetch(analyzeUrl, {
              method: 'POST',
              headers: {
                'Ocp-Apim-Subscription-Key': azureKey,
                'Content-Type': 'application/json',
              },
              body: JSON.stringify({ base64Source: b64 }),
            });
            if (!startRes.ok) throw new Error(`Azure start: ${startRes.status} ${await startRes.text()}`);

            const operationUrl = startRes.headers.get('Operation-Location');
            if (!operationUrl) throw new Error('Azure returned no Operation-Location');

            // Poll until complete — Azure typically finishes in 3-8 seconds
            let azureResult;
            await new Promise(r => setTimeout(r, 2000)); // initial 2s wait
            for (let i = 0; i < 12; i++) { // max ~20s polling (12 × 1.5s + 2s initial)
              const pollRes = await fetch(operationUrl, { headers: {'Ocp-Apim-Subscription-Key': azureKey} });
              azureResult = await pollRes.json();
              if (azureResult.status === 'succeeded' || azureResult.status === 'failed') break;
              await new Promise(r => setTimeout(r, 1500));
            }
            if (!azureResult || azureResult.status !== 'succeeded') {
              throw new Error(`Azure: ${azureResult?.status || 'timeout'}`);
            }

            // Build structured plain text from Azure output
            // Only use page 1 — page 2 has Dropbox Locations, payment info etc. which confuse extraction
            const page1 = azureResult.analyzeResult?.pages?.[0];
            const pageHeight = page1?.height || 792;
            const pageWidth  = page1?.width  || 612;

            // Skip paragraphs that are clearly from the back page
            // (collection instructions, dropbox table, payment info)
            const BACK_PAGE_KEYWORDS = [
              'sample collection instructions',
              'dropbox locations',
              'payment information',
              'cardholder',
              'whenever possible, collect sample from a faucet',
              'mastercard',
              'cvv code',
              'there is a 4% tech fee',
            ];
            const isBackPage = (text) => {
              const t = text.toLowerCase();
              return BACK_PAGE_KEYWORDS.some(k => t.includes(k));
            };

            // Use paragraphs from ALL pages, but skip back-page content
            const paragraphs = (azureResult.analyzeResult?.paragraphs || [])
              .filter(p => p.content && !isBackPage(p.content));

            // Organize paragraphs into spatial sections based on Y position
            const topSection    = []; // top 25% of page — header, Lab Use Only, Report To
            const middleSection = []; // 25-70% — Well Owner, Date Sampled
            const bottomSection = []; // bottom 30% — checkboxes, test types

            for (const para of paragraphs) {
              if (!para.content) continue;
              const y = para.boundingRegions?.[0]?.polygon?.[1] ?? 0; // top-left Y
              const x = para.boundingRegions?.[0]?.polygon?.[0] ?? 0; // top-left X
              const normalY = y / pageHeight;
              const normalX = x / pageWidth;
              const line = para.content
                .replace(/:selected:/g,   '[CHECKED]')
                .replace(/:unselected:/g, '[unchecked]');
              // Lab Use Only is upper-right; test checkboxes are right column
              // Use both Y and X position to better categorize
              if (normalY < 0.35)                      topSection.push(line);
              else if (normalY < 0.75)                 middleSection.push(line);
              else                                     bottomSection.push(line);
            }

            let azureText = '';
            if (topSection.length)    azureText += '=== TOP OF FORM (Lab Use Only, Report To, Header) ===\n' + topSection.join('\n') + '\n\n';
            if (middleSection.length) azureText += '=== MIDDLE OF FORM (Well Owner Address, Date/Time Sampled) ===\n' + middleSection.join('\n') + '\n\n';
            if (bottomSection.length) azureText += '=== BOTTOM OF FORM (Test Type Checkboxes, Individual Elements) ===\n' + bottomSection.join('\n') + '\n\n';

            // Key-value pairs (page 1 only)
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

            console.log(`[scan] Azure: ${paragraphs.length} page-1 paragraphs, ${kvPairs.length} key-value pairs`);
            console.log(`[scan] Azure full text:\n${azureText}`);

            // Claude Sonnet structures Azure's text into JSON
            const extractRes = await fetch('https://api.anthropic.com/v1/messages', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
                'x-api-key': process.env.ANTHROPIC_API_KEY,
                'anthropic-version': '2023-06-01',
              },
              body: JSON.stringify({
                model: 'claude-sonnet-4-6',
                max_tokens: 600,
                system: 'You are a JSON extraction API. Output ONLY a valid JSON object. No markdown, no explanation.',
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
  • BUSINESS FORM: Look in "REPORT TO BE SENT TO:" — company name may be on that line OR in the "Name:" field below. Match to known clients if confident. If nothing written → return "". ⛔ Never use form titles or guess.
- location: Street address — for public forms use the Address field in CUSTOMER & PROPERTY INFORMATION (TOP section); for business forms use the Well Owner section (MIDDLE section). Include ALL lines under the "Address:" label until the next label (City:, State:, etc.) even if they look like partial words — e.g. "was" after a street name is likely "Way" misread by OCR, include it. No periods or commas. If a circled T or ⊕ symbol appears → set waterType to "Treated". ⛔ Never use Report To address.
- city/state/zip: from MIDDLE OF FORM (Well Owner) ONLY. Maine zip starts with 04, NH starts with 03.
- email: For PUBLIC forms, extract from "E-mail:" field in CUSTOMER & PROPERTY INFORMATION section. The email may span TWO lines (e.g. "travis.john.gould" on one line and "@gmail.com" on next) — concatenate them into one address with no spaces. If not a public form or email is blank → return "".
- dateDrawn: Date CLIENT collected sample — next to "Date Sampled:" label → YYYY-MM-DD. ⛔ If date is crossed out/scribed through, do NOT use it — return "" instead. 2-digit years: 26=2026, 25=2025.
- timeDrawn: Time CLIENT collected sample — next to "Time Sampled:" label → HH:MM 24-hour. IMPORTANT: look carefully for this — it is often written as "9:59 AM", "9:59 A", "9:59a", "3:50 p", "3:50pm". "a"=AM, "p"=PM. Convert to 24hr (3:50 PM = 15:50, 9:59 AM = 09:59). ⛔ If time is crossed out, ignore it. Digit confusion: "0" vs "1", "5" vs "6". Minutes must be 00-59.
- receivedDate: Date LAB received sample — in the small "Lab Use Only" box (upper right corner). Stamped or written by lab staff. Looks like "JUN 24" or "07/14/26". Year always 2026 → YYYY-MM-DD.
- receivedTime: Time LAB received sample — in same "Lab Use Only" box. Looks like "14:06". DIFFERENT from Time Sampled. Minutes 00-59 only.
- tests: Package names where [CHECKED] appears in the right-side TEST TYPE column. Valid: "Basic Safety (FHA)","Standard Safety","Expanded Safety (Mortgage Test)","WW - Expanded Safety","Comprehensive","Pro Plus"
- hasRadon: true ONLY if [CHECKED] next to "Radon Water" in the TEST TYPE column. ⛔ NOT the "Radon Water Mitigation" note checkbox.
- individualElements: Elements where [CHECKED] appears in the LEFT column of individual tests. Look carefully — these are separate from package tests. Valid names include: Alkalinity, Arsenic Total, Bacteria, Cadmium Total, Calcium Total, Chloride Total, Copper Total, Fluoride, Hardness Total, Iron Total, Lead Total, Magnesium Total, Manganese Total, Nitrate, Nitrite, pH, Sodium Total, Sulfate, Sulfur, Tannins, TDS, Uranium Total. Include ALL that have [CHECKED].
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
            console.log(`[scan] Claude extracted: ${raw.slice(0, 300)}`);

          } catch(azureErr) {
            console.error(`[scan] Azure hybrid failed: ${azureErr.message}`);
            results.push({ azureError: azureErr.message }); // visible in response JSON
          }
        }

        // ── FALLBACK: Claude Opus vision (if Azure not configured or failed) ──
        if (!raw) {
          const imageBlock = isPdf
            ? { type:'document', source:{ type:'base64', media_type:'application/pdf', data:b64 } }
            : { type:'image',    source:{ type:'base64', media_type:file.mimeType, data:b64 } };

          const step1Res = await fetch('https://api.anthropic.com/v1/messages', {
            method:'POST',
            headers:{
              'Content-Type':'application/json',
              'x-api-key':process.env.ANTHROPIC_API_KEY,
              'anthropic-version':'2023-06-01',
              'anthropic-beta':'pdfs-2024-09-25',
            },
            body:JSON.stringify({
              model:'claude-opus-4-6',
              max_tokens:800,
              messages:[{ role:'user', content:[
                imageBlock,
                { type:'text', text:'Describe this water testing COC form: Lab Use Only box (date/time received), Report To company name, Well Owner address, Date/Time Sampled, and which checkboxes have marks. Be literal and specific.' }
              ]}],
            }),
          });
          if (!step1Res.ok) throw new Error(`Claude vision: ${step1Res.status}`);
          const description = (await step1Res.json()).content?.find(c=>c.type==='text')?.text || '';

          const step2Res = await fetch('https://api.anthropic.com/v1/messages', {
            method:'POST',
            headers:{
              'Content-Type':'application/json',
              'x-api-key':process.env.ANTHROPIC_API_KEY,
              'anthropic-version':'2023-06-01',
            },
            body:JSON.stringify({
              model:'claude-opus-4-6',
              max_tokens:600,
              system:'You are a JSON extraction API. Output ONLY a valid JSON object. No markdown, no explanation.',
              messages:[{ role:'user', content:
`Extract from this COC form description. Known business clients: ${aliasCtx}

TWO FORM TYPES EXIST:
- BUSINESS form: Report To = a company (inspector, water company). Well Owner section has property address.
- PUBLIC form: Report To = "Chanalytical Laboratories" (the lab itself). Has "CUSTOMER & PROPERTY INFORMATION" section with submitter's name/address/email/phone.

For PUBLIC forms: customer = person's name, location/city/state/zip = from CUSTOMER & PROPERTY INFORMATION section.
For BUSINESS forms: customer = company name from Report To, location = from WELL OWNER section only.

Description: ${description}
Return ONLY: {"barcodeId":"","customer":"","email":"","dateDrawn":"","timeDrawn":"","receivedDate":"","receivedTime":"","location":"","city":"","state":"ME","zip":"","tests":[],"individualElements":[],"hasRadon":false,"notes":"","waterType":"","confidence":0}`
              }],
            }),
          });
          if (!step2Res.ok) throw new Error(`Claude step2: ${step2Res.status}`);
          raw = (await step2Res.json()).content?.find(c=>c.type==='text')?.text || '';
        }

        // Parse JSON from OCR output
        let ocr;
        try {
          const s = raw.indexOf('{'), e = raw.lastIndexOf('}');
          if (s < 0 || e < 0) throw new Error('no JSON braces');
          ocr = JSON.parse(raw.slice(s, e+1));
        } catch {
          throw new Error(`OCR JSON parse failed: ${raw.slice(0,200)}`);
        }

        // Validate OCR has something useful — log warning if completely empty
        if (!ocr.customer && !ocr.location && !ocr.dateDrawn && !ocr.tests?.length) {
          console.warn(`[scan] OCR returned empty for ${file.name} — writing to queue for manual review`);
          ocr.confidence = 0;
        }

        // Normalize ALL-CAPS handwriting to Title Case
        ocr.customer = normalizeCase(ocr.customer);
        ocr.location = normalizeCase(ocr.location);
        ocr.city     = normalizeCase(ocr.city);

        // Remove periods and commas from street address
        if (ocr.location) ocr.location = ocr.location.replace(/[.,]+/g, ' ').replace(/\s+/g, ' ').trim();

        // Fix common OCR misreads in addresses
        if (ocr.location) {
          ocr.location = ocr.location
            .replace(/\bwas\b/gi, 'Way')   // "Way" misread as "was"
            .replace(/\bSt\b(?!\s*\d)/g, 'St') // keep St abbreviation
            .replace(/\s+/g, ' ').trim();
        }

        // Abbreviate road type words in address
        if (ocr.location) {
          const abbrevMap = {
            '\\bRoad\\b': 'Rd', '\\bStreet\\b': 'St', '\\bAvenue\\b': 'Ave',
            '\\bDrive\\b': 'Dr', '\\bLane\\b': 'Ln', '\\bTrail\\b': 'Trl',
            '\\bBoulevard\\b': 'Blvd', '\\bCourt\\b': 'Ct', '\\bPlace\\b': 'Pl',
            '\\bCircle\\b': 'Cir', '\\bHighway\\b': 'Hwy', '\\bParkway\\b': 'Pkwy',
            '\\bRoute\\b': 'Rte', '\\bTerrace\\b': 'Ter', '\\bExtension\\b': 'Ext',
          };
          for (const [pattern, abbrev] of Object.entries(abbrevMap)) {
            ocr.location = ocr.location.replace(new RegExp(pattern, 'gi'), abbrev);
          }
        }

        // Fix zip for ME/NH — first digit must be 0 (e.g. 04253 not 84253)
        if (ocr.zip) {
          const z = String(ocr.zip).replace(/\D/g,'').padStart(5,'0');
          if ((ocr.state === 'ME' || ocr.state === 'NH') && !z.startsWith('0')) {
            ocr.zip = '0' + z.slice(1);
            console.log(`[scan] Corrected zip: ${z} → ${ocr.zip}`);
          } else {
            ocr.zip = z;
          }
        }

        // Barcode lookup — merge pre-registered data if found
        let barcodeMatch = null;
        let reviewStatus = 'Ready to Review';
        let ocrStatus    = 'Complete';

        if (ocr.barcodeId?.startsWith('CHA-')) {
          barcodeMatch = await lookupBarcode(sheets, ocr.barcodeId);
          if (barcodeMatch) {
            ocrStatus = 'Barcode Match';
            if (!ocr.customer)       ocr.customer  = barcodeMatch.customer;
            if (!ocr.email)          ocr.email     = barcodeMatch.email;
            if (!ocr.location)       ocr.location  = barcodeMatch.location;
            if (!ocr.city)           ocr.city      = barcodeMatch.city;
            if (!ocr.state)          ocr.state     = barcodeMatch.state;
            if (!ocr.zip)            ocr.zip       = barcodeMatch.zip;
            if (!ocr.tests?.length && barcodeMatch.services) {
              ocr.tests = barcodeMatch.services.split(';').map(s=>s.trim()).filter(Boolean);
            }
          }
        }

        const client      = matchClient(ocr.customer, clients);
        console.log(`[scan] customer: ocr="${ocr.customer}" → matched="${client?.clientName || 'none'}"`);

        // ── AIO (All In One) test conversion ──────────────────────────────
        // If customer is All In One, convert FHA → AIO FHA, Portability → AIO Portability
        const clientName = client?.clientName || ocr.customer || '';
        const isAIO = clientName.toLowerCase().includes('all in one') || clientName.toLowerCase().includes('aio');
        if (isAIO) {
          ocr.tests = (ocr.tests || []).map(t => {
            if (/^basic safety(\s*\(fha\))?$/i.test(t)) return 'AIO FHA';
            if (/portability/i.test(t))                  return 'AIO Portability';
            return t;
          });
          console.log(`[scan] AIO client detected — converted tests: ${ocr.tests.join(', ')}`);
        }        const pkgTests    = ocr.tests || [];
        const indElements = ocr.individualElements || [];
        const radonTests  = ocr.hasRadon ? ['Radon Water'] : [];
        const tests       = [...pkgTests, ...indElements, ...radonTests];
        const valError    = validateTests(pkgTests);
        if (valError) reviewStatus = 'Validation Error';

        const stamp = `${todayDisplayET()} ${nowTimeET()}`;
        const waterTypeSuffix  = ocr.waterType ? ` - ${ocr.waterType}` : '';
        const locationWithType = (ocr.location || '') + waterTypeSuffix;

        // Write to Review Queue
        await sheets.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: `'${SHEETS.REVIEW_QUEUE}'!A:V`,
          valueInputOption: 'RAW',
          insertDataOption: 'INSERT_ROWS',
          requestBody: { values: [[
            ocrStatus,
            reviewStatus,
            '',                                                          // C: Lab ID (assigned at approval)
            client ? client.clientName : (ocr.customer||''),            // D: Client Name
            locationWithType,                                            // E: Address
            ocr.city        ||'',                                        // F: City
            ocr.state       ||'ME',                                      // G: State
            ocr.zip ? String(ocr.zip).padStart(5,'0') : '',             // H: Zip
            client ? client.email : (ocr.email||''),                    // I: Email
            ocr.dateDrawn   ||'',                                        // J: Sample Date
            ocr.timeDrawn   ||'',                                        // K: Sample Time
            ocr.receivedDate||'',                                        // L: Received Date
            ocr.receivedTime||'',                                        // M: Received Time
            tests.join(', '),                                            // N: Tests
            valError        ||'',                                        // O: Validation Errors
            ocr.confidence  ||0,                                         // P: Confidence
            file.id,                                                     // Q: File ID
            stamp,                                                       // R: Processed Date
            ocr.barcodeId   ||'',                                        // S: Barcode ID
            scannedByName,                                               // T: Scanned By
            '',                                                          // U: Approved By
            ocr.waterType   ||'',                                        // V: Water Type
          ]]},
        });

        results.push({
          fileName:    file.name,
          barcodeId:   ocr.barcodeId||'',
          barcodeMatch: !!barcodeMatch,
          client:      client?.clientName||ocr.customer,
          tests,
          confidence:  ocr.confidence,
        });

        console.log(`[scan] ✓ ${file.name} | ${client?.clientName||ocr.customer} | ${tests.join(',')} | ${ocr.confidence}%`);

      } catch(err) {
        console.error(`[scan] ✗ ${file.name}: ${err.message}`);
        console.error(`[scan] Moving ${file.name} back to INCOMING`);
        await moveFile(drive, file.id, FOLDERS.INCOMING); // Move back on failure
      }
    }

    return {statusCode:200, body:JSON.stringify({checked:files.length, processed:results.length, results})};

  } catch(err) {
    console.error('[scan-folder]', err);
    return {statusCode:500, body:JSON.stringify({error:err.message})};
  }
};
