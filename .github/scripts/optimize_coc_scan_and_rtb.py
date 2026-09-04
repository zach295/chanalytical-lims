from pathlib import Path

# ---- scan-folder.js ----
p = Path('api/src/scan-folder.js')
s = p.read_text()

s = s.replace("// ── PRIMARY: Azure Document Intelligence + Claude Sonnet ─────────────", "// ── PRIMARY: Azure Document Intelligence + Claude Haiku ─────────────", 1)

# File timing setup
old = """      for (let _fi = 0; _fi < toProcess.length; _fi++) {
        const file = toProcess[_fi];
        // Add delay between files to avoid Azure DI rate limiting
        if (_fi > 0) await new Promise(r => setTimeout(r, 2000));
        try {
          // Move to REVIEW immediately to prevent duplicate processing
          await moveSpFile(file.id, SCAN_REVIEW, token);

          // Download file as Buffer → base64 for Azure Doc Intel
          const buf  = await downloadSpFile(file.id, token);
"""
new = """      for (let _fi = 0; _fi < toProcess.length; _fi++) {
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
"""
if old not in s: raise SystemExit('scan timing setup block not found')
s = s.replace(old,new,1)

# Azure timing
old = """            try {
              // Start Azure analysis
              const analyzeUrl = `${endpoint}/documentintelligence/documentModels/prebuilt-layout:analyze?api-version=2024-11-30`;
"""
new = """            try {
              const azureStartedAt = Date.now();
              // Start Azure analysis
              const analyzeUrl = `${endpoint}/documentintelligence/documentModels/prebuilt-layout:analyze?api-version=2024-11-30`;
"""
if old not in s: raise SystemExit('azure timing start block not found')
s=s.replace(old,new,1)

old = """              scanLog.push('OK Azure succeeded');
              context.log(`[scan] STEP 3 OK — Azure succeeded`);
"""
new = """              timing.azureMs = Date.now() - azureStartedAt;
              scanLog.push(`OK Azure succeeded in ${timing.azureMs}ms`);
              context.log(`[scan] STEP 3 OK — Azure succeeded in ${(timing.azureMs/1000).toFixed(1)}s`);
"""
if old not in s: raise SystemExit('azure timing success block not found')
s=s.replace(old,new,1)

# Primary model + timing
s = s.replace("// Claude Sonnet structures Azure's text into JSON", "// Claude Haiku structures Azure's text into JSON; Sonnet is reserved for recovery", 1)
old = """              const extractRes = await fetch('https://api.anthropic.com/v1/messages', {
"""
new = """              const claudeStartedAt = Date.now();
              const extractRes = await fetch('https://api.anthropic.com/v1/messages', {
"""
if old not in s: raise SystemExit('primary claude call not found')
s=s.replace(old,new,1)

# Only first model occurrence = primary
s = s.replace("model:      'claude-sonnet-4-6',", "model:      'claude-haiku-4-5',", 1)

# Remove whole known client directory from primary prompt
old = """
KNOWN CLIENTS (match exactly if name appears on form, else ""):
${aliasCtx}

RULES:
"""
new = """
RULES:
- Extract only what is present in the OCR text. Do not infer or guess a known client identity; local matching happens after extraction.
"""
if old not in s: raise SystemExit('primary known clients prompt block not found')
s=s.replace(old,new,1)

old = """              const extractData = await extractRes.json();
              raw = extractData.content?.find(c => c.type === 'text')?.text || '';
              scanLog.push(`Claude raw: ${raw.length}chars`);
              context.log(`[scan] STEP 5 — Claude raw response length: ${raw.length}`);
"""
new = """              const extractData = await extractRes.json();
              raw = extractData.content?.find(c => c.type === 'text')?.text || '';
              timing.haikuMs = Date.now() - claudeStartedAt;
              scanLog.push(`Haiku raw: ${raw.length}chars in ${timing.haikuMs}ms`);
              context.log(`[scan] STEP 5 — Haiku response length: ${raw.length} in ${(timing.haikuMs/1000).toFixed(1)}s`);
"""
if old not in s: raise SystemExit('primary response timing block not found')
s=s.replace(old,new,1)

# Broaden Sonnet recovery modestly: only clearly weak extraction
old = """                    if (!testParse.customer && !testParse.tests?.length && testParse.confidence < 30) {
                      context.log('[scan] STEP 5 — Primary extraction empty, retrying with minimal prompt');
"""
new = """                    const weakExtraction = (Number(testParse.confidence || 0) < 50) &&
                      !testParse.customer && !testParse.location &&
                      !testParse.tests?.length && !testParse.individualElements?.length && !testParse.hasRadon;
                    if (weakExtraction) {
                      context.log('[scan] STEP 5 — Haiku extraction weak, retrying once with Sonnet');
                      const sonnetStartedAt = Date.now();
"""
if old not in s: raise SystemExit('retry condition block not found')
s=s.replace(old,new,1)

old = """                        const retryData = await retryRes.json();
                        const retryRaw = retryData.content?.find(c => c.type === 'text')?.text || '';
                        if (retryRaw && retryRaw.includes('{')) raw = retryRaw;
                        context.log('[scan] STEP 5 retry — result:', retryRaw.slice(0, 200));
"""
new = """                        const retryData = await retryRes.json();
                        const retryRaw = retryData.content?.find(c => c.type === 'text')?.text || '';
                        timing.sonnetRetryMs = Date.now() - sonnetStartedAt;
                        if (retryRaw && retryRaw.includes('{')) raw = retryRaw;
                        context.log(`[scan] STEP 5 Sonnet retry — ${(timing.sonnetRetryMs/1000).toFixed(1)}s — result:`, retryRaw.slice(0, 200));
"""
if old not in s: raise SystemExit('retry timing block not found')
s=s.replace(old,new,1)

# Review queue + total timing
old = """          await writeToReviewQueue({
"""
new = """          const queueStartedAt = Date.now();
          await writeToReviewQueue({
"""
# replace only first success-path call after customer debug; there is only one before error call in this section
if old not in s: raise SystemExit('review queue call not found')
s=s.replace(old,new,1)

old = """          results.push({
            fileName:     file.name,
"""
new = """          timing.reviewQueueMs = Date.now() - queueStartedAt;
          timing.totalMs = Date.now() - fileStartedAt;
          scanLog.push(`TIMING move=${timing.moveMs||0}ms download=${timing.downloadMs||0}ms azure=${timing.azureMs||0}ms haiku=${timing.haikuMs||0}ms sonnetRetry=${timing.sonnetRetryMs||0}ms reviewQueue=${timing.reviewQueueMs||0}ms total=${timing.totalMs}ms`);
          context.log(`[scan] TIMING ${file.name} — move ${(timing.moveMs||0)/1000}s | download ${(timing.downloadMs||0)/1000}s | Azure ${((timing.azureMs||0)/1000).toFixed(1)}s | Haiku ${((timing.haikuMs||0)/1000).toFixed(1)}s | Sonnet retry ${((timing.sonnetRetryMs||0)/1000).toFixed(1)}s | Queue ${((timing.reviewQueueMs||0)/1000).toFixed(1)}s | TOTAL ${(timing.totalMs/1000).toFixed(1)}s`);

          results.push({
            fileName:     file.name,
"""
if old not in s: raise SystemExit('results timing insertion point not found')
s=s.replace(old,new,1)

# include timing in result for dashboard/API diagnostics
old = """            confidence:   ocr.confidence,
            ocrExtracted: { phone: ocr.phone, billingAddress: ocr.billingAddress, email: ocr.email, customer: ocr.customer },
"""
new = """            confidence:   ocr.confidence,
            timing,
            ocrExtracted: { phone: ocr.phone, billingAddress: ocr.billingAddress, email: ocr.email, customer: ocr.customer },
"""
if old not in s: raise SystemExit('result timing field insertion not found')
s=s.replace(old,new,1)

p.write_text(s)

# ---- approve-scan.js ----
p = Path('api/src/approve-scan.js')
s = p.read_text()

s = s.replace("async function writeReportsToBilled(siteId, token, params, context) {", "async function writeReportsToBilled(siteId, token, params, context, cache = {}) {", 1)

# Cache client lookup
old = """        const cRes = await fetch(
          `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields($select=ClientName,PricingCategory)&$top=500`,
          { headers: authHdr });
        if (cRes.ok) {
          const cData = await cRes.json();
          const cLow  = params.customer.toLowerCase().trim();
          const found = (cData.value||[]).find(i => (i.fields?.ClientName||'').toLowerCase().trim() === cLow);
          pricingCategory = found?.fields?.PricingCategory || '';
        }
"""
new = """        if (!cache.clientPricingRows) {
          const cRes = await fetch(
            `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields($select=ClientName,PricingCategory)&$top=500`,
            { headers: authHdr });
          cache.clientPricingRows = cRes.ok ? ((await cRes.json()).value || []) : [];
        }
        const cLow  = params.customer.toLowerCase().trim();
        const found = cache.clientPricingRows.find(i => (i.fields?.ClientName||'').toLowerCase().trim() === cLow);
        pricingCategory = found?.fields?.PricingCategory || '';
"""
if old not in s: raise SystemExit('RTB client lookup block not found')
s=s.replace(old,new,1)

# Cache pricing rows
old = """      const pRes = await fetch(
        `${GRAPH}/sites/${siteId}/lists/Current%20Pricing-V1/items?$expand=fields&$top=200`,
        { headers: authHdr });
      if (pRes.ok) {
        const pData  = await pRes.json();
"""
new = """      if (!cache.pricingRows) {
        const pRes = await fetch(
          `${GRAPH}/sites/${siteId}/lists/Current%20Pricing-V1/items?$expand=fields&$top=200`,
          { headers: authHdr });
        cache.pricingRows = pRes.ok ? ((await pRes.json()).value || []) : [];
      }
      if (cache.pricingRows.length) {
        const pData  = { value: cache.pricingRows };
"""
if old not in s: raise SystemExit('RTB pricing lookup block not found')
s=s.replace(old,new,1)

# Cache column map
old = """    const colsRes = await fetch(
      `${GRAPH}/sites/${siteId}/lists/${listId}/columns?$select=name,displayName&$top=50`,
      { headers: authHdr }
    );
    const colMap = {}; // displayName → internalName
    if (colsRes.ok) {
      const colData = await colsRes.json();
      (colData.value || []).forEach(c => { colMap[c.displayName] = c.name; });
      if (context) context.log('[RTB] Column map:', JSON.stringify(colMap));
    }
"""
new = """    if (!cache.colMap) {
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
"""
if old not in s: raise SystemExit('RTB column map block not found')
s=s.replace(old,new,1)

# Approval-local cache passed to every RTB row
old = """      const rtbResults = [];
      for (const item of labItems) {
"""
new = """      const rtbResults = [];
      const rtbCache = {}; // pricing/client/column metadata loaded once per approval
      for (const item of labItems) {
"""
if old not in s: raise SystemExit('RTB cache caller insertion not found')
s=s.replace(old,new,1)

old = """          }, context).catch(e => ({ success:false, error:e.message }));
"""
new = """          }, context, rtbCache).catch(e => ({ success:false, error:e.message }));
"""
if old not in s: raise SystemExit('RTB caller cache arg not found')
s=s.replace(old,new,1)

p.write_text(s)
print('Optimized COC scanning and Reports-to-be-Billed metadata access.')
