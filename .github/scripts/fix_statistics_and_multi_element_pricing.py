from pathlib import Path
import re

# 1) Fix approval billing so a combined lab ID is priced as the SUM of all selected tests/elements.
p = Path('api/src/approve-scan.js')
s = p.read_text()
old = """        const pData  = await pRes.json();
        const tLow   = (params.testName||'').toLowerCase().trim();
        const sLow   = (params.suffix||'').toLowerCase().trim();
        const match  = (pData.value||[]).find(i => {
          const svc = (i.fields?.Service||'').toLowerCase().trim();
          const abbr = (i.fields?.CoreAbbr_x002f_Symbol||'').toLowerCase().trim();
          return svc === tLow || abbr === sLow || svc.includes(tLow) || tLow.includes(svc);
        });
        if (match) rate = parseFloat(String(match.fields?.[priceCol]||'').replace(/[$,]/g,'')) || 0;
"""
new = """        const pData  = await pRes.json();
        // A single Lab ID can contain multiple separately ordered elements, e.g.
        // \"Iron, Total | Manganese, Total | Total Hardness\". Price EACH component
        // and sum them instead of accidentally matching only one element.
        const testParts   = String(params.testName || '').split(/\\s*\\|\\s*/).map(v => v.trim()).filter(Boolean);
        const suffixParts = String(params.suffix || '').split(/\\s*,\\s*/).map(v => v.trim()).filter(Boolean);
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
"""
if old not in s:
    raise SystemExit('approve-scan pricing block not found')
s = s.replace(old, new, 1)
p.write_text(s)

# 2) Fix Sample Correction repricing using the same sum-of-components rule.
p = Path('api/src/update-sample.js')
s = p.read_text()
pat = re.compile(r"""\s*const tLow\s*= newTest\.toLowerCase\(\)\.trim\(\);\n\s*const sLow\s*= newSuffix\.toLowerCase\(\)\.trim\(\);\n\s*const pMatch = \(\(await pRes2\.json\(\)\)\.value\|\|\[\]\)\.find\(i => \{\n\s*const svc\s*= \(i\.fields\?\.Service\|\|'\'\)\.toLowerCase\(\)\.trim\(\);\n\s*const abbr = \(i\.fields\?\.CoreAbbr_x002f_Symbol\|\|'\'\)\.toLowerCase\(\)\.trim\(\);\n\s*return svc === tLow \|\| abbr === sLow \|\| svc\.includes\(tLow\) \|\| tLow\.includes\(svc\);\n\s*\}\);\n\s*if \(pMatch\) rate = parseFloat\(String\(pMatch\.fields\?\.\[priceCol\]\|\|'\'\)\.replace\(/\[\$,\]/g,''\)\) \|\| 0;""")
m = pat.search(s)
if not m:
    raise SystemExit('update-sample pricing block not found')
replacement = """
                  const pricingRows = (await pRes2.json()).value || [];
                  const testParts   = newTest.split(/\\s*\\|\\s*/).map(v => v.trim()).filter(Boolean);
                  const suffixParts = newSuffix.split(/\\s*,\\s*/).map(v => v.trim()).filter(Boolean);
                  let totalRate = 0;
                  for (let idx = 0; idx < Math.max(testParts.length, 1); idx++) {
                    const testName = testParts[idx] || newTest;
                    const suffix   = suffixParts[idx] || (testParts.length === 1 ? newSuffix : '');
                    const tLow = testName.toLowerCase();
                    const sLow = suffix.toLowerCase();
                    let pMatch = pricingRows.find(i => (i.fields?.Service || '').toLowerCase().trim() === tLow);
                    if (!pMatch && sLow) pMatch = pricingRows.find(i => (i.fields?.CoreAbbr_x002f_Symbol || '').toLowerCase().trim() === sLow);
                    if (!pMatch) pMatch = pricingRows.find(i => {
                      const svc = (i.fields?.Service || '').toLowerCase().trim();
                      return svc && tLow && (svc.includes(tLow) || tLow.includes(svc));
                    });
                    totalRate += pMatch ? (parseFloat(String(pMatch.fields?.[priceCol] || '').replace(/[$,]/g, '')) || 0) : 0;
                  }
                  rate = parseFloat(totalRate.toFixed(2));"""
s = s[:m.start()] + replacement + s[m.end():]
p.write_text(s)

# 3) Statistics: accession date from Lab ID, unique kits, total revenue across all billing rows, bounded periods.
p = Path('admin-dashboard.html')
s = p.read_text()
old_getdate = """      const getDate = r => {
        // Try explicit date fields first
        const d = toISO(r.reportDate || r.stmtDate || r.dateRec || r.datePaid || '');
        if (d) return d;
        // Fall back to labId prefix: MMDDYY-NNN → 2026-MM-DD
        const m = (r.labId || '').match(/^(\\d{2})(\\d{2})(\\d{2})-/);
"""
new_getdate = """      const getDate = r => {
        // Statistics are based on accession date. The Lab ID prefix is authoritative:
        // MMDDYY-NNN → YYYY-MM-DD. Fall back to received date only for legacy rows.
        const m = (r.labId || '').match(/^(\\d{2})(\\d{2})(\\d{2})-/);
"""
if old_getdate not in s:
    raise SystemExit('statistics getDate header not found')
s = s.replace(old_getdate, new_getdate, 1)
# Replace remainder immediately after lab ID match return/fallback if it still falls back to explicit report dates.
s = s.replace("""        if (m) return `20${m[3]}-${m[1]}-${m[2]}`;
        return '';
      };
""", """        if (m) return `20${m[3]}-${m[1]}-${m[2]}`;
        return toISO(r.dateRec || r.reportDate || r.stmtDate || r.datePaid || '');
      };
""", 1)
old_calc = """      const calc = subset => {
        const income = subset.reduce((s, r) => s + (parseFloat(r.amt) || 0), 0);
        const kits   = subset.length;
        return { income, kits, avg: kits > 0 ? income / kits : 0 };
      };
"""
new_calc = """      const baseLabId = r => {
        const raw = String(r.labId || '').trim();
        const m = raw.match(/(\\d{6}-\\d{3})/);
        return m ? m[1] : raw;
      };
      const calc = subset => {
        // Revenue is the sum of ALL billing lines; kit count is UNIQUE accession IDs.
        // This keeps FE + MN + HRD on one Lab ID as one kit while counting all revenue.
        const income = subset.reduce((sum, r) => sum + (parseFloat(r.amt) || 0), 0);
        const kits   = new Set(subset.map(baseLabId).filter(Boolean)).size;
        return { income, kits, avg: kits > 0 ? income / kits : 0 };
      };
"""
if old_calc not in s:
    raise SystemExit('statistics calc block not found')
s = s.replace(old_calc, new_calc, 1)
old_filters = """      const ws = calc(rows.filter(r => getDate(r) >= weekISO));
      const ms = calc(rows.filter(r => getDate(r) >= monthISO));
      const ys = calc(rows.filter(r => getDate(r) >= yearISO));
"""
new_filters = """      const todayISO = toISO(fmt(now));
      const inRange = (r, start) => {
        const d = getDate(r);
        return d && d >= start && d <= todayISO;
      };
      const ws = calc(rows.filter(r => inRange(r, weekISO)));
      const ms = calc(rows.filter(r => inRange(r, monthISO)));
      const ys = calc(rows.filter(r => inRange(r, yearISO)));
"""
if old_filters not in s:
    raise SystemExit('statistics filter block not found')
s = s.replace(old_filters, new_filters, 1)
p.write_text(s)

print('Patched approval pricing, Sample Correction pricing, and Statistics calculations.')
