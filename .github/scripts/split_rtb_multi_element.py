from pathlib import Path

p = Path('api/src/approve-scan.js')
s = p.read_text()
old = """      const rtbResults = [];
      for (const item of labItems) {
        const rtbResult = await writeReportsToBilled(_siteId, _token, {
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
        }, context).catch(e => ({ success:false, error:e.message }));
        rtbResults.push(rtbResult);
        context.log('[RTB]', item.fullId, rtbResult.success ? `row ${rtbResult.row}` : rtbResult.error);
      }
"""
new = """      const rtbResults = [];
      for (const item of labItems) {
        // Match COA behavior: one Reports-to-be-Billed row per separately ordered
        // test/element while keeping the same base Lab ID and sample metadata.
        const rtbTests = String(item.coaTest || '')
          .split(/\\s*\\|\\s*/)
          .map(v => v.trim())
          .filter(Boolean);
        const rtbSuffixes = String(item.suffix || '')
          .split(/\\s*,\\s*/)
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
          }, context).catch(e => ({ success:false, error:e.message }));
          rtbResults.push(rtbResult);
          context.log('[RTB]', item.baseId, rowTest, rtbResult.success ? `id ${rtbResult.id}` : rtbResult.error);
        }
      }
"""
if old not in s:
    raise SystemExit('RTB approval loop not found')
s = s.replace(old, new, 1)
p.write_text(s)
print('Split Reports to be Billed into one row per test/element.')
