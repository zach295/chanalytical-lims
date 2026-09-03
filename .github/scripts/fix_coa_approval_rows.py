from pathlib import Path

p = Path('api/src/approve-scan.js')
s = p.read_text()
old = '''        const dateRec  = receivedDate || `${pad(etNow.getMonth()+1)}/${pad(etNow.getDate())}/${String(etNow.getFullYear()).slice(-2)}`;
        const timeRec  = receivedTime || `${pad(etNow.getHours())}:${pad(etNow.getMinutes())}`;
        const sheetRows = labItems
          .filter(l => !l.isRejected)
          .map(l => [
            dateRec,
            timeRec,
            dateDrawn  || '',
            timeDrawn  || '',
            formalName || customer || '',
            clientCode || '',
            '',
            l.fullId,
            location   || '',
            city       || '',
            state      || 'ME',
            zip        || '',
            l.coaTest  || tests?.join(', ') || '',
          ]);
'''
new = '''        const dateRec  = receivedDate || `${pad(etNow.getMonth()+1)}/${pad(etNow.getDate())}/${String(etNow.getFullYear()).slice(-2)}`;
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
              .split(/\\s*\\|\\s*/)
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
              l.fullId,
              location   || '',
              city       || '',
              state      || 'ME',
              zip        || '',
              coaTestName,
            ]);
          });
'''
if old not in s:
    raise SystemExit('Google COA row block not found')
p.write_text(s.replace(old, new, 1))
