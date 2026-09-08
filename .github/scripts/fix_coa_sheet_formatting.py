from pathlib import Path

# --- approve-scan.js ---
p = Path('api/src/approve-scan.js')
s = p.read_text()

start = s.index('async function writeToGoogleSheet(rows, context) {')
end = s.index('\n}\n\n// Module-level cache for list IDs', start) + 2
new_func = r'''async function writeToGoogleSheet(rows, context) {
  // Identify each expected COA row by Lab ID + displayed test name. This makes
  // retries idempotent if Google accepted a write but the verification call failed.
  const rowKey = row => `${String(row?.[7] || '').trim()}|${String(row?.[12] || '').trim()}`;
  const expectedKeys = rows.map(rowKey);
  let lastError = null;

  for (let attempt = 1; attempt <= 3; attempt++) {
    const attemptStart = Date.now();
    try {
      const authStart = Date.now();
      const token = await getSheetsToken();
      context.log(`[Sheets][timing] attempt ${attempt} auth: ${Date.now()-authStart}ms`);

      // Get the actual row count so we can use existing blank rows rather than
      // appending/inserting new rows at the bottom of the sheet.
      const metaStart = Date.now();
      const metaRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${SHEETS_ID}?fields=sheets(properties(title,gridProperties(rowCount)))`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!metaRes.ok) throw new Error(`Google Sheets metadata lookup failed (${metaRes.status})`);
      const meta = await metaRes.json();
      const sheetMeta = (meta.sheets || []).find(sh => sh.properties?.title === SHEETS_TAB);
      const rowCount = Math.max(Number(sheetMeta?.properties?.gridProperties?.rowCount || 1000), 2);
      context.log(`[Sheets][timing] attempt ${attempt} metadata: ${Date.now()-metaStart}ms rows=${rowCount}`);

      // Read the current working grid. A completely blank A:N row is available.
      const locateStart = Date.now();
      const gridRange = encodeURIComponent(`${SHEETS_TAB}!A1:N${rowCount}`);
      const gridRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${SHEETS_ID}/values/${gridRange}`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      context.log(`[Sheets][timing] attempt ${attempt} locate-read: ${Date.now()-locateStart}ms status=${gridRes.status}`);
      if (!gridRes.ok) throw new Error(`Google Sheets row lookup failed (${gridRes.status})`);
      const gridVals = (await gridRes.json()).values || [];

      const existingCounts = new Map();
      for (let r = 1; r < gridVals.length; r++) {
        const key = `${String(gridVals[r]?.[7] || '').trim()}|${String(gridVals[r]?.[12] || '').trim()}`;
        if (key !== '|') existingCounts.set(key, (existingCounts.get(key) || 0) + 1);
      }

      // Only write expected rows that do not already exist from a prior attempt.
      const usedCounts = new Map();
      const missingRows = [];
      for (let i = 0; i < rows.length; i++) {
        const key = expectedKeys[i];
        const already = existingCounts.get(key) || 0;
        const used = usedCounts.get(key) || 0;
        if (used < already) usedCounts.set(key, used + 1);
        else missingRows.push(rows[i]);
      }

      if (missingRows.length) {
        const emptyRowNums = [];
        for (let rowNum = 2; rowNum <= rowCount && emptyRowNums.length < missingRows.length; rowNum++) {
          const vals = gridVals[rowNum - 1] || [];
          const isEmpty = Array.from({ length: 14 }, (_, c) => String(vals[c] || '').trim()).every(v => !v);
          if (isEmpty) emptyRowNums.push(rowNum);
        }
        if (emptyRowNums.length < missingRows.length) {
          throw new Error(`COA sheet does not have enough existing empty rows (${emptyRowNums.length} available, ${missingRows.length} needed)`);
        }

        const writeStart = Date.now();
        const writeData = missingRows.map((row, i) => ({
          range: `${SHEETS_TAB}!A${emptyRowNums[i]}:N${emptyRowNums[i]}`,
          values: [row],
        }));
        const writeRes = await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${SHEETS_ID}/values:batchUpdate`,
          {
            method: 'POST',
            headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
            // RAW preserves ZIP leading zeroes and exact MM-DD-YY text formatting.
            body: JSON.stringify({ valueInputOption: 'RAW', data: writeData }),
          }
        );
        context.log(`[Sheets][timing] attempt ${attempt} write existing rows ${emptyRowNums.join(',')}: ${Date.now()-writeStart}ms status=${writeRes.status}`);
        if (!writeRes.ok) {
          const err = await writeRes.text().catch(()=> '');
          throw new Error(`Google Sheets write failed (${writeRes.status}): ${err.slice(0,300)}`);
        }
      } else {
        context.log(`[Sheets] attempt ${attempt}: all ${rows.length} expected row(s) already present`);
      }

      // Verify Lab ID + Test Name counts, not just Lab ID presence.
      const verifyStart = Date.now();
      const verifyRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${SHEETS_ID}/values/${gridRange}`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      context.log(`[Sheets][timing] attempt ${attempt} verify-read: ${Date.now()-verifyStart}ms status=${verifyRes.status}`);
      if (!verifyRes.ok) throw new Error(`Google Sheets verification read failed (${verifyRes.status})`);
      const verifyVals = (await verifyRes.json()).values || [];
      const verifyCounts = new Map();
      for (let r = 1; r < verifyVals.length; r++) {
        const key = `${String(verifyVals[r]?.[7] || '').trim()}|${String(verifyVals[r]?.[12] || '').trim()}`;
        if (key !== '|') verifyCounts.set(key, (verifyCounts.get(key) || 0) + 1);
      }
      const neededCounts = new Map();
      for (const key of expectedKeys) neededCounts.set(key, (neededCounts.get(key) || 0) + 1);
      const missingKeys = [...neededCounts.entries()]
        .filter(([key, count]) => (verifyCounts.get(key) || 0) < count)
        .map(([key, count]) => `${key} (${verifyCounts.get(key) || 0}/${count})`);
      if (missingKeys.length) throw new Error(`Google Sheets verification missing COA row(s): ${missingKeys.join(', ')}`);

      context.log(`[Sheets] Wrote/verified ${rows.length} COA row(s) using existing empty rows in ${Date.now()-attemptStart}ms`);
      return { success:true, rows:rows.length, attempt };
    } catch (e) {
      lastError = e;
      context.log(`[Sheets] attempt ${attempt}/3 failed after ${Date.now()-attemptStart}ms: ${e.message}`);
      if (attempt < 3) await new Promise(resolve => setTimeout(resolve, attempt * 1500));
    }
  }
  throw lastError || new Error('Google Sheets write failed after 3 attempts');
}'''
s = s[:start] + new_func + s[end:]

# COA date formatting helper + display test abbreviation + public client name + blank report date.
old = r'''        const pad = n => String(n).padStart(2, '0');
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
'''
new = r'''        const pad = n => String(n).padStart(2, '0');
        const toCoaDate = value => {
          const raw = String(value || '').trim();
          if (!raw) return '';
          let m = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
          if (m) return `${pad(m[2])}-${pad(m[3])}-${m[1].slice(-2)}`;
          m = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
          if (m) return `${pad(m[1])}-${pad(m[2])}-${m[3].slice(-2)}`;
          return raw;
        };
        const dateRec  = toCoaDate(receivedDate) || `${pad(etNow.getMonth()+1)}-${pad(etNow.getDate())}-${String(etNow.getFullYear()).slice(-2)}`;
        const dateDrawnCoa = toCoaDate(dateDrawn);
        const timeRec  = receivedTime || `${pad(etNow.getHours())}:${pad(etNow.getMinutes())}`;
        const coaDisplayTest = testName => {
          const name = String(testName || '').trim();
          if (name === 'Basic Safety (FHA)') return 'Basic Safety';
          if (name === 'Expanded Safety (Mortgage Test)') return 'Expanded Safety (Mtg Test)';
          return name;
        };
        const coaCustomerName = (() => {
          const current = String(formalName || customer || '').trim();
          if (!usePublic || current.startsWith('Public-')) return current;
          const parts = String(customer || current).trim().split(/\s+/).filter(Boolean);
          if (parts.length >= 2) {
            const last = parts[parts.length - 1];
            const first = parts.slice(0, -1).join(' ');
            return `Public-${last}, ${first}`;
          }
          return current ? `Public-${current}` : current;
        })();
'''
if old not in s:
    raise SystemExit('COA formatting block not found')
s = s.replace(old, new, 1)

s = s.replace("              dateDrawn  || '',\n", "              dateDrawnCoa,\n", 1)
s = s.replace("              formalName || customer || '',\n              clientCode || '',\n              reportDate,\n", "              coaCustomerName,\n              clientCode || '',\n              '', // Report Date is written only when the report is actually sent\n", 1)
s = s.replace("              zip        || '',\n", "              String(zip || '').replace(/[^0-9]/g, '').padStart(5, '0'),\n", 1)

p.write_text(s)

# --- send-report.js: COA Report Date should be exact MM-DD-YY text in column G ---
p2 = Path('api/src/send-report.js')
s2 = p2.read_text()
old_date = "        const reportDate5 = `${pad5(etNow5.getMonth()+1)}/${pad5(etNow5.getDate())}/${String(etNow5.getFullYear()).slice(-2)}`;\n"
new_date = "        const reportDate5 = `${pad5(etNow5.getMonth()+1)}-${pad5(etNow5.getDate())}-${String(etNow5.getFullYear()).slice(-2)}`;\n"
if old_date not in s2:
    raise SystemExit('send-report COA date format target not found')
s2 = s2.replace(old_date, new_date, 1)
s2 = s2.replace("            body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data: updateData5 }),\n", "            body: JSON.stringify({ valueInputOption: 'RAW', data: updateData5 }),\n", 1)
p2.write_text(s2)
