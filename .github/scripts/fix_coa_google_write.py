from pathlib import Path
p=Path('api/src/approve-scan.js')
s=p.read_text()
old="""async function writeToGoogleSheet(rows, context) {
  try {
    const token = await getSheetsToken();
    const range = encodeURIComponent(`${SHEETS_TAB}!A:N`);
    const res = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SHEETS_ID}/values/${range}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
      {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: rows }),
      }
    );
    if (!res.ok) {
      const err = await res.text();
      context.log('[Sheets] Write failed:', err.slice(0, 200));
    } else {
      context.log('[Sheets] Wrote', rows.length, 'row(s) to Google Sheet');
    }
  } catch(e) { context.log('[Sheets] Error (non-fatal):', e.message); }
}
"""
new="""async function writeToGoogleSheet(rows, context) {
  const token = await getSheetsToken();
  const range = encodeURIComponent(`${SHEETS_TAB}!A:N`);
  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SHEETS_ID}/values/${range}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: rows }),
    }
  );
  if (!res.ok) {
    const err = await res.text().catch(()=>'');
    throw new Error(`Google Sheets write failed (${res.status}): ${err.slice(0,300)}`);
  }
  const data = await res.json().catch(()=>({}));
  context.log('[Sheets] Wrote', rows.length, 'row(s) to Google Sheet', data?.updates?.updatedRange || '');
  return { success:true, updatedRange:data?.updates?.updatedRange || '', rows:rows.length };
}
"""
if old not in s: raise SystemExit('writeToGoogleSheet block not found')
s=s.replace(old,new,1)
old2="""      // ── Write to Google Sheet ────────────────────────────────────────────────
      try {
"""
new2="""      // ── Write to Google Sheet ────────────────────────────────────────────────
      let coaSheetWarning = '';
      let coaSheetStatus = 'skipped';
      try {
"""
if old2 not in s: raise SystemExit('google section header not found')
s=s.replace(old2,new2,1)
old3="""        if (sheetRows.length) {
          // Fire-and-forget with 8s timeout — never blocks billing or activity log
          Promise.race([
            writeToGoogleSheet(sheetRows, context),
            new Promise((_, rej) => setTimeout(() => rej(new Error('Sheets timeout')), 8000)),
          ]).catch(e => context.log('[Sheets] Non-fatal:', e.message));
        }
      } catch(e) { context.log('[Sheets] Row build error:', e.message); }
"""
new3="""        if (sheetRows.length) {
          // Await the write before returning from the Azure Function. Fire-and-forget work
          // can be terminated as soon as the request completes, which caused intermittent/missed COA rows.
          const sheetResult = await Promise.race([
            writeToGoogleSheet(sheetRows, context),
            new Promise((_, rej) => setTimeout(() => rej(new Error('Google Sheets write timed out after 12 seconds')), 12000)),
          ]);
          coaSheetStatus = `written:${sheetResult.rows}`;
        }
      } catch(e) {
        coaSheetStatus = 'failed';
        coaSheetWarning = e.message || 'Google Sheets write failed';
        context.log('[Sheets] COA write failed:', coaSheetWarning);
      }
"""
if old3 not in s: raise SystemExit('fire and forget block not found')
s=s.replace(old3,new3,1)
old4="""          `COA scan archived | Review Queue row deleted`,
"""
new4="""          `COA scan archived | Review Queue row deleted`,
          `COA Google Sheet: ${coaSheetStatus}${coaSheetWarning ? ' — ' + coaSheetWarning : ''}`,
"""
if old4 not in s: raise SystemExit('activity detail block not found')
s=s.replace(old4,new4,1)
old5="""          csWarning:  csWarning || undefined,
"""
new5="""          csWarning:  csWarning || undefined,
          coaSheetWarning: coaSheetWarning || undefined,
          coaSheetStatus,
"""
if old5 not in s: raise SystemExit('response warning block not found')
s=s.replace(old5,new5,1)
p.write_text(s)
print('Patched COA Google Sheet write to await completion and surface failures.')
