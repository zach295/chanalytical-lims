const fs = require('fs');
const path = 'api/src/approve-scan.js';
let s = fs.readFileSync(path, 'utf8');

const oldWrite = `  const wbBase = \`${GRAPH}/sites/${siteId}/drive/items/${rcsFileId}/workbook\`;
  const sesRes = await fetch(\`${wbBase}/createSession\`, {
    method: 'POST', headers: { ...authHdr, 'Content-Type': 'application/json' },
    body: JSON.stringify({ persistChanges: true }),
  });
  const { id: sid } = await sesRes.json();
  const wbHdr = { ...authHdr, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
  try {
    const sheetsRes = await fetch(\`${wbBase}/worksheets\`, { headers: wbHdr });
    const wsId = ((await sheetsRes.json()).value || [])[0]?.id;
    if (!wsId) throw new Error('No worksheets in RCS file');
    const colARes  = await fetch(\`${wbBase}/worksheets/${wsId}/range(address='A1:A200')?$select=values\`, { headers: wbHdr });
    const colAVals = (await colARes.json()).values || [];
    let targetRow  = 2;
    for (let i = 1; i < colAVals.length; i++) {
      if (!String(colAVals[i][0] || '').trim()) { targetRow = i + 1; break; }
      targetRow = i + 2;
    }
    if (context) context.log(\`[RCS] colA rows=${colAVals.length} targetRow=${targetRow}\`);
    const todayET = new Date().toLocaleDateString('en-US', {
      timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: 'numeric'
    });
    // Use the reviewed/corrected dates from the approval card. Previously column G
    // always used today's approval date, so correcting Date Received before approval
    // never reached the Radon Control Sheet.
    const drawnFmt    = fmtExcel(dateDrawn) || dateDrawn || '';
    const receivedFmt = fmtExcel(receivedDate) || receivedDate || todayET;
    const writeRes = await fetch(\`${wbBase}/worksheets/${wsId}/range(address='A${targetRow}:G${targetRow}')\`, {
      method: 'PATCH', headers: wbHdr,
      body: JSON.stringify({ values: [[newLabId, '', '', '', drawnFmt, timeDrawn || '', receivedFmt]] }),
    });
    if (!writeRes.ok) {
      const errText = await writeRes.text().catch(()=>'');
      throw new Error(\`RCS write failed (${writeRes.status}): ${errText.slice(0,200)}\`);
    }
    if (context) context.log(\`[RCS] Wrote ${newLabId} to ${rcsName} row ${targetRow}\`);
    return { success: true, file: rcsName, row: targetRow };
  } finally {
    await fetch(\`${wbBase}/closeSession\`, { method: 'POST', headers: wbHdr }).catch(() => {});
  }
}`;

const newWrite = `  const wbBase = \`${GRAPH}/sites/${siteId}/drive/items/${rcsFileId}/workbook\`;

  // Excel/Graph can intermittently be busy or slow to commit workbook changes.
  // Retry the workbook session, then verify the Lab ID after every write attempt.
  let sid = '';
  let sessionLastError = '';
  for (let attempt = 1; attempt <= 3 && !sid; attempt++) {
    const started = Date.now();
    try {
      const sesRes = await fetch(\`${wbBase}/createSession\`, {
        method: 'POST', headers: { ...authHdr, 'Content-Type': 'application/json' },
        body: JSON.stringify({ persistChanges: true }),
      });
      const sesData = await sesRes.json().catch(() => ({}));
      if (context) context.log(\`[RCS][timing] createSession attempt ${attempt}: ${Date.now()-started}ms status=${sesRes.status}\`);
      if (sesRes.ok && sesData.id) sid = sesData.id;
      else sessionLastError = \`createSession ${sesRes.status}\`;
    } catch (e) {
      sessionLastError = e.message;
      if (context) context.log(\`[RCS] createSession attempt ${attempt} error: ${e.message}\`);
    }
    if (!sid && attempt < 3) await new Promise(r => setTimeout(r, 750 * attempt));
  }
  if (!sid) throw new Error(\`RCS workbook session failed after 3 attempts: ${sessionLastError || 'unknown error'}\`);

  const wbHdr = { ...authHdr, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
  try {
    const sheetsRes = await fetch(\`${wbBase}/worksheets\`, { headers: wbHdr });
    if (!sheetsRes.ok) throw new Error(\`RCS worksheet list failed (${sheetsRes.status})\`);
    const wsId = ((await sheetsRes.json()).value || [])[0]?.id;
    if (!wsId) throw new Error('No worksheets in RCS file');

    const todayET = new Date().toLocaleDateString('en-US', {
      timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: 'numeric'
    });
    const drawnFmt    = fmtExcel(dateDrawn) || dateDrawn || '';
    const receivedFmt = fmtExcel(receivedDate) || receivedDate || todayET;

    let lastWriteError = '';
    for (let attempt = 1; attempt <= 3; attempt++) {
      // Re-read before every attempt so a prior successful-but-slow write is detected
      // and so retries never create a duplicate Lab ID.
      const readRes = await fetch(\`${wbBase}/worksheets/${wsId}/range(address='A1:A200')?$select=values\`, { headers: wbHdr });
      if (!readRes.ok) {
        lastWriteError = \`RCS column read failed (${readRes.status})\`;
        if (context) context.log(\`[RCS] attempt ${attempt}: ${lastWriteError}\`);
        if (attempt < 3) await new Promise(r => setTimeout(r, 750 * attempt));
        continue;
      }
      const colAVals = (await readRes.json()).values || [];
      const existingRow = colAVals.findIndex((row, i) => i > 0 && String(row?.[0] || '').trim() === newLabId.trim());
      if (existingRow >= 1) {
        if (context) context.log(\`[RCS] ${newLabId} already present in ${rcsName} row ${existingRow + 1}; no duplicate written\`);
        return { success: true, file: rcsName, row: existingRow + 1, alreadyPresent: true };
      }

      let targetRow = 2;
      for (let i = 1; i < colAVals.length; i++) {
        if (!String(colAVals[i]?.[0] || '').trim()) { targetRow = i + 1; break; }
        targetRow = i + 2;
      }
      if (context) context.log(\`[RCS] attempt ${attempt}: colA rows=${colAVals.length} targetRow=${targetRow}\`);

      try {
        const writeRes = await fetch(\`${wbBase}/worksheets/${wsId}/range(address='A${targetRow}:G${targetRow}')\`, {
          method: 'PATCH', headers: wbHdr,
          body: JSON.stringify({ values: [[newLabId, '', '', '', drawnFmt, timeDrawn || '', receivedFmt]] }),
        });
        if (!writeRes.ok) {
          const errText = await writeRes.text().catch(()=>'');
          lastWriteError = \`RCS write failed (${writeRes.status}): ${errText.slice(0,200)}\`;
          if (context) context.log(\`[RCS] attempt ${attempt}: ${lastWriteError}\`);
        } else {
          const verifyRes = await fetch(\`${wbBase}/worksheets/${wsId}/range(address='A1:A200')?$select=values\`, { headers: wbHdr });
          if (verifyRes.ok) {
            const verifyVals = (await verifyRes.json()).values || [];
            const verifiedRow = verifyVals.findIndex((row, i) => i > 0 && String(row?.[0] || '').trim() === newLabId.trim());
            if (verifiedRow >= 1) {
              if (context) context.log(\`[RCS] Wrote and verified ${newLabId} in ${rcsName} row ${verifiedRow + 1}\`);
              return { success: true, file: rcsName, row: verifiedRow + 1, attempt };
            }
            lastWriteError = \`verification missing ${newLabId}\`;
          } else {
            lastWriteError = \`verification read failed (${verifyRes.status})\`;
          }
          if (context) context.log(\`[RCS] attempt ${attempt}: ${lastWriteError}\`);
        }
      } catch (e) {
        lastWriteError = e.message;
        if (context) context.log(\`[RCS] write attempt ${attempt} error: ${e.message}\`);
      }

      if (attempt < 3) await new Promise(r => setTimeout(r, 750 * attempt));
    }

    throw new Error(\`RCS write could not be verified for ${newLabId} after 3 attempts: ${lastWriteError || 'unknown error'}\`);
  } finally {
    await fetch(\`${wbBase}/closeSession\`, { method: 'POST', headers: wbHdr }).catch(() => {});
  }
}`;

if (!s.includes(oldWrite)) throw new Error('Expected Radon Control Sheet write block not found');
s = s.replace(oldWrite, newWrite);

const oldCall = `      let rcsStatus = 'skipped';
      if (radonLabItem) {
        try {
          const rcsResult = await writeRadonControlSheet(
            _siteId, _token, radonLabItem.fullId,
            dateDrawn || '',
            to24h(timeDrawn) || timeDrawn || '',
            receivedDate || '',
            context
          );
          rcsStatus = JSON.stringify(rcsResult);
          context.log(\`[RCS] ${rcsStatus}\`);
        } catch(e) {
          rcsStatus = 'ERROR: ' + e.message;
          context.log('[RCS] Error:', e.message);
        }
      }`;

const newCall = `      let rcsStatus = 'skipped';
      let rcsWarning = '';
      if (radonLabItem) {
        try {
          const rcsResult = await writeRadonControlSheet(
            _siteId, _token, radonLabItem.fullId,
            dateDrawn || '',
            to24h(timeDrawn) || timeDrawn || '',
            receivedDate || '',
            context
          );
          rcsStatus = JSON.stringify(rcsResult);
          context.log(\`[RCS] ${rcsStatus}\`);
        } catch(e) {
          rcsStatus = 'ERROR: ' + e.message;
          rcsWarning = \`⚠️ Sample approved, but Radon Control Sheet write could not be verified for ${radonLabItem.fullId}. Check the RCS manually. ${e.message}\`;
          context.log('[RCS] Error:', e.message);
          context.log('[RCS] Warning:', rcsWarning);
        }
      }`;

if (!s.includes(oldCall)) throw new Error('Expected RCS approval call block not found');
s = s.replace(oldCall, newCall);

const oldReturn = `          csWarning:  csWarning || undefined,
          coaSheetWarning: coaSheetWarning || undefined,`;
const newReturn = `          csWarning:  csWarning || undefined,
          rcsWarning: rcsWarning || undefined,
          rcsStatus,
          coaSheetWarning: coaSheetWarning || undefined,`;
if (!s.includes(oldReturn)) throw new Error('Expected approval response warning block not found');
s = s.replace(oldReturn, newReturn);

if (!s.includes('RCS write could not be verified for ${newLabId} after 3 attempts')) throw new Error('RCS retry/verify patch missing');
if (!s.includes('rcsWarning: rcsWarning || undefined')) throw new Error('RCS warning response missing');

fs.writeFileSync(path, s);
console.log('Radon Control Sheet retry, duplicate prevention, verification, and warning response added');
