const { app } = require('@azure/functions');
const { createItem, listItems, updateItem, deleteItem, LISTS } = require('../shared/graph');
const { syncCoaFromArchivedRows, getCoaRowsByBaseIds, clearCoaRows } = require('../shared/coa-sheet');

// ── Control Sheet Helper ──────────────────────────────────────────────────────
// Finds C_MMDDYY.xlsx in Test C folder and updates the lab ID cell in column A
async function updateControlSheet(siteId, datePrefix, baseId, newLabId, token, context) {
  const GRAPH        = 'https://graph.microsoft.com/v1.0';
  const controlFolder = process.env.SP_CONTROL_FOLDER ||
    '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
  const marker    = 'Shared Documents/';
  const idx       = controlFolder.indexOf(marker);
  const relPath   = idx >= 0 ? controlFolder.slice(idx + marker.length) : controlFolder.replace(/^\/+/, '');
  const authHdr   = { Authorization: `Bearer ${token}` };
  const MONTHS    = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];

  // Build month subfolder: "August 2026" from MMDDYY prefix
  const mm         = parseInt(datePrefix.slice(0, 2)) - 1;
  const yy         = datePrefix.slice(4, 6);
  const year       = '20' + yy;
  const monthName  = MONTHS[mm] || datePrefix.slice(0, 2);
  const monthFolder = `${relPath}/${monthName} ${year}`;
  const fileName   = `C_${datePrefix}.xlsx`;

  // Try month subfolder first, then flat folder
  let fileId = null;
  for (const tryPath of [`${monthFolder}/${fileName}`, `${relPath}/${fileName}`]) {
    const enc = tryPath.split('/').map(encodeURIComponent).join('/');
    const r   = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${enc}`, { headers: authHdr });
    if (r.ok) { fileId = (await r.json()).id; break; }
  }
  if (!fileId) throw new Error(`Control sheet C_${datePrefix}.xlsx not found`);

  // 2. Open session
  const sesRes  = await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook/createSession`,
    { method: 'POST', headers: { ...authHdr, 'Content-Type': 'application/json' },
      body: JSON.stringify({ persistChanges: true }) }
  );
  const sesData = await sesRes.json();
  if (!sesData.id) throw new Error(`Could not open workbook session: ${JSON.stringify(sesData).slice(0,100)}`);
  const sid   = sesData.id;
  const wbHdr = { ...authHdr, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
  const wbBase = `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook`;

  try {
    // 3. Get first worksheet
    const sheetsRes = await fetch(`${wbBase}/worksheets`, { headers: wbHdr });
    const sheets    = (await sheetsRes.json()).value || [];
    if (!sheets.length) throw new Error('No worksheets in control sheet');
    const wsId = sheets[0].id;

    // 4. Read used range to find matching row
    const rangeRes  = await fetch(
      `${wbBase}/worksheets/${wsId}/range(address='A1:A150')`,
      { headers: wbHdr }
    );
    const rangeData = await rangeRes.json();
    const rawVals   = rangeData.values || rangeData.text || [];
    const rows      = Array.isArray(rawVals[0]) ? rawVals : rawVals.map(v => [v]);

    let targetRow = -1;
    for (let i = 0; i < rows.length; i++) {
      const cell     = String(rows[i][0] || '').trim().replace(/\s+/g, ' ');
      const cellBase = cell.split(' ')[0].replace(/[^\w-]/g, '').trim();
      if (cellBase === baseId.replace(/[^\w-]/g,'') || cell.startsWith(baseId)) { targetRow = i + 1; break; }
    }

    if (targetRow < 0) throw new Error(`Lab ID ${baseId} not found in column A of C_${datePrefix}.xlsx (scanned ${rows.length} rows)`);

    // 5. Update the cell with new lab ID
    await fetch(
      `${wbBase}/worksheets/${wsId}/range(address='A${targetRow}')`,
      { method: 'PATCH', headers: wbHdr, body: JSON.stringify({ values: [[newLabId]] }) }
    );

    if (context) context.log(`[controlSheet] Updated A${targetRow}: ${newLabId}`);
    return { updated: true, row: targetRow };
  } finally {
    await fetch(`${wbBase}/closeSession`, { method: 'POST', headers: wbHdr }).catch(() => {});
  }
}

// ── Radon Control Sheet Helper ────────────────────────────────────────────────
async function updateRadonSheet(siteId, datePrefix, baseId, newLabId, token, context) {
  const GRAPH      = 'https://graph.microsoft.com/v1.0';
  const controlFolder = process.env.SP_CONTROL_FOLDER ||
    '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
  const marker     = 'Shared Documents/';
  const idx        = controlFolder.indexOf(marker);
  const relPath    = idx >= 0 ? controlFolder.slice(idx + marker.length) : controlFolder.replace(/^\/+/, '');

  // Build month folder name: "July Radon 2026" from MMDDYY prefix
  const mm   = datePrefix.slice(0, 2);
  const yy   = datePrefix.slice(4, 6);
  const year = '20' + yy;
  const months = ['January','February','March','April','May','June',
                  'July','August','September','October','November','December'];
  const monthName = months[parseInt(mm) - 1] || mm;
  const radonFolder = `${relPath}/${monthName} Radon ${year}`;
  const fileName    = `C_${datePrefix}.xlsx`; // same naming as control sheet
  const filePath    = `${radonFolder}/${fileName}`.split('/').map(encodeURIComponent).join('/');
  const authHdr     = { Authorization: `Bearer ${token}` };

  const fileRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${filePath}`, { headers: authHdr });
  if (!fileRes.ok) {
    if (context) context.log(`[radonSheet] File not found: ${fileName} in ${monthName} Radon ${year}`);
    return { updated: false, reason: 'Radon sheet not found' };
  }
  const { id: fileId } = await fileRes.json();

  // Reuse same logic as control sheet
  const sesRes = await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook/createSession`,
    { method: 'POST', headers: { ...authHdr, 'Content-Type': 'application/json' },
      body: JSON.stringify({ persistChanges: true }) }
  );
  const { id: sid } = await sesRes.json();
  const wbHdr  = { ...authHdr, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
  const wbBase = `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook`;

  try {
    const sheetsRes = await fetch(`${wbBase}/worksheets`, { headers: wbHdr });
    const wsId      = ((await sheetsRes.json()).value || [])[0]?.id;
    if (!wsId) throw new Error('No worksheets in radon sheet');
    const rangeRes  = await fetch(`${wbBase}/worksheets/${wsId}/range(address='A1:A150')`, { headers: wbHdr });
    const rangeData = await rangeRes.json();
    const rawVals   = rangeData.values || rangeData.text || [];
    const rows      = Array.isArray(rawVals[0]) ? rawVals : rawVals.map(v => [v]);
    let   targetRow = -1;
    for (let i = 0; i < rows.length; i++) {
      const cell     = String(rows[i][0] || '').trim().replace(/\s+/g,' ');
      const cellBase = cell.split(' ')[0].replace(/[^\w-]/g,'').trim();
      if (cellBase === baseId.replace(/[^\w-]/g,'') || cell.startsWith(baseId)) { targetRow = i + 1; break; }
    }
    if (targetRow < 0) return { updated: false, reason: `${baseId} not found in radon sheet` };
    await fetch(`${wbBase}/worksheets/${wsId}/range(address='A${targetRow}')`,
      { method: 'PATCH', headers: wbHdr, body: JSON.stringify({ values: [[newLabId]] }) });
    if (context) context.log(`[radonSheet] Updated A${targetRow}: ${newLabId}`);
    return { updated: true, row: targetRow };
  } finally {
    await fetch(`${wbBase}/closeSession`, { method: 'POST', headers: wbHdr }).catch(() => {});
  }
}

async function clearDuplicateControlRow(siteId, datePrefix, baseId, radon, token, context) {
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const controlFolder = process.env.SP_CONTROL_FOLDER || '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
  const marker = 'Shared Documents/';
  const idx = controlFolder.indexOf(marker);
  const relPath = idx >= 0 ? controlFolder.slice(idx + marker.length) : controlFolder.replace(/^\\/+/, '');
  const monthNum = parseInt(datePrefix.slice(0, 2), 10) - 1;
  const year = '20' + datePrefix.slice(4, 6);
  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const monthName = months[monthNum] || datePrefix.slice(0, 2);
  const fileName = 'C_' + datePrefix + '.xlsx';
  const paths = radon
    ? [relPath + '/' + monthName + ' Radon ' + year + '/' + fileName]
    : [relPath + '/' + monthName + ' ' + year + '/' + fileName, relPath + '/' + fileName];
  const auth = { Authorization:'Bearer ' + token, 'Content-Type':'application/json' };

  let fileId = null;
  for (const p of paths) {
    const enc = p.split('/').map(encodeURIComponent).join('/');
    const fr = await fetch(GRAPH + '/sites/' + siteId + '/drive/root:/' + enc, { headers:auth });
    if (fr.ok) { fileId = (await fr.json()).id; break; }
  }
  if (!fileId) return { cleared:false, reason:radon ? 'Radon control sheet not found' : 'Control sheet not found' };

  const sr = await fetch(GRAPH + '/sites/' + siteId + '/drive/items/' + fileId + '/workbook/createSession', {
    method:'POST', headers:auth, body:JSON.stringify({ persistChanges:true })
  });
  const sd = await sr.json().catch(()=>({}));
  if (!sr.ok || !sd.id) throw new Error('Could not open ' + (radon ? 'radon ' : '') + 'control sheet session');
  const wbHdr = { ...auth, 'workbook-session-id':sd.id };
  const wbBase = GRAPH + '/sites/' + siteId + '/drive/items/' + fileId + '/workbook';

  try {
    const wr = await fetch(wbBase + '/worksheets', { headers:wbHdr });
    const ws = ((await wr.json()).value || [])[0];
    if (!ws) throw new Error('No worksheet found');
    const ar = await fetch(wbBase + "/worksheets/" + ws.id + "/range(address='A1:A250')?$select=values", { headers:wbHdr });
    if (!ar.ok) throw new Error('Control sheet row lookup failed (' + ar.status + ')');
    const vals = (await ar.json()).values || [];
    let row = -1;
    for (let i = 0; i < vals.length; i++) {
      const cell = String(vals[i]?.[0] || '').trim();
      const cellBase = cell.split(' ')[0].replace(/[^\\w-]/g, '').trim();
      if (cellBase === baseId.replace(/[^\\w-]/g, '') || cell.startsWith(baseId)) { row = i + 1; break; }
    }
    if (row < 0) return { cleared:false, reason:baseId + ' not found' };
    const endCol = radon ? 'G' : 'AE';
    const width = radon ? 7 : 31;
    const range = "A" + row + ":" + endCol + row;
    const cr = await fetch(wbBase + "/worksheets/" + ws.id + "/range(address='" + range + "')", {
      method:'PATCH', headers:wbHdr, body:JSON.stringify({ values:[Array(width).fill('')] })
    });
    if (!cr.ok) throw new Error((radon ? 'Radon ' : '') + 'control sheet clear failed (' + cr.status + ')');
    if (context) context.log('[duplicate] Cleared ' + (radon ? 'radon ' : '') + 'control sheet row ' + row + ' for ' + baseId);
    return { cleared:true, row };
  } finally {
    await fetch(wbBase + '/closeSession', { method:'POST', headers:wbHdr }).catch(()=>{});
  }
}
app.http('reject-sample', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { labId, rejectionType, reason, rejectedBy, duplicateOf } = await request.json();
      const isDuplicate = rejectionType === 'Rejected - Duplicate';
      if (!labId)          return { status: 400, body: JSON.stringify({ error: 'labId required' }) };
      if (!rejectionType)  return { status: 400, body: JSON.stringify({ error: 'rejectionType required' }) };
      if (!reason?.trim() && !isDuplicate) return { status: 400, body: JSON.stringify({ error: 'reason required' }) };

      const siteId = process.env.SP_SITE_ID;
      const { getToken } = require('../shared/graph');
      const token  = await getToken();
      const now    = new Date().toISOString();
      const baseId = labId.split(' ')[0].trim();
      const cleanReason = String(reason || '').trim();
      const cleanDuplicateOf = String(duplicateOf || '').trim();
      const rejNote = isDuplicate
        ? ['Duplicate', cleanDuplicateOf && `Duplicate of ${cleanDuplicateOf}`, cleanReason].filter(Boolean).join(': ')
        : `${rejectionType}: ${cleanReason}`;
      const rejLabId = isDuplicate ? `${baseId} Dup` : `${baseId} REJ`;
      const archivedService = isDuplicate ? 'Duplicate' : rejectionType;
      const log = [];

      // 1. Write to Rejected list
      // field_1=LabId, field_2=RejectionType, field_3=Reason, field_4=RejectedBy
      await createItem(LISTS.REJECTED, {
        Title:   now,
        field_1: labId,
        field_2: rejectionType,
        field_3: isDuplicate
          ? [cleanDuplicateOf && `Duplicate of ${cleanDuplicateOf}`, cleanReason].filter(Boolean).join(' | ')
          : cleanReason,
        field_4: rejectedBy || 'Lab Staff',
      }).catch(e => context.log('[Rejected] Error:', e.message));
      log.push('✅ Written to Rejected list');

      // ── Control sheets ───────────────────────────────────────────────────
      const datePrefix = baseId.slice(0, 6);
      const archivedAll = await listItems(LISTS.ARCHIVED_INTAKE, { top: 2000 }).catch(() => []);
      const isRadonSample = archivedAll.some(r =>
        r.field_1?.startsWith(baseId) && (r.field_2||'').toLowerCase().includes('radon')
      );

      if (isDuplicate) {
        // Duplicate cleanup removes the operational row entirely rather than
        // relabeling it. Always attempt both regular and radon control sheets.
        try {
          const cs = await clearDuplicateControlRow(siteId, datePrefix, baseId, false, token, context);
          log.push(cs.cleared ? `✅ Control sheet row cleared (row ${cs.row})` : `ℹ️ Control sheet: ${cs.reason}`);
        } catch(e) { log.push(`⚠️ Control sheet: ${e.message}`); }
        try {
          const rs = await clearDuplicateControlRow(siteId, datePrefix, baseId, true, token, context);
          log.push(rs.cleared ? `✅ Radon control sheet row cleared (row ${rs.row})` : `ℹ️ Radon control sheet: ${rs.reason}`);
        } catch(e) { log.push(`⚠️ Radon control sheet: ${e.message}`); }
      } else {
        try {
          const csResult = await updateControlSheet(siteId, datePrefix, baseId, rejLabId, token, context);
          log.push(csResult.updated ? `✅ Control sheet updated to REJ (row ${csResult.row})` : `ℹ️ Control sheet: ${csResult.reason || 'row not found'}`);
        } catch(e) { log.push(`⚠️ Control sheet: ${e.message}`); }
        if (isRadonSample) {
          try {
            const rwResult = await updateRadonSheet(siteId, datePrefix, baseId, rejLabId, token, context);
            log.push(rwResult.updated ? '✅ Radon sheet updated to REJ' : `ℹ️ Radon: ${rwResult.reason}`);
          } catch(e) { log.push(`⚠️ Radon sheet: ${e.message}`); }
        }
      }
      // ── Delete from Review Queue ─────────────────────────────────────────
      try {
        const rqItems = await listItems(LISTS.REVIEW_QUEUE, { top: 500 });
        const rqRow   = rqItems.find(r => {
          const bc = (r.BarcodeID || r.barcodeId || '').split(' ')[0].trim();
          return bc === baseId;
        });
        if (rqRow) {
          await deleteItem(LISTS.REVIEW_QUEUE, rqRow._id);
          log.push(`✅ Review Queue row deleted`);
        }
      } catch(e) { log.push(`⚠️ Review Queue: ${e.message}`); }

      // ── Write to Activity Log ─────────────────────────────────────────────
      const actNow = new Date();
      const dateStr = actNow.toLocaleDateString('en-US', { timeZone: 'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
      const timeStr = actNow.toLocaleTimeString('en-US', { timeZone: 'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
      try {
        const updateResults = log.filter(l => !l.includes('Written to Activity Log')).join(' | ');
        const fullNotes = isDuplicate
          ? [
              'Type: Duplicate',
              cleanDuplicateOf && `Duplicate of: ${cleanDuplicateOf}`,
              cleanReason && `Reason: ${cleanReason}`,
              updateResults && `Updates: ${updateResults}`,
            ].filter(Boolean).join('\n')
          : [`Rejection Type: ${rejectionType}`, `Reason: ${cleanReason}`, updateResults && `Updates: ${updateResults}`]
              .filter(Boolean).join('\n');
        await createItem('Activity Log', {
          Title:        `${dateStr} ${labId}`,
          Client:       labId,
          ActivityType: isDuplicate ? 'Duplicate Sample' : 'Sample Rejected',
          Notes:        fullNotes.slice(0, 3000),
          By:           rejectedBy || 'Lab Staff',
          LogDate:      dateStr,
          LogTime:      timeStr,
          Quantity:     0,
        });
        log.push('✅ Written to Activity Log');
      } catch(e) {
        context.log('[ActivityLog] Error:', e.message);
        log.push('⚠️ Activity Log write failed: ' + e.message);
      }

      // ── Delete from Results Cache ─────────────────────────────────────────
      try {
        const GRAPH   = 'https://graph.microsoft.com/v1.0';
        const authHdr = { Authorization: `Bearer ${token}` };
        const rcListRes = await fetch(`${GRAPH}/sites/${siteId}/lists?$select=id,displayName`, { headers: authHdr });
        const rcListId  = ((await rcListRes.json()).value || []).find(l => l.displayName === 'Results Cache')?.id;
        if (rcListId) {
          const rcItemsRes = await fetch(
            `${GRAPH}/sites/${siteId}/lists/${rcListId}/items?$expand=fields($select=LabID)&$top=500`,
            { headers: authHdr }
          );
          const rcItem = ((await rcItemsRes.json()).value || [])
            .find(i => String(i.fields?.LabID || '').split(' ')[0].trim() === baseId);
          if (rcItem) {
            await fetch(
              `${GRAPH}/sites/${siteId}/lists/${rcListId}/items/${rcItem.id}`,
              { method: 'DELETE', headers: authHdr }
            );
            log.push(`✅ Results Cache entry deleted for ${baseId}`);
          }
        }
      } catch(e) { log.push(`⚠️ Results Cache delete: ${e.message}`); }

      // ── Update Accession Log ──────────────────────────────────────────────
      try {
        const GRAPH   = 'https://graph.microsoft.com/v1.0';
        const authHdr = { Authorization: `Bearer ${token}` };
        const accListRes = await fetch(`${GRAPH}/sites/${siteId}/lists?$select=id,displayName`, { headers: authHdr });
        const accListId  = ((await accListRes.json()).value || []).find(l => l.displayName === 'Accession Log')?.id;
        if (accListId) {
          const accItemsRes = await fetch(
            `${GRAPH}/sites/${siteId}/lists/${accListId}/items?$expand=fields($select=field_1,field_2,field_3,field_4)&$top=500`,
            { headers: authHdr }
          );
          const accItems = ((await accItemsRes.json()).value || [])
            .filter(i => String(i.fields?.field_1 || '').split(' ')[0].trim() === baseId);
          for (const item of accItems) {
            await fetch(
              `${GRAPH}/sites/${siteId}/lists/${accListId}/items/${item.id}/fields`,
              { method: 'PATCH', headers: { ...authHdr, 'Content-Type': 'application/json' },
                body: JSON.stringify({ field_2: rejLabId, field_3: archivedService, field_4: isDuplicate ? 'Dup' : 'REJ' }) }
            );
          }
          if (accItems.length) log.push(`✅ Accession Log updated (${accItems.length} row(s))`);
          else log.push(`ℹ️ Accession Log: no rows found for ${baseId}`);
        }
      } catch(e) { log.push(`⚠️ Accession Log: ${e.message}`); }


      // Report Date for rejected samples = next business day after the Lab ID date.
      // Lab ID format: MMDDYY-NNN.
      const nextBusinessDayFromLabId = id => {
        const m = String(id || '').match(/^(\d{2})(\d{2})(\d{2})-/);
        if (!m) return '';
        const d = new Date(Number(`20${m[3]}`), Number(m[1]) - 1, Number(m[2]));
        if (isNaN(d.getTime())) return '';
        do {
          d.setDate(d.getDate() + 1);
        } while (d.getDay() === 0 || d.getDay() === 6);
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        const yy = String(d.getFullYear()).slice(-2);
        return `${mm}/${dd}/${yy}`;
      };

      // ── Update Reports to be Billed ───────────────────────────────────────
      try {
        const GRAPH2   = 'https://graph.microsoft.com/v1.0';
        const authHdr2 = { Authorization: `Bearer ${token}` };
        const lRes = await fetch(`${GRAPH2}/sites/${siteId}/lists?$select=id,displayName`, { headers: authHdr2 });
        const rtbListId = ((await lRes.json()).value || []).find(l => l.displayName === 'Reports to be Billed')?.id;
        if (rtbListId) {
          // Get column display name → internal name map
          const colsRes = await fetch(`${GRAPH2}/sites/${siteId}/lists/${rtbListId}/columns?$select=name,displayName&$top=100`, { headers: authHdr2 });
          const colMap  = {};
          if (colsRes.ok) { ((await colsRes.json()).value || []).forEach(c => { colMap[c.displayName] = c.name; }); }

          // Read the entire Reports to be Billed list. A single Graph page can
          // contain only the first 500 items, which caused older samples to be missed.
          let allBilledItems = [];
          let billedNext = `${GRAPH2}/sites/${siteId}/lists/${rtbListId}/items?$expand=fields($select=Title)&$top=500`;
          while (billedNext) {
            const billedRes = await fetch(billedNext, { headers: authHdr2 });
            if (!billedRes.ok) {
              throw new Error(`RTB read failed ${billedRes.status}: ${(await billedRes.text()).slice(0,120)}`);
            }
            const billedData = await billedRes.json();
            allBilledItems.push(...(billedData.value || []));
            billedNext = billedData['@odata.nextLink'] || null;
          }

          const billedItems = allBilledItems
            .filter(i => (i.fields?.Title || '').split(' ')[0].trim() === baseId);
          let billedOk = 0;
          for (const item of billedItems) {
            if (isDuplicate) {
              const dRes = await fetch(
                `${GRAPH2}/sites/${siteId}/lists/${rtbListId}/items/${item.id}`,
                { method:'DELETE', headers:authHdr2 }
              );
              if (dRes.ok || dRes.status === 204 || dRes.status === 404) billedOk++;
              else { const t = await dRes.text(); log.push(`⚠️ RTB DELETE failed ${dRes.status}: ${t.slice(0,100)}`); }
              continue;
            }

            const pFields = {};
            pFields[colMap['Lab ID'] || 'Title'] = rejLabId;
            pFields[colMap['Item/Service']  || 'Item_x002F_Service'] = rejectionType;
            pFields[colMap['Test Type SKU'] || 'Test_x0020_Type_x0020_SKU'] = 'REJ';
            const rejectionReportDate = nextBusinessDayFromLabId(baseId);
            if (rejectionReportDate) pFields[colMap['Report Date'] || 'Report_x0020_Date'] = rejectionReportDate;
            const pRes = await fetch(
              `${GRAPH2}/sites/${siteId}/lists/${rtbListId}/items/${item.id}/fields`,
              { method:'PATCH', headers:{ ...authHdr2, 'Content-Type':'application/json' }, body:JSON.stringify(pFields) }
            );
            if (pRes.ok) billedOk++;
            else { const t = await pRes.text(); log.push(`⚠️ RTB PATCH failed ${pRes.status}: ${t.slice(0,100)}`); }
          }
          if (billedOk > 0) log.push(isDuplicate
            ? `✅ Reports to be Billed deleted (${billedOk} row(s))`
            : `✅ Reports to be Billed updated (${billedOk} row(s))`);
        }
      } catch(e) { log.push(`⚠️ Reports to be Billed: ${e.message}`); }

      // 2. Update Archived Intake. Duplicate samples remain as the permanent
      // historical record but use "[Lab ID] Dup" + service "Duplicate".
      const archived = await listItems(LISTS.ARCHIVED_INTAKE, { top: 2000 });
      const matches  = archived.filter(r => (r.field_1 || '').startsWith(baseId));

      for (const item of matches) {
        const existingNotes = item.field_13 || '';
        const newNotes = existingNotes ? `${existingNotes} | ${rejNote}` : rejNote;
        await updateItem(LISTS.ARCHIVED_INTAKE, item._id, {
          field_1:  rejLabId,
          field_2:  archivedService,
          field_13: newNotes,
          field_14: 'Rejected',
        }).catch(e => context.log('[ArchivedIntake] Error:', e.message));
      }
      log.push(`✅ Archived Intake: updated ${matches.length} row(s) → ${rejLabId} / ${archivedService} / Rejected`);

      // ── COA / Form Responses ─────────────────────────────────────────────
      if (isDuplicate) {
        try {
          const coaRows = await getCoaRowsByBaseIds([baseId]);
          const cleared = await clearCoaRows(coaRows);
          log.push(cleared
            ? `✅ COA row(s) deleted/cleared (${cleared})`
            : 'ℹ️ COA: no matching rows found');
        } catch(e) {
          log.push(`⚠️ COA delete: ${e.message}`);
        }
      } else {
        try {
          const refreshed = await listItems(LISTS.ARCHIVED_INTAKE, { top: 2000 });
          const rejectedRows = refreshed.filter(r => (r.field_1 || '').split(' ')[0].trim() === baseId);
          const coaSync = await syncCoaFromArchivedRows(baseId, rejectedRows, context);
          log.push(`✅ COA sheet synchronized to rejection (${coaSync.updated} row(s))`);
        } catch(e) {
          log.push(`⚠️ COA sheet: ${e.message}`);
        }
      }

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, labId, newLabId: rejLabId, rejectionType, duplicateOf: cleanDuplicateOf, log }),
      };
    } catch(e) {
      context.log('[reject-sample] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
