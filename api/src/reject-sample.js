const { app } = require('@azure/functions');
const { createItem, listItems, updateItem, LISTS } = require('../shared/graph');

// ── Control Sheet Helper ──────────────────────────────────────────────────────
// Finds C_MMDDYY.xlsx in Test C folder and updates the lab ID cell in column A
async function updateControlSheet(siteId, datePrefix, baseId, newLabId, token, context) {
  const GRAPH        = 'https://graph.microsoft.com/v1.0';
  const controlFolder = process.env.SP_CONTROL_FOLDER ||
    '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
  const marker    = 'Shared Documents/';
  const idx       = controlFolder.indexOf(marker);
  const relPath   = idx >= 0 ? controlFolder.slice(idx + marker.length) : controlFolder.replace(/^\/+/, '');
  const fileName  = `C_${datePrefix}.xlsx`;
  const filePath  = `${relPath}/${fileName}`.split('/').map(encodeURIComponent).join('/');
  const authHdr   = { Authorization: `Bearer ${token}` };

  // 1. Get file
  const fileRes = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${filePath}`, { headers: authHdr });
  if (!fileRes.ok) throw new Error(`Control sheet C_${datePrefix}.xlsx not found (${fileRes.status})`);
  const { id: fileId } = await fileRes.json();

  // 2. Open session
  const sesRes  = await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook/createSession`,
    { method: 'POST', headers: { ...authHdr, 'Content-Type': 'application/json' },
      body: JSON.stringify({ persistChanges: true }) }
  );
  const { id: sid } = await sesRes.json();
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
      `${wbBase}/worksheets/${wsId}/usedRange?$select=values`,
      { headers: wbHdr }
    );
    const rows = (await rangeRes.json()).values || [];

    let targetRow = -1;
    for (let i = 0; i < rows.length; i++) {
      const cell     = String(rows[i][0] || '').trim();
      const cellBase = cell.split(' ')[0].trim();
      if (cellBase === baseId || cell === baseId) { targetRow = i + 1; break; }
    }

    if (targetRow < 0) throw new Error(`Lab ID ${baseId} not found in column A of C_${datePrefix}.xlsx`);

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
    const rangeRes  = await fetch(`${wbBase}/worksheets/${wsId}/usedRange?$select=values`, { headers: wbHdr });
    const rows      = (await rangeRes.json()).values || [];
    let   targetRow = -1;
    for (let i = 0; i < rows.length; i++) {
      const cell = String(rows[i][0] || '').trim();
      if (cell.split(' ')[0] === baseId || cell === baseId) { targetRow = i + 1; break; }
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

app.http('reject-sample', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { labId, rejectionType, reason, rejectedBy } = await request.json();
      if (!labId)          return { status: 400, body: JSON.stringify({ error: 'labId required' }) };
      if (!rejectionType)  return { status: 400, body: JSON.stringify({ error: 'rejectionType required' }) };
      if (!reason?.trim()) return { status: 400, body: JSON.stringify({ error: 'reason required' }) };

      const siteId = process.env.SP_SITE_ID;
      const now    = new Date().toISOString();
      const baseId = labId.split(' ')[0].trim();
      const rejNote = `${rejectionType}: ${reason.trim()}`;
      const log = [];

      // 1. Write to Rejected list
      // field_1=LabId, field_2=RejectionType, field_3=Reason, field_4=RejectedBy
      await createItem(LISTS.REJECTED, {
        Title:   now,
        field_1: labId,
        field_2: rejectionType,
        field_3: reason.trim(),
        field_4: rejectedBy || 'Lab Staff',
      }).catch(e => context.log('[Rejected] Error:', e.message));
      log.push('✅ Written to Rejected list');

      // ── Update Control Sheet — change suffix to REJ ───────────────────────
      const datePrefix = baseId.slice(0, 6); // MMDDYY from lab ID prefix
      const rejLabId   = `${baseId} REJ`;
      try {
        const csResult = await updateControlSheet(siteId, datePrefix, baseId, rejLabId, token, context);
        log.push(csResult.updated
          ? `✅ Control sheet updated to REJ (row ${csResult.row})`
          : `ℹ️ Control sheet: ${csResult.reason || 'row not found'}`);
      } catch(e) { log.push(`⚠️ Control sheet: ${e.message}`); }

      // ── Update Radon Sheet if applicable ─────────────────────────────────
      const archivedAll = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 }).catch(() => []);
      const isRadonSample = archivedAll.some(r =>
        r.field_1?.startsWith(baseId) && (r.field_2||'').toLowerCase().includes('radon')
      );
      if (isRadonSample) {
        try {
          const rwResult = await updateRadonSheet(siteId, datePrefix, baseId, rejLabId, token, context);
          log.push(rwResult.updated ? `✅ Radon sheet updated to REJ` : `ℹ️ Radon: ${rwResult.reason}`);
        } catch(e) { log.push(`⚠️ Radon sheet: ${e.message}`); }
      }

      // ── Write to Activity Log ─────────────────────────────────────────────
      const actNow = new Date();
      const dateStr = actNow.toLocaleDateString('en-US', { timeZone: 'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
      const timeStr = actNow.toLocaleTimeString('en-US', { timeZone: 'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
      try {
        // Field names from SP Activity Log list (check exact internal names)
        await createItem('Activity Log', {
          Title:        `${dateStr} ${labId}`,
          Client:       labId,
          ActivityType: rejectionType,
          Notes:        reason.trim(),
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

      // 2. Update Archived Intake — change suffix to REJ and append notes
      // field_1=fullId, field_2=coaTest, field_13=notes, field_14=status
      const archived = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
      const matches  = archived.filter(r => (r.field_1 || '').startsWith(baseId));

      for (const item of matches) {
        const existingNotes = item.field_13 || '';
        const newNotes = existingNotes ? `${existingNotes} | ${rejNote}` : rejNote;
        // Change full lab ID suffix to REJ (e.g. "072826-003 COMP" → "072826-003 REJ")
        const newFullId = `${baseId} REJ`;
        await updateItem(LISTS.ARCHIVED_INTAKE, item._id, {
          field_1:  newFullId,
          field_2:  rejectionType,
          field_13: newNotes,
          field_14: 'Rejected',
        }).catch(e => context.log('[ArchivedIntake] Error:', e.message));
      }
      log.push(`✅ Archived Intake: updated ${matches.length} row(s) → suffix changed to REJ`);

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, labId, newLabId: `${baseId} REJ`, rejectionType, log }),
      };
    } catch(e) {
      context.log('[reject-sample] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
