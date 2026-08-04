/**
 * update-sample.js — Azure version (v361)
 * Updates sample information in SharePoint Archived Intake and Accession Log.
 * v361 adds: test type update with suffix recalculation.
 */
const { app }   = require('@azure/functions');
const { listItems, updateItem, createItem, LISTS } = require('../shared/graph');

const SUFFIX_MAP = {
  'Basic Safety (FHA)':'BS','Basic Safety':'BS','Standard Safety':'SS',
  'Expanded Safety (Mortgage Test)':'EXP','WW - Expanded Safety':'WW',
  'Comprehensive':'COMP','Pro Plus':'PP','Radon Water':'RW',
  'AIO FHA':'AIOFHA','AIO Portability':'PORT',
  'Alkalinity':'ALK','Arsenic, Total':'AS','Arsenic, Speciation':'AS-SPEC',
  'Bacteria':'BAC','Cadmium, Total':'CD','Calcium, Total':'CA',
  'Chloride, Total':'CL','Copper, Total':'CU','Fluoride':'FL',
  'Iron, Total':'FE','Lead, Total':'PB','Magnesium, Total':'MG',
  'Manganese, Total':'MN','Nitrate':'NO3','Nitrite':'NO2',
  'pH':'PH','Sodium, Total':'NA','Sulfate':'SO4','Tannins':'TAN',
  'Total Dissolved Solids (TDS)':'TDS','Total Hardness':'HRD','Uranium, Total':'U',
  'Rejected - Timeout':'REJ','Rejected - Missing Information':'REJ',
  'Rejected - Chlorine':'REJ','Rejected - Other':'REJ',
};

function getSuffix(testName) {
  if (!testName) return '';
  const parts = testName.split(' | ').map(t => t.trim()).filter(Boolean);
  const suffixes = parts.map(t => SUFFIX_MAP[t] || t.substring(0,3).toUpperCase());
  return suffixes.join(', ');
}

function to24h(t) {
  if (!t) return '';
  const s = String(t).trim().replace(/^[^\d]*/, '');
  const extracted = s.match(/(\d{1,2}:\d{2}(?:\s*(?:AM|PM))?)$/i)?.[1] || s;
  const plain = extracted.match(/^(\d{1,2}):(\d{2})$/);
  if (plain) {
    const h = parseInt(plain[1]), m = parseInt(plain[2]);
    if (h >= 0 && h <= 23 && m >= 0 && m <= 59)
      return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
  }
  const ampm = extracted.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (ampm) {
    let h = parseInt(ampm[1]); const m = parseInt(ampm[2]);
    const isPM = ampm[3].toUpperCase() === 'PM';
    if (isPM && h < 12) h += 12;
    if (!isPM && h === 12) h = 0;
    if (h >= 0 && h <= 23 && m >= 0 && m <= 59)
      return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
  }
  return extracted;
}

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

app.http('update-sample', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { baseId, updates, updatedBy } = await request.json();
      if (!baseId) return { status: 400, jsonBody: { error: 'baseId required' } };

      const log = [];
      let rowsUpdated = 0;

      // ── Find all Archived Intake items for this baseId ─────────────────────────
      // Filter client-side — OData filter unreliable for field_N columns
      const allIntake = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 }).catch(() => []);
      const archivedItems = allIntake.filter(r => {
        const rowBase = (r.field_1 || '').split(' ')[0].trim();
        return rowBase === baseId;
      });

      // ── Update standard fields (Archived Intake uses field_N internal names) ────
      for (const item of archivedItems) {
        const fields = {};
        // field_2=coaTest, field_3=customer, field_4=dateDrawn, field_5=timeDrawn
        // field_6=dateReceived, field_7=timeReceived, field_8=address, field_9=city
        // field_10=state, field_11=zip, field_13=notes
        if (updates.customer     !== undefined) fields.field_3  = updates.customer;
        if (updates.dateDrawn    !== undefined) fields.field_4  = updates.dateDrawn;
        if (updates.timeDrawn    !== undefined) fields.field_5  = to24h(updates.timeDrawn);
        if (updates.receivedDate !== undefined) fields.field_6  = updates.receivedDate;
        if (updates.receivedTime !== undefined) fields.field_7  = to24h(updates.receivedTime);
        if (updates.location     !== undefined) fields.field_8  = updates.location;
        if (updates.city         !== undefined) fields.field_9  = updates.city;
        if (updates.state        !== undefined) fields.field_10 = updates.state;
        if (updates.zip          !== undefined) fields.field_11 = updates.zip;
        if (updates.notes        !== undefined) fields.field_13 = updates.notes;

        if (Object.keys(fields).length > 0) {
          await updateItem(LISTS.ARCHIVED_INTAKE, item._id, fields);
          rowsUpdated++;
        }
      }
      log.push(`Archived Intake: ${archivedItems.length} row(s) found`);

      // ── Update test type ──────────────────────────────────────────────────────
      if (updates.coaTest) {
        const newTest   = updates.coaTest.trim();
        const newSuffix = getSuffix(newTest);

        try {
          for (const item of archivedItems) {
            const currentFullId = item.field_1 || '';
            const rowBase = currentFullId.split(' ')[0].trim();
            if (rowBase !== baseId) continue;
            const newFullId = `${rowBase} ${newSuffix}`;
            await updateItem(LISTS.ARCHIVED_INTAKE, item._id, {
              field_1: newFullId,
              field_2: newTest,
              Title:   new Date().toISOString(), // Title is timestamp in Archived Intake
            });
            rowsUpdated++;
          }
          log.push(`Test type updated to ${newTest} (${newSuffix})`);
          log.push(`Note: update control sheet lab ID manually if suffix changed`);
        } catch(testErr) {
          log.push(`Test update failed: ${testErr.message}`);
        }
      }

      // ── Update customer name in Accession Log ─────────────────────────────────
      if (updates.customer !== undefined) {
        const accItems = await listItems(LISTS.ACCESSION_LOG, {
          filter: `startswith(fields/BaseId,'${baseId}')`,
          top: 20,
        }).catch(() => []);

        for (const item of accItems) {
          await updateItem(LISTS.ACCESSION_LOG, item._id, { Customer: updates.customer });
          rowsUpdated++;
        }
        log.push(`Accession Log: ${accItems.length} row(s) updated`);
      }

      context.log(`[update-sample] ${baseId} by ${updatedBy||'staff'}: ${log.join(', ')}`);

      // ── Update Control Sheet ─────────────────────────────────────────────────
      const datePrefix = baseId.slice(0, 6); // MMDDYY
      if (updates.coaTest || rowsUpdated > 0) {
        const finalLabId = updates.coaTest
          ? `${baseId} ${getSuffix(updates.coaTest.trim())}`
          : (archivedItems[0]?.field_1 || baseId);
        try {
          const csResult = await updateControlSheet(siteId, datePrefix, baseId, finalLabId, token, context);
          log.push(csResult.updated
            ? `✅ Control sheet updated (row ${csResult.row})`
            : `ℹ️ Control sheet: ${csResult.reason || 'no change needed'}`);
        } catch(e) { log.push(`⚠️ Control sheet: ${e.message}`); }

        // Radon sheet — only if sample has radon component
        const isRadon = (updates.coaTest || '').toLowerCase().includes('radon') ||
          archivedItems.some(r => (r.field_2||'').toLowerCase().includes('radon'));
        if (isRadon) {
          try {
            const rwResult = await updateRadonSheet(siteId, datePrefix, baseId, finalLabId, token, context);
            log.push(rwResult.updated ? `✅ Radon sheet updated` : `ℹ️ Radon: ${rwResult.reason}`);
          } catch(e) { log.push(`⚠️ Radon sheet: ${e.message}`); }
        }
      }

      // ── Write to Activity Log ───────────────────────────────────────────────
      try {
        const actNow  = new Date();
        const logDate = actNow.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
        const logTime = actNow.toLocaleTimeString('en-US', { timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
        const changes = Object.entries(updates).filter(([,v])=>v!==undefined&&v!=='').map(([k,v])=>`${k}: ${v}`).join('; ');
        await createItem('Activity Log', {
          Title:        `${logDate} ${baseId}`,
          Client:       baseId,
          ActivityType: 'Sample Correction',
          Notes:        changes,
          By:           updatedBy || 'Lab Staff',
          LogDate:      logDate,
          LogTime:      logTime,
          Quantity:     0,
        });
        log.push('✅ Written to Activity Log');
      } catch(e) {
        log.push('⚠️ Activity Log: ' + e.message);
      }

      return {
        status: 200,
        jsonBody: { success: true, baseId, rowsUpdated, log },
      };

    } catch (e) {
      context.log('[update-sample] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  },
});
