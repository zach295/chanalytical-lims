/**
 * update-sample.js — Azure version (v361)
 * Updates sample information in SharePoint Archived Intake and Accession Log.
 * v361 adds: test type update with suffix recalculation.
 */
const { app }   = require('@azure/functions');
const { listItems, updateItem, createItem, getToken, LISTS } = require('../shared/graph');

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
async function updateControlSheet(siteIdArg, datePrefix, baseId, newLabId, tokenArg, context) {
  const GRAPH        = 'https://graph.microsoft.com/v1.0';
  const siteId       = siteIdArg || process.env.SP_SITE_ID;
  const token        = tokenArg  || (await getToken());
  const controlFolder = process.env.SP_CONTROL_FOLDER ||
    '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
  const marker    = 'Shared Documents/';
  const idx       = controlFolder.indexOf(marker);
  const relPath   = idx >= 0 ? controlFolder.slice(idx + marker.length) : controlFolder.replace(/^\/+/, '');
  const authHdr   = { Authorization: `Bearer ${token}` };
  const MONTHS    = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];

  // Build month subfolder: "August 2026" from MMDDYY prefix
  const mm        = parseInt(datePrefix.slice(0, 2)) - 1;
  const yy        = datePrefix.slice(4, 6);
  const year      = '20' + yy;
  const monthName = MONTHS[mm] || datePrefix.slice(0, 2);
  const fileName  = `C_${datePrefix}.xlsx`;

  // Try month subfolder first, then flat folder
  let fileId = null;
  for (const tryPath of [`${relPath}/${monthName} ${year}/${fileName}`, `${relPath}/${fileName}`]) {
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
  const sid    = sesData.id;
  const wbHdr  = { ...authHdr, 'workbook-session-id': sid, 'Content-Type': 'application/json' };
  const wbBase = `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook`;

  try {
    // 3. Get first worksheet
    const sheetsRes = await fetch(`${wbBase}/worksheets`, { headers: wbHdr });
    const sheetsData = await sheetsRes.json();
    const sheets    = sheetsData.value || [];
    if (!sheets.length) throw new Error('No worksheets in control sheet');
    const wsId = sheets[0].id;

    // 4. Read used range to find matching row
    const rangeRes  = await fetch(
      `${wbBase}/worksheets/${wsId}/usedRange?$select=values`,
      { headers: wbHdr }
    );
    if (!rangeRes.ok) {
      const errTxt = await rangeRes.text();
      throw new Error(`usedRange failed (${rangeRes.status}): ${errTxt.slice(0, 150)}`);
    }
    const rangeData = await rangeRes.json();
    if (rangeData.error) throw new Error(`usedRange error: ${JSON.stringify(rangeData.error).slice(0, 150)}`);
    const rows = rangeData.values || [];
    if (context) context.log(`[controlSheet] Sheet="${sheets[0].name}" rows=${rows.length} fileId=${fileId}`);

    let targetRow = -1;
    for (let i = 0; i < rows.length; i++) {
      const cell     = String(rows[i][0] || '').trim().replace(/\s+/g, ' ');
      const cellBase = cell.split(' ')[0].replace(/[^\w-]/g, '').trim();
      const baseClean = baseId.replace(/[^\w-]/g, '').trim();
      if (cellBase === baseClean || cell.startsWith(baseId)) { targetRow = i + 1; break; }
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
async function updateRadonSheet(siteIdArg, datePrefix, baseId, newLabId, tokenArg, context) {
  const GRAPH      = 'https://graph.microsoft.com/v1.0';
  const siteId     = siteIdArg || process.env.SP_SITE_ID;
  const token      = tokenArg  || (await getToken());
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

// ── Reports to be Billed helpers ─────────────────────────────────────────────
let _rtbListId = null;
async function getRTBListId(siteId, token, GRAPH, authHdr) {
  if (_rtbListId) return _rtbListId;
  const r = await fetch(`${GRAPH}/sites/${siteId}/lists?$select=id,displayName`, { headers: authHdr });
  _rtbListId = ((await r.json()).value || []).find(l => l.displayName === 'Reports to be Billed')?.id || null;
  return _rtbListId;
}
async function getRTBColMap(siteId, listId, token, GRAPH, authHdr) {
  const r = await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/columns?$select=name,displayName&$top=100`, { headers: authHdr });
  const colMap = {};
  if (r.ok) { ((await r.json()).value || []).forEach(c => { colMap[c.displayName] = c.name; }); }
  return colMap;
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
      // Fetch Archived Intake items directly via Graph API
      const GRAPH   = 'https://graph.microsoft.com/v1.0';
      const token   = await getToken();
      const siteId  = process.env.SP_SITE_ID;
      const authHdr = { Authorization: `Bearer ${token}` };

      // Find Archived Intake list ID by name
      let listId = null;
      try {
        const listsRes = await fetch(`${GRAPH}/sites/${siteId}/lists?$select=id,displayName&$top=50`, { headers: authHdr });
        const listsData = await listsRes.json();
        const found = (listsData.value || []).find(l => l.displayName === 'Archived Intake');
        listId = found?.id;
        context.log(`[update-sample] Archived Intake list id=${listId}`);
      } catch(e) { log.push('⚠️ Could not find Archived Intake list: ' + e.message); }

      let archivedItems = [];
      if (listId) {
        try {
          const itemsRes  = await fetch(
            `${GRAPH}/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`,
            { headers: authHdr }
          );
          const itemsData = await itemsRes.json();
          const allItems  = (itemsData.value || []).map(i => ({ ...i.fields, _id: i.id }));
          archivedItems   = allItems.filter(r => (r.field_1 || '').split(' ')[0].trim() === baseId);
          context.log(`[update-sample] total=${allItems.length} matched=${archivedItems.length} baseId=${baseId}`);
        } catch(e) { log.push('⚠️ Archived Intake read error: ' + e.message); }
      }

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
          if (listId) {
            await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items/${item._id}/fields`,
              { method: 'PATCH', headers: { ...authHdr, 'Content-Type': 'application/json' },
                body: JSON.stringify(fields) });
            rowsUpdated++;
          }
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
            if (listId) await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items/${item._id}/fields`,
              { method: 'PATCH', headers: { ...authHdr, 'Content-Type': 'application/json' },
                body: JSON.stringify({ field_1: newFullId, field_2: newTest }) });
            rowsUpdated++;
          }
          log.push(`Test type updated to ${newTest} (${newSuffix})`);
          log.push(`Note: update control sheet lab ID manually if suffix changed`);

          // ── If previously rejected — restore status to Approved ─────────────
          const prevTest = archivedItems[0]?.field_2 || '';
          const wasRejected = /^rejected|rej/i.test(prevTest) || /\bREJ\b/i.test(archivedItems[0]?.field_1 || '');
          if (wasRejected) {
            for (const item of archivedItems) {
              if (listId) await fetch(
                `${GRAPH}/sites/${siteId}/lists/${listId}/items/${item._id}/fields`,
                { method: 'PATCH', headers: { ...authHdr, 'Content-Type': 'application/json' },
                  body: JSON.stringify({ field_14: 'Pending' }) }
              );
            }
            log.push(`✅ Status restored to Pending (was Rejected)`);
          }

          // ── Update Accession Log test type ──────────────────────────────────
          try {
            const accListRes  = await fetch(`${GRAPH}/sites/${siteId}/lists?$select=id,displayName`, { headers: authHdr });
            const accListId   = ((await accListRes.json()).value || []).find(l => l.displayName === 'Accession Log')?.id;
            if (accListId) {
              const accItemsRes = await fetch(
                `${GRAPH}/sites/${siteId}/lists/${accListId}/items?$expand=fields($select=field_1,field_2,field_3,field_4)&$top=500`,
                { headers: authHdr }
              );
              const accItems = ((await accItemsRes.json()).value || [])
                .filter(i => String(i.fields?.field_1 || '').split(' ')[0].trim() === baseId);
              const newFullId = `${baseId} ${newSuffix}`;
              for (const item of accItems) {
                await fetch(
                  `${GRAPH}/sites/${siteId}/lists/${accListId}/items/${item.id}/fields`,
                  { method: 'PATCH', headers: { ...authHdr, 'Content-Type': 'application/json' },
                    body: JSON.stringify({ field_2: newFullId, field_3: newTest, field_4: newSuffix }) }
                );
              }
              if (accItems.length) log.push(`✅ Accession Log updated (${accItems.length} row(s))`);
              else log.push(`ℹ️ Accession Log: no rows found for ${baseId}`);
            }
          } catch(e) { log.push(`⚠️ Accession Log: ${e.message}`); }

          // ── Update Reports to be Billed ─────────────────────────────────────
          try {
            const rtbListId  = await getRTBListId(siteId, token, GRAPH, authHdr);
            if (rtbListId) {
              const colMap     = await getRTBColMap(siteId, rtbListId, token, GRAPH, authHdr);
              const billedRes  = await fetch(
                `${GRAPH}/sites/${siteId}/lists/${rtbListId}/items?$expand=fields($select=Title)&$top=500`,
                { headers: authHdr }
              );
              const billedItems = ((await billedRes.json()).value || [])
                .filter(i => (i.fields?.Title || '').split(' ')[0].trim() === baseId);
              let billedOk = 0;
              for (const item of billedItems) {
                const pFields = {};
                pFields[colMap['Item/Service']   || 'Item_x002F_Service']  = newTest;
                pFields[colMap['Test Type SKU']  || 'Test_x0020_Type_x0020_SKU'] = newSuffix;
                const pRes = await fetch(
                  `${GRAPH}/sites/${siteId}/lists/${rtbListId}/items/${item.id}/fields`,
                  { method: 'PATCH', headers: { ...authHdr, 'Content-Type': 'application/json' },
                    body: JSON.stringify(pFields) }
                );
                if (pRes.ok) billedOk++;
                else { const t = await pRes.text(); log.push(`⚠️ RTB PATCH failed ${pRes.status}: ${t.slice(0,100)}`); }
              }
              if (billedOk > 0) log.push(`✅ Reports to be Billed updated (${billedOk} row(s))`);
            }
          } catch(e) { log.push(`⚠️ Reports to be Billed: ${e.message}`); }
        } catch(testErr) {
          log.push(`Test update failed: ${testErr.message}`);
        }
      }

      // ── Update Reports to be Billed for any field change ──────────────────────
      try {
        const rtbListId2 = await getRTBListId(siteId, token, GRAPH, authHdr);
        if (rtbListId2) {
          const colMap2 = await getRTBColMap(siteId, rtbListId2, token, GRAPH, authHdr);
          const bRes    = await fetch(
            `${GRAPH}/sites/${siteId}/lists/${rtbListId2}/items?$expand=fields($select=Title)&$top=500`,
            { headers: authHdr }
          );
          const bItems  = ((await bRes.json()).value || [])
            .filter(i => (i.fields?.Title || '').split(' ')[0].trim() === baseId);
          for (const item of bItems) {
            const bFields = {};
            if (updates.customer)     bFields[colMap2['Customer']       || 'Customer']       = updates.customer;
            if (updates.location)     bFields[colMap2['Location']       || 'Location']       = updates.location;
            if (updates.city)         bFields[colMap2['City/Town']      || 'City_x002F_Town']= updates.city;
            if (updates.state)        bFields[colMap2['State']          || 'State']          = updates.state;
            if (updates.zip)          bFields[colMap2['Zip']            || 'Zip']            = updates.zip;
            if (updates.dateDrawn)    bFields[colMap2['Date Drawn']     || 'Date_x0020_Drawn']           = updates.dateDrawn;
            if (updates.timeDrawn)    bFields[colMap2['Time Drawn']     || 'Time_x0020_Drawn']           = to24h(updates.timeDrawn);
            if (updates.receivedDate) bFields[colMap2["Date Rec'd"]     || 'Date_x0020_Rec_x0027_d']    = updates.receivedDate;
            if (updates.receivedTime) bFields[colMap2["Time Rec'd"]     || 'Time_x0020_Rec_x0027_d']    = to24h(updates.receivedTime);
            if (Object.keys(bFields).length) {
              const pRes2 = await fetch(
                `${GRAPH}/sites/${siteId}/lists/${rtbListId2}/items/${item.id}/fields`,
                { method: 'PATCH', headers: { ...authHdr, 'Content-Type': 'application/json' },
                  body: JSON.stringify(bFields) }
              );
              if (!pRes2.ok) { const t = await pRes2.text(); log.push(`⚠️ RTB fields PATCH failed ${pRes2.status}: ${t.slice(0,100)}`); }
            }
          }
          const sentFields = Object.keys((() => {
            const bf = {};
            if (updates.customer)     bf.x = 1;
            if (updates.location)     bf.x = 1;
            if (updates.city)         bf.x = 1;
            if (updates.state)        bf.x = 1;
            if (updates.zip)          bf.x = 1;
            if (updates.dateDrawn)    bf.x = 1;
            if (updates.timeDrawn)    bf.x = 1;
            if (updates.receivedDate) bf.x = 1;
            if (updates.receivedTime) bf.x = 1;
            return bf;
          })()).length;
          if (bItems.length && sentFields > 0) log.push(`✅ Reports to be Billed fields updated (${bItems.length} row(s))`);
        }
      } catch(e) { log.push(`⚠️ Reports to be Billed fields: ${e.message}`); }

      // ── Update customer name in Accession Log ─────────────────────────────────
      if (updates.customer !== undefined) {
        const accItems = await listItems(LISTS.ACCESSION_LOG, {
          filter: `startswith(fields/field_1,'${baseId}')`,
          top: 20,
        }).catch(() => []);

        for (const item of accItems) {
          await updateItem(LISTS.ACCESSION_LOG, item._id, { Customer: updates.customer });
          rowsUpdated++;
        }
        log.push(`Accession Log: ${accItems.length} row(s) updated`);
      }

      context.log(`[update-sample] ${baseId} by ${updatedBy||'staff'}: ${log.join(', ')}`);

      // ── Remove from Rejected list if correcting away from a rejected type ───────
      if (updates.coaTest) {
        const wasRejected = /^rejected|^wq.*reject|^rw.*reject/i.test(updates.coaTest) === false;
        const prevTest    = archivedItems[0]?.field_2 || '';
        const wasRejectedBefore = /^rejected|^wq.*reject|^rw.*reject/i.test(prevTest);
        if (wasRejectedBefore && !updates.coaTest.match(/^rejected|^wq.*reject|^rw.*reject/i)) {
          try {
            // Find Rejected list ID
            const rejListRes = await fetch(`${GRAPH}/sites/${siteId}/lists?$select=id,displayName`, {headers:authHdr});
            const rejListId  = ((await rejListRes.json()).value||[]).find(l=>l.displayName==='Rejected')?.id;
            if (rejListId) {
              const rejItemsRes = await fetch(
                `${GRAPH}/sites/${siteId}/lists/${rejListId}/items?$expand=fields($select=field_1)&$top=200`,
                {headers:authHdr}
              );
              const rejItems = ((await rejItemsRes.json()).value||[])
                .filter(i => (i.fields?.field_1||'').split(' ')[0].trim() === baseId);
              for (const r of rejItems) {
                await fetch(`${GRAPH}/sites/${siteId}/lists/${rejListId}/items/${r.id}`,
                  { method:'DELETE', headers:authHdr });
              }
              if (rejItems.length) log.push(`✅ Removed ${rejItems.length} row(s) from Rejected list`);
            }
          } catch(e) { log.push(`⚠️ Rejected list cleanup: ${e.message}`); }
        }
      }

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

        // Radon sheet — if new OR old test type includes radon
        const isRadon = (updates.coaTest || '').toLowerCase().includes('radon') ||
          (updates.coaTest || '').toLowerCase().includes('rw') ||
          archivedItems.some(r => (r.field_2||'').toLowerCase().includes('radon') || (r.field_2||'').toLowerCase().includes('rw'));
        if (isRadon) {
          try {
            const rwResult = await updateRadonSheet(siteId, datePrefix, baseId, finalLabId, token, context);
            log.push(rwResult.updated ? `✅ Radon sheet updated` : `ℹ️ Radon: ${rwResult.reason}`);
          } catch(e) { log.push(`⚠️ Radon sheet: ${e.message}`); }
        }
      }

      // Results Cache: lab ID stores base ID only — no suffix — no update needed on correction

      // ── Write to Activity Log ───────────────────────────────────────────────
      try {
        const actNow  = new Date();
        const logDate = actNow.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
        const logTime = actNow.toLocaleTimeString('en-US', { timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
        const fieldLabels = { coaTest:'Test Type', customer:'Customer', dateDrawn:'Date Drawn',
          timeDrawn:'Time Drawn', receivedDate:'Date Received', receivedTime:'Time Received',
          location:'Address', city:'City', state:'State', zip:'Zip', notes:'Notes' };
        const changes = Object.entries(updates)
          .filter(([,v]) => v !== undefined && v !== '')
          .map(([k,v]) => `${fieldLabels[k] || k} → ${v}`)
          .join('; ');
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
