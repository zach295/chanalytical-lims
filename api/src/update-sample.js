/**
 * update-sample.js — Azure version (v361)
 * Updates sample information in SharePoint Archived Intake and Accession Log.
 * v361 adds: test type update with suffix recalculation.
 */
const { app }   = require('@azure/functions');
const { listItems, updateItem, LISTS } = require('../shared/graph');

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

app.http('update-sample', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { baseId, updates, updatedBy } = await request.json();
      if (!baseId) return { status: 400, jsonBody: { error: 'baseId required' } };

      const log = [];
      let rowsUpdated = 0;

      // ── Find all Archived Intake items for this baseId ────────────────────────
      const archivedItems = await listItems(LISTS.ARCHIVED_INTAKE, {
        filter: `startswith(fields/FullId,'${baseId}')`,
        top: 20,
      }).catch(() => []);

      // ── Update standard fields ────────────────────────────────────────────────
      for (const item of archivedItems) {
        const fields = {};
        if (updates.customer     !== undefined) fields.ClientName   = updates.customer;
        if (updates.dateDrawn    !== undefined) fields.DateDrawn    = updates.dateDrawn;
        if (updates.timeDrawn    !== undefined) fields.TimeDrawn    = to24h(updates.timeDrawn);
        if (updates.receivedDate !== undefined) fields.ReceivedDate = updates.receivedDate;
        if (updates.receivedTime !== undefined) fields.ReceivedTime = updates.receivedTime;
        if (updates.location     !== undefined) fields.Address      = updates.location;
        if (updates.city         !== undefined) fields.City         = updates.city;
        if (updates.state        !== undefined) fields.State        = updates.state;
        if (updates.zip          !== undefined) fields.Zip          = updates.zip;
        if (updates.notes        !== undefined) fields.Notes        = updates.notes;

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
            const currentFullId = item.FullId || '';
            const rowBase = currentFullId.match(/(\d{6}-\d{3})/)?.[1] || '';
            if (rowBase !== baseId) continue;
            const newFullId = `${rowBase} ${newSuffix}`;
            await updateItem(LISTS.ARCHIVED_INTAKE, item._id, {
              FullId:  newFullId,
              CoaTest: newTest,
              Title:   newFullId,
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
