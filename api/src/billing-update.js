/**
 * billing-update.js — updates QB/paid status on Reports to be Billed list
 */
const { app } = require('@azure/functions');
const { getToken } = require('../shared/graph');
const { writeActivityLog } = require('../shared/audit');

const GRAPH = 'https://graph.microsoft.com/v1.0';

function reportDateToISO(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  let m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) return `${m[3]}-${String(m[1]).padStart(2,'0')}-${String(m[2]).padStart(2,'0')}`;
  m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2})$/);
  if (m) return `20${m[3]}-${String(m[1]).padStart(2,'0')}-${String(m[2]).padStart(2,'0')}`;
  return '';
}

app.http('billing-update', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body = await request.json().catch(() => ({}));
      const { itemId, fields } = body;

      const token = await getToken();
      const siteId = process.env.SP_SITE_ID;

      const res = await fetch(
        `${GRAPH}/sites/${siteId}/lists?$select=id,displayName`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const listId = ((await res.json()).value || [])
        .find(l => l.displayName === 'Reports to be Billed')?.id;
      if (!listId) return { status: 404, jsonBody: { error: 'List not found' } };

      // Admin bulk status update. Date range is based only on Report Date.
      // "paid" intentionally changes ONLY the Pd checkbox; amount/date fields are untouched.
      if (body.bulkAction) {
        const action = String(body.bulkAction || '').toLowerCase();
        const from = String(body.fromReportDate || '');
        const to = String(body.toReportDate || '');
        if (!['qb','paid'].includes(action)) return { status:400, jsonBody:{ error:'bulkAction must be qb or paid' } };
        if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) {
          return { status:400, jsonBody:{ error:'Valid From/To Report Dates are required.' } };
        }

        let items = [];
        let next = `${GRAPH}/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`;
        while (next) {
          const rr = await fetch(next, { headers:{ Authorization:`Bearer ${token}` } });
          if (!rr.ok) throw new Error(`Billing list read failed (${rr.status})`);
          const dd = await rr.json();
          items.push(...(dd.value || []));
          next = dd['@odata.nextLink'] || null;
        }

        const fieldName = action === 'qb' ? 'QB' : 'Pd';
        const matching = [];
        let outOfRange = 0;
        for (const item of items) {
          const f = item.fields || {};
          const d = reportDateToISO(f.Report_x0020_Date || f.ReportDate || '');
          if (!d || d < from || d > to) { outOfRange++; continue; }
          matching.push(item);
        }
        const pending = matching.filter(item => !(item.fields || {})[fieldName]);
        const skipped = matching.length - pending.length;

        let updated = 0;
        const failures = [];
        // Microsoft Graph batch accepts up to 20 requests. Batching keeps large
        // date ranges from requiring hundreds of sequential browser/server round trips.
        for (let offset = 0; offset < pending.length; offset += 20) {
          const chunk = pending.slice(offset, offset + 20);
          const requests = chunk.map((item, idx) => ({
            id:String(idx + 1),
            method:'PATCH',
            url:`/sites/${siteId}/lists/${listId}/items/${item.id}/fields`,
            headers:{ 'Content-Type':'application/json' },
            body:{ [fieldName]:true },
          }));
          const br = await fetch(`${GRAPH}/$batch`, {
            method:'POST',
            headers:{ Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
            body:JSON.stringify({ requests }),
          });
          if (!br.ok) {
            const err = await br.text().catch(()=>'');
            chunk.forEach(item => failures.push(`${item.fields?.Title || item.id}: batch HTTP ${br.status} ${err.slice(0,80)}`));
            continue;
          }
          const bd = await br.json();
          const responses = bd.responses || [];
          for (let idx = 0; idx < chunk.length; idx++) {
            const rr = responses.find(x => x.id === String(idx + 1));
            if (rr && rr.status >= 200 && rr.status < 300) updated++;
            else failures.push(`${chunk[idx].fields?.Title || chunk[idx].id}: HTTP ${rr?.status || 'no response'}`);
          }
        }

        const audit = await writeActivityLog({
          labId: `Report Dates ${from} to ${to}`,
          type:'Billing Bulk Updated',
          notes:`${action === 'qb' ? 'QuickBooks' : 'Paid'} checked for ${updated} row(s). Already checked: ${skipped}.${failures.length ? ' Failed: ' + failures.slice(0,5).join(', ') : ''}`,
          by:body.updatedBy || 'Admin Bulk Update',
          context,
        });

        return {
          status: failures.length ? 207 : 200,
          jsonBody:{
            success: failures.length === 0,
            updated, skipped, matched:matching.length, outOfRange,
            failures,
            auditWarning:audit.success ? null : audit.error,
          },
        };
      }

      if (!itemId || !fields) return { status: 400, jsonBody: { error: 'itemId and fields required' } };

      // Read the current row first so billing edits are recorded old -> new.
      const currentRes = await fetch(
        `${GRAPH}/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const current = currentRes.ok ? ((await currentRes.json()).fields || {}) : {};
      const labId = String(current.Title || body.labId || `Billing item ${itemId}`).trim();

      const spFields = {};
      const changes = [];
      const addChange = (label, internal, next) => {
        const prev = current[internal];
        const prevText = prev === undefined || prev === null ? '' : String(prev);
        const nextText = next === undefined || next === null ? '' : String(next);
        if (prevText !== nextText) changes.push(`${label}: "${prevText}" → "${nextText}"`);
        spFields[internal] = next;
      };

      if (fields.qb      !== undefined) addChange('QuickBooks', 'QB', !!fields.qb);
      if (fields.paid    !== undefined) addChange('Paid', 'Pd', !!fields.paid);
      if (fields.amtPaid !== undefined) addChange('Amount Paid', 'Amt_x0020_Pd', parseFloat(fields.amtPaid) || 0);
      if (fields.datePaid !== undefined) addChange('Date Paid', 'Date_x0020_Pd', fields.datePaid);
      if (fields.stmtDate !== undefined) addChange('Statement/Invoice Date', 'Statement_x002F_Inv_x0020_Date', fields.stmtDate);
      if (fields.disc !== undefined) addChange('Discount', 'Disc', parseFloat(fields.disc) || 0);

      const upRes = await fetch(
        `${GRAPH}/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
        {
          method: 'PATCH',
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify(spFields),
        }
      );

      if (!upRes.ok) {
        const err = await upRes.text();
        return { status: 500, jsonBody: { error: `SP update failed: ${err.slice(0, 100)}` } };
      }

      const audit = changes.length ? await writeActivityLog({
        labId,
        type: 'Billing Updated',
        notes: changes.join(' | '),
        by: body.updatedBy || body.changedBy || 'Lab Staff',
        context,
      }) : { success: true };

      return {
        status: 200,
        jsonBody: { success: true, auditWarning: audit.success ? null : audit.error },
      };

    } catch (err) {
      context.log('[billing-update] error:', err.message);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
