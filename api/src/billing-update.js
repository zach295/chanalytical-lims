/**
 * billing-update.js — updates QB/paid status on Reports to be Billed list
 */
const { app } = require('@azure/functions');
const { getToken } = require('../shared/graph');
const { writeActivityLog } = require('../shared/audit');

const GRAPH = 'https://graph.microsoft.com/v1.0';

app.http('billing-update', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body = await request.json().catch(() => ({}));
      const { itemId, fields } = body;
      if (!itemId || !fields) return { status: 400, jsonBody: { error: 'itemId and fields required' } };

      const token = await getToken();
      const siteId = process.env.SP_SITE_ID;

      const res = await fetch(
        `${GRAPH}/sites/${siteId}/lists?$select=id,displayName`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const listId = ((await res.json()).value || [])
        .find(l => l.displayName === 'Reports to be Billed')?.id;
      if (!listId) return { status: 404, jsonBody: { error: 'List not found' } };

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
