/**
 * billing-update.js — updates QB/paid status on Reports to be Billed list
 */
const { app }    = require('@azure/functions');
const { getToken } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

app.http('billing-update', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body   = await request.json().catch(() => ({}));
      const { itemId, fields } = body;
      if (!itemId || !fields) return { status: 400, jsonBody: { error: 'itemId and fields required' } };

      const token  = await getToken();
      const siteId = process.env.SP_SITE_ID;

      // Get list ID
      const res    = await fetch(
        `${GRAPH}/sites/${siteId}/lists?$select=id,displayName`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const listId = ((await res.json()).value || [])
        .find(l => l.displayName === 'Reports to be Billed')?.id;
      if (!listId) return { status: 404, jsonBody: { error: 'List not found' } };

      // Map clean field names back to SharePoint column names
      const spFields = {};
      if (fields.qb       !== undefined) spFields.QB                              = !!fields.qb;
      if (fields.paid      !== undefined) spFields.Pd                              = !!fields.paid;
      if (fields.amtPaid   !== undefined) spFields.Amt_x0020_Pd                   = parseFloat(fields.amtPaid) || 0;
      if (fields.datePaid  !== undefined) spFields.Date_x0020_Pd                  = fields.datePaid;
      if (fields.stmtDate  !== undefined) spFields.Statement_x002F_Inv_x0020_Date = fields.stmtDate;
      if (fields.disc      !== undefined) spFields.Disc                            = parseFloat(fields.disc) || 0;

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

      return { status: 200, jsonBody: { success: true } };

    } catch (err) {
      context.log('[billing-update] error:', err.message);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
