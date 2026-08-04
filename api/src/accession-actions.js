/**
 * accession-actions.js
 * POST /api/accession-actions
 * Handles: read-clients, today-approved, list-intake, fix-field, read-raw,
 *          mark-reported, mark-pending
 */
const { app }                          = require('@azure/functions');
const { listItems, updateItem, LISTS } = require('../shared/graph');

app.http('accession-actions', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body   = await request.json().catch(() => ({}));
        const action = body.action || '';

        // ── read-clients ─────────────────────────────────────────────────────
        if (action === 'read-clients') {
          const siteId = process.env.SP_SITE_ID;
          const token  = await (require('../shared/graph').getToken)();
          const GRAPH  = 'https://graph.microsoft.com/v1.0';
          const res    = await fetch(
            `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields&$top=500`,
            { headers: { Authorization: `Bearer ${token}` } }
          );
          if (!res.ok) return { status: 200, jsonBody: { clients: [] } };
          const data    = await res.json();
          const clients = (data.value || []).map(item => {
            const f = item.fields || {};
            return {
              clientName:       f.ClientName        || f.Title || '',
              clientCode:       f.ClientCode        || '',
              reportEmail:      f.Aliases           || '',
              billingEmail:     f.Notes             || '',
              phone:            f.Active            || '',
              dbaName:          f.Phone             || '',
              abbrev:           f.Abbrev            || '',
              billingAddress:   f.BillingAddress    || '',
              pricingCategory:  f.PricingCategory   || '',
              billingFrequency: f.BillingFrequency  || '',
              frequency:        f.Frequency         || '',
              startDate:        f.StartDate         || '',
              status:           f.Status            || '',
              radonLic:         f.RadonLic_x0023_   || '',
            };
          });
          return { status: 200, jsonBody: { clients } };
        }

        // ── today-approved ───────────────────────────────────────────────────
        if (action === 'today-approved') {
          const allI     = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 }).catch(() => []);
          const todayStr = new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' });
          const approved = allI
            .filter(r => String(r.Title || '').startsWith(todayStr))
            .map(r => ({
              fullId:     r.field_1  || '',
              coaTest:    r.field_2  || '',
              customer:   r.field_3  || '',
              approvedBy: r.field_12 || '',
              timestamp:  r.Title    || '',
            }))
            .sort((a, b) => b.timestamp.localeCompare(a.timestamp))
            .slice(0, 5);
          return { status: 200, jsonBody: { approved } };
        }

        // ── list-intake ──────────────────────────────────────────────────────
        if (action === 'list-intake') {
          const searchId = body.baseId || '';
          if (!searchId) return { status: 400, jsonBody: { error: 'baseId required' } };
          const allItems = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 }).catch(() => []);
          const matches  = allItems.filter(r =>
            (r.field_1 || '').split(' ')[0].trim() === searchId
          ).map(r => ({
            _id:      r._id,
            field_1:  r.field_1  || '',
            field_2:  r.field_2  || '',
            field_3:  r.field_3  || '',
            field_4:  r.field_4  || '',
            field_5:  r.field_5  || '',
            field_6:  r.field_6  || '',
            field_7:  r.field_7  || '',
            field_8:  r.field_8  || '',
            field_9:  r.field_9  || '',
            field_10: r.field_10 || '',
            field_11: r.field_11 || '',
            field_12: r.field_12 || '',
            field_13: r.field_13 || '',
            field_14: r.field_14 || '',
            resolvedCoaTest: (r.field_2 || '').trim(),
          }));
          return { status: 200, jsonBody: { matches } };
        }

        // ── fix-field ────────────────────────────────────────────────────────
        if (action === 'fix-field') {
          const { itemId, field, value } = body;
          await updateItem(LISTS.ARCHIVED_INTAKE, itemId, { [field]: value });
          return { status: 200, jsonBody: { success: true, itemId, field, value } };
        }

        // ── read-raw ─────────────────────────────────────────────────────────
        if (action === 'read-raw') {
          const { itemId } = body;
          if (!itemId) return { status: 400, jsonBody: { error: 'itemId required' } };
          const allItems = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 }).catch(() => []);
          const found = allItems.find(r => String(r._id) === String(itemId));
          return { status: 200, jsonBody: { item: found || null } };
        }

        // ── mark-reported / mark-pending ─────────────────────────────────────
        const baseId = body.baseId || '';
        if (!baseId) return { status: 400, jsonBody: { error: 'baseId required' } };
        const isReport = action === 'mark-reported';
        const isPend   = action === 'mark-pending';
        if (!isReport && !isPend) {
          return { status: 400, jsonBody: { error: 'Unknown action: ' + action } };
        }
        const newStatus = isReport ? 'Reported' : 'Pending';
        const allItems  = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 }).catch(() => []);
        const matches   = allItems.filter(r => (r.field_1 || '').split(' ')[0].trim() === baseId);
        await Promise.all(matches.map(r =>
          updateItem(LISTS.ARCHIVED_INTAKE, r._id, { field_14: newStatus })
        ));
        return { status: 200, jsonBody: { success: true, baseId, status: newStatus, rowsUpdated: matches.length } };
      return { status: 400, jsonBody: { error: 'Unknown action' } };
        } catch(e) {
      context.log('[accession-status] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});