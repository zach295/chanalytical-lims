const { app } = require('@azure/functions');
const { listItems, findItem, updateItem, LISTS } = require('../shared/graph');

app.http('accession-status', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      // ── GET — list pending and sent lab IDs ─────────────────────────────────
      if (request.method === 'GET') {
        const items = await listItems(LISTS.ARCHIVED_INTAKE, {
          orderby: 'fields/Created asc',
          top: 500,
        });

        const byBase = {};
        items.forEach(r => {
          const fullId = (r.FullId || '').trim();
          if (!fullId) return;
          const baseId = fullId.replace(/\s+\S+$/, '').trim();
          if (!byBase[baseId]) {
            byBase[baseId] = {
              baseId,
              fullIds:      [],
              tests:        [],
              customer:     r.Customer     || '',
              location:     r.Location     || '',
              city:         r.City         || '',
              state:        r.State        || 'ME',
              zip:          r.Zip          || '',
              dateDrawn:    r.DateDrawn    || '',
              timeDrawn:    r.TimeDrawn    || '',
              dateReceived: r.ReceivedDate || '',
              timeReceived: r.ReceivedTime || '',
              reviewedBy:   r.ReviewedBy   || '',
              status:       'Pending',
              timestamp:    r.Timestamp    || '',
              _itemIds:     [],
            };
          }
          byBase[baseId]._itemIds.push(r._id);
          if (fullId && !byBase[baseId].fullIds.includes(fullId)) byBase[baseId].fullIds.push(fullId);
          if (r.CoaTest && !byBase[baseId].tests.includes(r.CoaTest)) byBase[baseId].tests.push(r.CoaTest);
          const status = (r.ReportStatus || 'Pending').trim();
          if (status === 'Sent' || status === 'Reported') byBase[baseId].status = 'Sent';
        });

        const all = Object.values(byBase);
        const pending  = all.filter(r => r.status === 'Pending');
        const reported = all.filter(r => r.status === 'Sent');

        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ pending, reported }),
        };
      }

      // ── POST — mark reported / pending ─────────────────────────────────────
      if (request.method === 'POST') {
        const body = await request.json();
        const { action, baseId } = body;
        if (!baseId) return { status: 400, body: JSON.stringify({ error: 'baseId required' }) };

        const status = action === 'mark-reported' ? 'Sent' : 'Pending';
        const items = await listItems(LISTS.ARCHIVED_INTAKE);
        const matches = items.filter(r =>
          (r.FullId || '').trim().startsWith(baseId)
        );

        for (const item of matches) {
          await updateItem(LISTS.ARCHIVED_INTAKE, item._id, { ReportStatus: status });
        }

        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: true, baseId, status, rowsUpdated: matches.length }),
        };
      }

      return { status: 405, body: 'Method Not Allowed' };
    } catch(e) {
      context.log('[accession-status] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
