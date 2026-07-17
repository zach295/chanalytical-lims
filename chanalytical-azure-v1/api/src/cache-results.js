const { app } = require('@azure/functions');
const { listItems, createItem, updateItem, deleteItem, findItem } = require('../shared/graph');

const LIST = 'Results Cache';

app.http('cache-results', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (request.method === 'GET') {
        const labId  = request.query.get('labId') || '';
        const baseId = labId.match(/(\d{6}-\d{3})/)?.[1] || labId;
        if (!baseId) return { status: 400, body: JSON.stringify({ error: 'labId required' }) };

        const item = await findItem(LIST, 'LabId', baseId).catch(() => null);
        if (!item) return { status: 404, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ found: false, labId: baseId }) };

        const data = JSON.parse(item.Data || '{}');
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ found: true, labId: baseId, data }) };
      }

      if (request.method === 'POST') {
        const body = await request.json();
        const { action, labId, results } = body;
        const baseId = (labId||'').match(/(\d{6}-\d{3})/)?.[1] || labId;
        if (!baseId) return { status: 400, body: JSON.stringify({ error: 'labId required' }) };

        if (action === 'delete') {
          const item = await findItem(LIST, 'LabId', baseId).catch(() => null);
          if (item) await deleteItem(LIST, item._id);
          return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
        }

        if (action === 'write') {
          const existing = await findItem(LIST, 'LabId', baseId).catch(() => null);
          let merged = {};
          if (existing?.Data) { try { merged = JSON.parse(existing.Data); } catch {} }
          merged = { ...merged, ...results };
          if (results.gallery && merged.gallery) merged.gallery = { ...merged.gallery, ...results.gallery };
          if (results.icpms   && merged.icpms)   merged.icpms   = { ...merged.icpms,   ...results.icpms   };

          const fields = { Title: baseId, LabId: baseId, Data: JSON.stringify(merged), Timestamp: new Date().toISOString() };
          if (existing) await updateItem(LIST, existing._id, fields);
          else          await createItem(LIST, fields);

          return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true, labId: baseId }) };
        }

        return { status: 400, body: JSON.stringify({ error: 'Unknown action' }) };
      }

      return { status: 405, body: 'Method Not Allowed' };
    } catch(e) {
      context.log('[cache-results] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
