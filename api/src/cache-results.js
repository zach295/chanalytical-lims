const { app } = require('@azure/functions');
const { createItem, updateItem, deleteItem, findItem } = require('../shared/graph');
const { writeActivityLog } = require('../shared/audit');

const LIST = 'Results Cache';

function summarizeChanges(before, after) {
  const keys = new Set([...Object.keys(before || {}), ...Object.keys(after || {})]);
  const out = [];
  for (const key of keys) {
    const a = before?.[key];
    const b = after?.[key];
    const aText = typeof a === 'object' ? JSON.stringify(a) : String(a ?? '');
    const bText = typeof b === 'object' ? JSON.stringify(b) : String(b ?? '');
    if (aText !== bText) out.push(`${key}: "${aText}" → "${bText}"`);
  }
  return out;
}

app.http('cache-results', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (request.method === 'GET') {
        const labId = request.query.get('labId') || '';
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
        const baseId = (labId || '').match(/(\d{6}-\d{3})/)?.[1] || labId;
        if (!baseId) return { status: 400, body: JSON.stringify({ error: 'labId required' }) };
        const actor = body.updatedBy || body.changedBy || body.importedBy || 'Lab Staff';

        if (action === 'delete') {
          const item = await findItem(LIST, 'LabId', baseId).catch(() => null);
          let prior = {};
          if (item?.Data) { try { prior = JSON.parse(item.Data); } catch {} }
          if (item) await deleteItem(LIST, item._id);
          const audit = await writeActivityLog({
            labId: baseId,
            type: 'Results Cache Deleted',
            notes: item ? `Results Cache record deleted. Prior top-level result groups: ${Object.keys(prior).join(', ') || 'none'}` : 'Delete requested; no Results Cache record existed.',
            by: actor,
            context,
          });
          return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true, auditWarning: audit.success ? null : audit.error }) };
        }

        if (action === 'write') {
          const existing = await findItem(LIST, 'LabId', baseId).catch(() => null);
          let before = {};
          if (existing?.Data) { try { before = JSON.parse(existing.Data); } catch {} }
          let merged = { ...before, ...(results || {}) };
          if (results?.gallery && merged.gallery) merged.gallery = { ...(before.gallery || {}), ...results.gallery };
          if (results?.icpms && merged.icpms) merged.icpms = { ...(before.icpms || {}), ...results.icpms };

          const fields = { Title: baseId, LabId: baseId, Data: JSON.stringify(merged), Timestamp: new Date().toISOString() };
          if (existing) await updateItem(LIST, existing._id, fields);
          else await createItem(LIST, fields);

          const changes = summarizeChanges(before, merged);
          const audit = await writeActivityLog({
            labId: baseId,
            type: existing ? 'Results Cache Updated' : 'Results Cache Created',
            notes: changes.length ? changes.join(' | ') : 'Results Cache write completed with no material value change.',
            by: actor,
            context,
          });

          return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true, labId: baseId, auditWarning: audit.success ? null : audit.error }) };
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
