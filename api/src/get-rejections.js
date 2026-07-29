const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

app.http('get-rejections', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const items = await listItems(LISTS.REJECTED, { top: 100 });
      const rows = items.map(r => [
        // Try both direct names and field_X — handles both manually created and Excel-imported lists
        r.Timestamp     || r.field_5 || r.Created || '',
        r.LabId         || r.field_1 || r.Title   || '',
        r.RejectionType || r.field_2 || '',
        r.Reason        || r.field_3 || '',
        r.RejectedBy    || r.field_4 || '',
      ]);
      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ rows }),
      };
    } catch(e) {
      context.log('[get-rejections] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
