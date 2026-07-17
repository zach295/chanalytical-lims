const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

app.http('get-rejections', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const items = await listItems(LISTS.REJECTED, { orderby: 'fields/Created desc', top: 100 });
      const rows = items.map(r => [
        r.Timestamp || '',
        r.LabId     || '',
        r.RejectionType || '',
        r.Reason    || '',
        r.RejectedBy || '',
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
