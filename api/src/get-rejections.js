const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

app.http('get-rejections', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const items = await listItems(LISTS.REJECTED, { top: 100 });

      // Sort newest first by Created date
      items.sort((a, b) => {
        const da = new Date(a.Created || 0);
        const db = new Date(b.Created || 0);
        return db - da;
      });

      // field_1=LabId, field_2=RejectionType, field_3=Reason, field_4=RejectedBy
      const rows = items.map(r => [
        r.Created || r.Title || '',
        r.field_1 || '',
        r.field_2 || '',
        r.field_3 || '',
        r.field_4 || '',
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
