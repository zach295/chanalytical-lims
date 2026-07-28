const { app } = require('@azure/functions');
const { updateItem, deleteItem, LISTS } = require('../shared/graph');

app.http('mark-scan-processed', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { fileId, outcome, reviewQueueRow, rowIndex } = await request.json();
      const row = reviewQueueRow || rowIndex;
      if (!row) return { status: 400, body: JSON.stringify({ error: 'rowIndex required' }) };

      if (outcome === 'discarded') {
        await deleteItem(LISTS.REVIEW_QUEUE, row).catch(() => {});
      } else {
        await updateItem(LISTS.REVIEW_QUEUE, row, { Title: 'Processed' }).catch(() => {});
      }

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true }),
      };
    } catch(e) {
      context.log('[mark-scan-processed] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
