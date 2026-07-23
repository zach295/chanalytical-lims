const { app } = require('@azure/functions');
const { updateItem, deleteItem, LISTS } = require('../shared/graph');

app.http('mark-scan-processed', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { fileId, outcome, reviewQueueRow } = await request.json();
      if (!fileId) return { status: 400, body: JSON.stringify({ error: 'fileId required' }) };
      if (outcome === 'discarded' && reviewQueueRow) {
        await deleteItem(LISTS.REVIEW_QUEUE, reviewQueueRow).catch(() => {});
      } else if (reviewQueueRow) {
        await updateItem(LISTS.REVIEW_QUEUE, reviewQueueRow, { ReviewStatus: 'Processed' }).catch(() => {});
      }
      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true }),
      };
    } catch (e) {
      context.log('[mark-scan-processed] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
