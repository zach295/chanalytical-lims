const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

app.http('inventory-read', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const items = await listItems(LISTS.INVENTORY).catch(() => []);
      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ items }),
      };
    } catch(e) {
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
