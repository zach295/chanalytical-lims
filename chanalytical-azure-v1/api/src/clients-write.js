const { app } = require('@azure/functions');
const { createItem, updateItem, deleteItem, findItem, LISTS } = require('../shared/graph');

app.http('clients-write', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body = await request.json();
      const { action, client } = body;

      if (action === 'create') {
        const item = await createItem(LISTS.CLIENTS, {
          Title:      client.clientName,
          ClientName: client.clientName,
          ClientCode: client.clientCode || '',
          Abbrev:     client.abbrev     || '',
          Email:      client.email      || '',
          Aliases:    client.aliases    || '',
          Phone:      client.phone      || '',
          Active:     'TRUE',
          Notes:      client.notes      || '',
        });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true, id: item._id }) };
      }

      if (action === 'update') {
        await updateItem(LISTS.CLIENTS, client._id, {
          Title:      client.clientName,
          ClientName: client.clientName,
          ClientCode: client.clientCode || '',
          Abbrev:     client.abbrev     || '',
          Email:      client.email      || '',
          Aliases:    client.aliases    || '',
          Phone:      client.phone      || '',
          Active:     client.active !== false ? 'TRUE' : 'FALSE',
          Notes:      client.notes      || '',
        });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      return { status: 400, body: JSON.stringify({ error: 'Unknown action: ' + action }) };
    } catch(e) {
      context.log('[clients-write] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
