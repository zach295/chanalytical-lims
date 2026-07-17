const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

app.http('clients-read', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const items = await listItems(LISTS.CLIENTS, { orderby: 'fields/ClientName asc' });
      const clients = items.map(r => ({
        _id:        r._id,
        clientName: r.ClientName  || r.Title || '',
        clientCode: r.ClientCode  || '',
        abbrev:     r.Abbrev      || '',
        email:      r.Email       || '',
        aliases:    r.Aliases     || '',
        phone:      r.Phone       || '',
        active:     r.Active !== false && r.Active !== 'FALSE',
        notes:      r.Notes       || '',
      }));
      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ clients }),
      };
    } catch(e) {
      context.log('[clients-read] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
