const { app } = require('@azure/functions');
const { getToken, getListId } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

app.http('list-columns', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const listName = request.query.get('list') || 'Results Cache';
      const token    = await getToken();
      const siteId   = process.env.SP_SITE_ID;
      const listId   = await getListId(listName);
      const res = await fetch(
        `${GRAPH}/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const data = res.ok ? await res.json() : { value: [], error: await res.text() };
      const columns = (data.value || []).map(c => ({ name: c.name, displayName: c.displayName }));
      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ listName, columns }),
      };
    } catch(e) {
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
