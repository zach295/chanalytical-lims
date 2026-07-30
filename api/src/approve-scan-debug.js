const { app } = require('@azure/functions');
const { listItems, getToken, getListId, LISTS } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

app.http('approve-debug', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const [accItems, intakeItems] = await Promise.all([
        listItems(LISTS.ACCESSION_LOG,   { top: 1 }).catch(() => []),
        listItems(LISTS.ARCHIVED_INTAKE, { top: 1 }).catch(() => []),
      ]);

      // Get Results Cache column definitions to find internal field names
      const token  = await getToken();
      const siteId = process.env.SP_SITE_ID;
      const listId = await getListId('Results Cache');
      const colRes = await fetch(
        `${GRAPH}/sites/${siteId}/lists/${listId}/columns?$select=name,displayName,columnGroup`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const colData = colRes.ok ? await colRes.json() : { value: [] };
      const columns = (colData.value || [])
        .filter(c => !c.columnGroup || c.columnGroup !== 'Hidden')
        .map(c => ({ name: c.name, displayName: c.displayName }));

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          accessionLog:    accItems[0]   || null,
          archivedIntake:  intakeItems[0] || null,
          resultsCacheColumns: columns,
        }),
      };
    } catch(e) {
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
