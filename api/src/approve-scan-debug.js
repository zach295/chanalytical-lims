const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

// Temporary debug endpoint — delete after field names confirmed
app.http('approve-debug', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const [accItems, intakeItems, cacheItems] = await Promise.all([
        listItems(LISTS.ACCESSION_LOG, { top: 1 }).catch(() => []),
        listItems(LISTS.ARCHIVED_INTAKE, { top: 1 }).catch(() => []),
        listItems('Results Cache', { top: 1 }).catch(() => []),
      ]);
      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          accessionLog:   accItems[0]   || null,
          archivedIntake: intakeItems[0] || null,
          resultsCache:   cacheItems[0]  || null,
        }),
      };
    } catch(e) {
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
