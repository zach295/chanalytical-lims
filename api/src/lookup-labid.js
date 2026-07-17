const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

app.http('lookup-labid', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const input = request.query.get('labId') || '';
    if (!input) return { status: 400, body: JSON.stringify({ error: 'labId required' }) };

    try {
      const base = input.split(' ')[0].trim().toLowerCase();
      const items = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
      const seen = new Set();
      const results = items
        .filter(r => {
          const id = (r.FullId || '').toLowerCase();
          return id.startsWith(base) || id === input.toLowerCase();
        })
        .map(r => ({ labId: r.FullId, tests: r.CoaTest, customer: r.Customer, source: 'Archived Intake' }))
        .filter(r => { if (seen.has(r.labId)) return false; seen.add(r.labId); return true; });

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ results }),
      };
    } catch(e) {
      context.log('[lookup-labid] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
