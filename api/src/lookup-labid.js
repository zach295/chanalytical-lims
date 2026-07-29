const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');
app.http('lookup-labid', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const input = (request.query.get('labId') || request.query.get('baseId') || '').trim();
    if (!input) return { status: 400, body: JSON.stringify({ error: 'labId required' }) };
    try {
      const items = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
      const baseInput = input.split(' ')[0].trim().toLowerCase();
      const seen = new Set();
      const results = [];
      for (const r of items) {
        const fullId = String(r.field_1 || '').trim().toLowerCase();
        const rowBase = fullId.split(' ')[0].trim();
        if (rowBase !== baseInput && fullId !== input.toLowerCase()) continue;
        if (!fullId || seen.has(fullId)) continue;
        seen.add(fullId);
        results.push({
          labId: r.field_1 || '',
          tests: r.field_2 || '',
          customer: r.field_3 || '',
          location: r.field_8 || '',
          city: r.field_9 || '',
          state: r.field_10 || 'ME',
          zip: r.field_11 ? String(r.field_11).padStart(5,'0') : '',
          dateDrawn: r.field_4 || '',
          timeDrawn: r.field_5 || '',
          receivedDate: r.field_6 || '',
          receivedTime: r.field_7 || '',
          approvedBy: r.field_12 || '',
          notes: r.field_13 || '',
          status: r.field_14 || 'Pending',
          source: 'Archived Intake',
        });
      }
      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ found: results.length > 0, count: items.length, version: 'v4', results, record: results[0] || null }),
      };
    } catch(e) {
      context.log('[lookup-labid] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
