const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

app.http('lookup-labid', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const input  = (request.query.get('labId') || request.query.get('baseId') || '').trim();
    if (!input) return { status: 400, body: JSON.stringify({ error: 'labId required' }) };
    try {
      // Strip suffix — match on base ID (MMDDYY-NNN) only
      const base = input.split(' ')[0].trim().toLowerCase();
      const items = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
      const seen  = new Set();
      const results = items
        .filter(r => {
          // field_1 = fullId (e.g. "072826-001 EXP")
          // Match if fullId starts with the base ID the user typed
          const fullId = (r.field_1 || '').toLowerCase();
          const rowBase = fullId.split(' ')[0].trim();
          return rowBase === base || fullId === input.toLowerCase();
        })
        .map(r => ({
          labId:       r.field_1  || '',
          tests:       r.field_2  || '',
          customer:    r.field_3  || '',
          location:    r.field_8  || '',
          city:        r.field_9  || '',
          state:       r.field_10 || 'ME',
          zip:         r.field_11 ? String(r.field_11).padStart(5,'0') : '',
          dateDrawn:   r.field_4  || '',
          timeDrawn:   r.field_5  || '',
          receivedDate: r.field_6 || '',
          receivedTime: r.field_7 || '',
          approvedBy:  r.field_12 || '',
          notes:       r.field_13 || '',
          status:      r.field_14 || 'Pending',
          source: 'Archived Intake',
        }))
        .filter(r => {
          if (!r.labId || seen.has(r.labId)) return false;
          seen.add(r.labId);
          return true;
        });

      const found = results.length > 0;
      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ found, results, record: results[0] || null }),
      };
    } catch(e) {
      context.log('[lookup-labid] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
