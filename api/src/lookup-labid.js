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

      // Debug mode — show raw fields of first item
      if (input === 'debug') {
        return { status: 200, headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ sample: items.slice(0, 2) }) };
      }

      const base = input.split(' ')[0].trim().toLowerCase();
      const seen  = new Set();
      const results = items
        .filter(r => {
          // Try both field_X and direct column names
          const fullId = (r.field_1 || r['Lab ID'] || r.LabID || r.Title || '').toLowerCase();
          const rowBase = fullId.split(' ')[0].trim();
          return rowBase === base || fullId === input.toLowerCase();
        })
        .map(r => ({
          labId:        r.field_1  || r['Lab ID'] || r.LabID || '',
          tests:        r.field_2  || r.Test      || '',
          customer:     r.field_3  || r.Client    || '',
          location:     r.field_8  || r.Address   || '',
          city:         r.field_9  || r.City      || '',
          state:        r.field_10 || r.State     || 'ME',
          zip:          r.field_11 || r.Zip       || '',
          dateDrawn:    r.field_4  || r['Date collected']  || '',
          timeDrawn:    r.field_5  || r['Time Collected']  || '',
          receivedDate: r.field_6  || r['Date Recieved']   || '',
          receivedTime: r.field_7  || r['Time Recieved']   || '',
          approvedBy:   r.field_12 || r['Approved By']     || '',
          notes:        r.field_13 || r.Notes     || '',
          status:       r.field_14 || r.Status    || 'Pending',
          source: 'Archived Intake',
        }))
        .filter(r => {
          if (!r.labId || seen.has(r.labId)) return false;
          seen.add(r.labId);
          return true;
        });

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ found: results.length > 0, results, record: results[0] || null }),
      };
    } catch(e) {
      context.log('[lookup-labid] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
