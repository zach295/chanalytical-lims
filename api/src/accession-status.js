/**
 * accession-status.js
 * GET  → { pending, reported }
 * POST → various actions
 */
const { app }                                = require('@azure/functions');
const { listItems, updateItem, LISTS }       = require('../shared/graph');

function fmtName(name) {
  if (!name || !name.startsWith('Public-')) return name || '';
  const inner = name.slice(7).trim();
  const ci = inner.indexOf(',');
  if (ci > 0) {
    const last = inner.slice(0, ci).trim();
    const first = inner.slice(ci + 1).trim();
    return first ? `${first} ${last}` : last;
  }
  return inner;
}

app.http('accession-status', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const items = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
        const byBase = {};
        for (const r of items) {
          const fullId   = (r.field_1  || '').trim();
          const coaTest  = (r.field_2  || '').trim();
          const customer = (r.field_3  || '').trim();
          const status   = (r.field_14 || 'Pending').trim();
          if (!fullId) continue;
          const baseId = fullId.split(' ')[0].trim();
          if (!byBase[baseId]) {
            byBase[baseId] = {
              baseId,
              fullIds:    [],
              tests:      [],
              customer:   fmtName(customer),
              rawCustomer: customer,
              approvedBy: r.field_12 || '',
              timestamp:  r.Title    || '',
              status:     'Pending',
            };
          }
          byBase[baseId].fullIds.push(fullId);
          if (coaTest && !byBase[baseId].tests.includes(coaTest)) {
            byBase[baseId].tests.push(coaTest);
          }
          if (status && status !== 'Pending') byBase[baseId].status = status;
        }
        const all      = Object.values(byBase);
        const pending  = all.filter(r => r.status === 'Pending');
        const reported = all.filter(r => r.status !== 'Pending');
        return {
          status: 200,
          headers: {
            'content-type': 'application/json',
            'Cache-Control': 'no-store',
          },
          body: JSON.stringify({ pending, reported }),
        };
        } catch(e) {
      context.log('[accession-status] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
