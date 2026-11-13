const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');
const { loadClientDirectory } = require('../shared/client-identity');
const { createReportLabIdDisplayResolver } = require('../shared/report-labid-display');

app.http('get-rejections', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const [items, archived, directory] = await Promise.all([
        listItems(LISTS.REJECTED, { top: 100 }),
        listItems(LISTS.ARCHIVED_INTAKE, { top: 2000 }).catch(() => []),
        loadClientDirectory().catch(() => null),
      ]);
      const displayClient = directory ? createReportLabIdDisplayResolver(directory) : null;
      const archivedCustomers = new Map();
      for (const item of archived) {
        const baseId = String(item.field_1 || '').split(' ')[0].trim();
        if (baseId && !archivedCustomers.has(baseId)) archivedCustomers.set(baseId, item.field_3 || '');
      }

      // Sort newest first by Created date
      items.sort((a, b) => {
        const da = new Date(a.Created || 0);
        const db = new Date(b.Created || 0);
        return db - da;
      });

      // field_1=LabId, field_2=RejectionType, field_3=Reason, field_4=RejectedBy
      const rows = items.map(r => [
        r.Created || r.Title || '',
        r.field_1 || '',
        r.field_2 || '',
        r.field_3 || '',
        r.field_4 || '',
        (() => {
          const original = archivedCustomers.get(String(r.field_1 || '').split(' ')[0].trim()) || '';
          return (displayClient && displayClient(original)) || original;
        })(),
      ]);

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ rows }),
      };
    } catch(e) {
      context.log('[get-rejections] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
