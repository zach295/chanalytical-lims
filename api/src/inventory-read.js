const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

app.http('inventory-read', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const [invItems, actItems] = await Promise.all([
        listItems(LISTS.INVENTORY).catch(() => []),
        listItems(LISTS.ACTIVITY_LOG, { top: 150 }).catch(() => []),
      ]);
      const inventory = {};
      invItems.forEach(r => {
        const key = r.ClientKey || r.Title || '';
        if (key) {
          inventory[key] = {
            inStock:       parseInt(r.InStock       || '0') || 0,
            sampled:       parseInt(r.Sampled       || '0') || 0,
            totalSent:     parseInt(r.TotalSent     || '0') || 0,
            totalReceived: parseInt(r.TotalReceived || '0') || 0,
            lastActivity:  r.LastActivity || '',
          };
        }
      });
      const activityLog = actItems.map(r => ({
        date:   r.LogDate           || '',
        time:   r.LogTime           || '',
        client: r.Client       || '',
        type:   r.ActivityType || '',
        qty: parseInt(r.Quantity || '0') || 0,
        notes:  r.Notes        || '',
        by:     r.By           || '',
      }));
      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ inventory, activityLog, bottles: {} }),
      };
    } catch(e) {
      context.log('[inventory-read] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
