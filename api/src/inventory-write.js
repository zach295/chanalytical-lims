const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

const BOTTLE_LIST = 'Bottle Inventory';

app.http('inventory-read', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const [invItems, actItems, bottleItems] = await Promise.all([
        listItems(LISTS.INVENTORY).catch(() => []),
        listItems(LISTS.ACTIVITY_LOG, { top: 150 }).catch(() => []),
        listItems(BOTTLE_LIST).catch(() => []),
      ]);

      const inventory = {};
      invItems.forEach(r => {
        context.log('[inv raw]', JSON.stringify(r));
        const key = r.ClientKey || r.Title || '';
        if (key) {
          inventory[key] = {
            inStock:       parseInt(r.InStock       || r.InStock0       || '0') || 0,
            sampled:       parseInt(r.Sampled       || r.Sampled0       || '0') || 0,
            totalSent:     parseInt(r.TotalSent     || r.TotalSent0     || '0') || 0,
            totalReceived: parseInt(r.TotalReceived || r.TotalReceived0 || '0') || 0,
            lastActivity:  r.LastActivity || '',
          };
        }
      });

      const activityLog = actItems.map(r => ({
        date:   r.Date         || '',
        time:   r.Time         || '',
        client: r.Client       || '',
        type:   r.ActivityType || '',
        qty:    parseInt(r.Qty || r.Qty0 || '0') || 0,
        notes:  r.Notes        || '',
        by:     r.By           || '',
      }));

      const bottles = {};
      bottleItems.forEach(r => {
        context.log('[bottles raw]', JSON.stringify(r));
        const key = r.BottleKey || r.Title || '';
        if (key) {
          bottles[key] = {
            count: parseInt(r.Count || r.Count0 || r.count || '0') || 0,
            label: r.Label || key,
          };
        }
      });

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ inventory, activityLog, bottles }),
      };
    } catch(e) {
      context.log('[inventory-read] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
