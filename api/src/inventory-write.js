const { app } = require('@azure/functions');
const { createItem, updateItem, findItem, LISTS } = require('../shared/graph');

app.http('inventory-write', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body = await request.json();
      const { action } = body;

      if (action === 'log_activity') {
        const entry = body.payload?.entry || body.entry || body.payload || body;
        const now = new Date();
        const serverDate = now.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
        const serverTime = now.toLocaleTimeString('en-US', { timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
        const logDate = entry.date || serverDate;
        const logTime = entry.time || serverTime;
        await createItem(LISTS.ACTIVITY_LOG, {
          Title:        `${logDate} ${entry.client || ''}`.trim(),
          LogDate:      logDate,
          LogTime:      logTime,
          Client:       entry.client || '',
          ActivityType: entry.type   || '',
          Quantity: Number(entry.qty) || 0,
          Notes:        entry.notes  || '',
          By:           entry.by     || 'Lab Staff',
        });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'write_inventory') {
        const { inventory } = body;
        for (const [key, data] of Object.entries(inventory || {})) {
          const existing = await findItem(LISTS.INVENTORY, 'Title', key).catch(() => null);
          const fields = {
            Title: key, ClientKey: key,
            InStock: data.inStock || 0, Sampled: data.sampled || 0,
            TotalSent: data.totalSent || 0, TotalReceived: data.totalReceived || 0,
            LastActivity: data.lastActivity || '',
          };
          if (existing) await updateItem(LISTS.INVENTORY, existing._id, fields);
          else await createItem(LISTS.INVENTORY, fields);
        }
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'write_bottles') {
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      return { status: 400, body: JSON.stringify({ error: 'Unknown action: ' + action }) };
    } catch(e) {
      context.log('[inventory-write] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
