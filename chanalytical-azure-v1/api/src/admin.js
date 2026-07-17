const { app } = require('@azure/functions');
const { listItems, createItem, updateItem, findItem } = require('../shared/graph');

const INV_LIST = 'Client Inventory';
const ACT_LIST = 'Activity Log';

// ── ADMIN READ ─────────────────────────────────────────────────────────────────
app.http('admin-read', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const [invItems, actItems] = await Promise.all([
        listItems(INV_LIST).catch(() => []),
        listItems(ACT_LIST, { orderby: 'fields/Created desc', top: 150 }).catch(() => []),
      ]);

      const inventory = {};
      invItems.forEach(r => {
        const key = r.ClientKey || r.Title || '';
        if (key) {
          inventory[key] = {
            inStock:       parseInt(r.InStock   || '0') || 0,
            sampled:       parseInt(r.Sampled   || '0') || 0,
            totalSent:     parseInt(r.TotalSent || '0') || 0,
            totalReceived: parseInt(r.TotalReceived || '0') || 0,
            lastActivity:  r.LastActivity || '',
          };
        }
      });

      const activityLog = actItems.map(r => ({
        date:   r.Date   || '',
        time:   r.Time   || '',
        client: r.Client || '',
        type:   r.Type   || '',
        qty:    r.Qty    || 0,
        notes:  r.Notes  || '',
        by:     r.By     || '',
      }));

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ inventory, activityLog }),
      };
    } catch(e) {
      context.log('[admin-read] Error:', e.message);
      return { status:500, body: JSON.stringify({ error: e.message }) };
    }
  }
});

// ── ADMIN WRITE ────────────────────────────────────────────────────────────────
app.http('admin-write', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body = await request.json();
      const { action } = body;

      if (action === 'log_activity') {
        const { entry } = body.payload || body;
        await createItem(ACT_LIST, {
          Title:  `${entry.date} ${entry.client}`,
          Date:   entry.date   || '',
          Time:   entry.time   || '',
          Client: entry.client || '',
          Type:   entry.type   || '',
          Qty:    entry.qty    || 0,
          Notes:  entry.notes  || '',
          By:     entry.by     || '',
        });
        return { status:200, headers:{'content-type':'application/json'}, body: JSON.stringify({ success:true }) };
      }

      if (action === 'write_inventory') {
        const { inventory } = body;
        for (const [key, data] of Object.entries(inventory || {})) {
          const existing = await findItem(INV_LIST, 'ClientKey', key).catch(() => null);
          const fields = {
            Title: key, ClientKey: key,
            InStock: data.inStock || 0, Sampled: data.sampled || 0,
            TotalSent: data.totalSent || 0, TotalReceived: data.totalReceived || 0,
            LastActivity: data.lastActivity || '',
          };
          if (existing) await updateItem(INV_LIST, existing._id, fields);
          else          await createItem(INV_LIST, fields);
        }
        return { status:200, headers:{'content-type':'application/json'}, body: JSON.stringify({ success:true }) };
      }

      return { status:400, body: JSON.stringify({ error:'Unknown action' }) };
    } catch(e) {
      context.log('[admin-write] Error:', e.message);
      return { status:500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
