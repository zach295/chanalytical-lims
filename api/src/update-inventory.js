/**
 * update-inventory.js — Azure version
 * Called from field app (index.html) when a sample is submitted.
 * Decrements inStock, increments sampled for the client in SP Client Inventory list.
 * Also logs to SP Activity Log list.
 *
 * POST { clientKey, sampleId, dateDrawn, timeDrawn, customerName }
 */
const { app }   = require('@azure/functions');
const { findItem, createItem, updateItem, LISTS } = require('../shared/graph');

function formatDate(dateStr) {
  if (!dateStr) return '';
  try {
    const [y, m, d] = dateStr.split('-');
    return `${m}-${d}-${y.slice(-2)}`;
  } catch { return dateStr; }
}

app.http('update-inventory', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { clientKey, sampleId, dateDrawn, timeDrawn, customerName } = await request.json();

      // Skip general/walk-in submissions — no inventory to track
      if (!clientKey || clientKey === 'general') {
        return { status: 200, jsonBody: { skipped: true } };
      }

      const formattedDate = formatDate(dateDrawn);
      const today = formatDate(new Date().toISOString().split('T')[0]);
      const activityDate = formattedDate || today;

      // ── Update Client Inventory ──────────────────────────────────────────────
      const existing = await findItem(LISTS.INVENTORY, 'ClientKey', clientKey).catch(() => null);

      if (existing) {
        const inStock = Math.max(0, (parseInt(existing.InStock) || 0) - 1);
        const sampled  = (parseInt(existing.Sampled) || 0) + 1;
        await updateItem(LISTS.INVENTORY, existing._id, {
          InStock:      inStock,
          Sampled:      sampled,
          LastActivity: activityDate,
        });
      } else {
        // Client not in inventory yet — create row
        await createItem(LISTS.INVENTORY, {
          Title:        clientKey,
          ClientKey:    clientKey,
          InStock:      0,
          Sampled:      1,
          TotalSent:    0,
          TotalReceived:0,
          LastActivity: activityDate,
        });
      }

      // ── Log to Activity Log ──────────────────────────────────────────────────
      const now      = new Date();
      const timeStr  = now.toLocaleTimeString('en-US', { hour:'2-digit', minute:'2-digit', hour12:true });
      const ts       = now.toISOString();

      await createItem(LISTS.ACTIVITY_LOG, {
        Title:  `${activityDate} ${clientKey}`,
        Date:   activityDate,
        Time:   timeStr,
        Client: clientKey,
        Type:   'sampled',
        Qty:    1,
        Notes:  `Sample ${sampleId || ''}`,
        By:     customerName || clientKey,
      }).catch(e => context.log('[update-inventory] Activity log failed:', e.message));

      context.log(`[update-inventory] ${clientKey} — inStock-1, sampled+1`);

      return { status: 200, jsonBody: { success: true } };

    } catch (err) {
      context.log('[update-inventory] Error:', err.message);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
