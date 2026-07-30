/**
 * sync-results-cache.js
 * One-time (or periodic) backfill:
 * Reads Archived Intake, finds any base IDs that don't have a
 * Results Cache row, and creates stub entries.
 *
 * GET /api/sync-results-cache?dry=true  → preview only
 * GET /api/sync-results-cache           → create stubs
 */
const { app } = require('@azure/functions');
const { listItems, createItem, LISTS } = require('../shared/graph');

app.http('sync-results-cache', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const dry = request.query.get('dry') === 'true';

      // Get all Archived Intake base IDs
      const intakeItems = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
      const intakeIds   = new Set();
      for (const r of intakeItems) {
        const fullId = (r.field_1 || '').trim();
        if (!fullId) continue;
        const baseId = fullId.split(' ')[0].trim();
        if (/^\d{6}-\d{3}$/.test(baseId)) intakeIds.add(baseId);
      }

      // Get existing Results Cache IDs
      const cacheItems = await listItems('Results Cache', { top: 500 });
      const cacheIds   = new Set();
      for (const r of cacheItems) {
        const id = String(r.LabID || '').split(' ')[0].trim();
        if (id) cacheIds.add(id);
      }

      // Find missing IDs
      const missing = [...intakeIds].filter(id => !cacheIds.has(id)).sort();

      context.log(`[sync] Intake: ${intakeIds.size}, Cache: ${cacheIds.size}, Missing: ${missing.length}`);

      if (dry) {
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ dry: true, intakeCount: intakeIds.size, cacheCount: cacheIds.size, missingCount: missing.length, missing }),
        };
      }

      // Create stub entries
      const log = []; let created = 0, errors = 0;
      for (const baseId of missing) {
        await createItem('Results Cache', { LabID: baseId })
          .then(() => { created++; log.push(`Created: ${baseId}`); })
          .catch(e => { errors++; log.push(`Error ${baseId}: ${e.message}`); });
      }

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, created, errors, log }),
      };

    } catch(e) {
      context.log('[sync-results-cache] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
