const { app } = require('@azure/functions');
const { createItem, listItems, updateItem, LISTS } = require('../shared/graph');

app.http('reject-sample', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { labId, rejectionType, reason, rejectedBy } = await request.json();
      if (!labId)          return { status: 400, body: JSON.stringify({ error: 'labId required' }) };
      if (!rejectionType)  return { status: 400, body: JSON.stringify({ error: 'rejectionType required' }) };
      if (!reason?.trim()) return { status: 400, body: JSON.stringify({ error: 'reason required' }) };

      const now    = new Date().toISOString();
      const baseId = labId.split(' ')[0].trim();
      const rejNote = `${rejectionType}: ${reason.trim()}`;
      const log = [];

      // 1. Write to Rejected list — use Title only plus whatever columns exist
      await createItem(LISTS.REJECTED, {
        Title:         labId,
        LabId:         labId,
        RejectionType: rejectionType,
        Reason:        reason.trim(),
        RejectedBy:    rejectedBy || 'Lab Staff',
        Timestamp:     now,
      }).catch(e => context.log('[Rejected] Write error:', e.message));
      log.push('✅ Written to Rejected list');

      // 2. Update Archived Intake — field_X mapping
      const archived = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
      const matches  = archived.filter(r => (r.field_1 || '').startsWith(baseId));

      for (const item of matches) {
        const existingNotes = item.field_13 || '';
        const newNotes = existingNotes ? `${existingNotes} | ${rejNote}` : rejNote;
        await updateItem(LISTS.ARCHIVED_INTAKE, item._id, {
          field_2:  rejectionType,
          field_13: newNotes,
          field_14: 'Pending',
        }).catch(e => context.log('[ArchivedIntake] Update error:', e.message));
      }
      log.push(`✅ Archived Intake: updated ${matches.length} row(s)`);

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, labId, rejectionType, log }),
      };
    } catch(e) {
      context.log('[reject-sample] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
