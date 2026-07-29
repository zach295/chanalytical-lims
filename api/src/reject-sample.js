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

      // 1. Write to Rejected list
      // Try direct field names first (manually created list), fall back handles field_X
      await createItem(LISTS.REJECTED, {
        Title:         labId,
        LabId:         labId,
        RejectionType: rejectionType,
        Reason:        reason.trim(),
        RejectedBy:    rejectedBy || 'Lab Staff',
        Timestamp:     now,
      }).catch(async () => {
        // If field names fail, try field_X naming
        await createItem(LISTS.REJECTED, {
          Title:   labId,
          field_1: labId,
          field_2: rejectionType,
          field_3: reason.trim(),
          field_4: rejectedBy || 'Lab Staff',
          field_5: now,
        });
      });
      log.push('✅ Written to Rejected list');

      // 2. Update Archived Intake — uses field_X mapping
      // field_1=fullId, field_2=coaTest, field_13=notes, field_14=status
      const archived = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
      const matches  = archived.filter(r => (r.field_1 || '').startsWith(baseId));

      for (const item of matches) {
        const existingNotes = item.field_13 || '';
        const newNotes = existingNotes ? `${existingNotes} | ${rejNote}` : rejNote;
        await updateItem(LISTS.ARCHIVED_INTAKE, item._id, {
          field_2:  rejectionType,
          field_13: newNotes,
          field_14: 'Pending',
        });
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
