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

      const now = new Date().toISOString();
      const baseId = labId.split(' ')[0].trim();
      const rejNote = `${rejectionType}: ${reason}`;
      const log = [];

      // 1. Write to Rejected list
      await createItem(LISTS.REJECTED, {
        Title:       labId,
        LabId:       labId,
        RejectionType: rejectionType,
        Reason:      reason.trim(),
        RejectedBy:  rejectedBy || 'Lab Staff',
        Timestamp:   now,
      });
      log.push('✅ Written to Rejected list');

      // 2. Update Archived Intake — replace CoaTest + append Notes
      const archived = await listItems(LISTS.ARCHIVED_INTAKE);
      const matches = archived.filter(r => (r.FullId||'').startsWith(baseId));
      for (const item of matches) {
        const newNotes = item.Notes ? `${item.Notes} | ${rejNote}` : rejNote;
        await updateItem(LISTS.ARCHIVED_INTAKE, item._id, {
          CoaTest:      rejectionType,
          Notes:        newNotes,
          ReportStatus: 'Pending',
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
