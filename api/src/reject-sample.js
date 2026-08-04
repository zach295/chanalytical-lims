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
      // field_1=LabId, field_2=RejectionType, field_3=Reason, field_4=RejectedBy
      await createItem(LISTS.REJECTED, {
        Title:   now,
        field_1: labId,
        field_2: rejectionType,
        field_3: reason.trim(),
        field_4: rejectedBy || 'Lab Staff',
      }).catch(e => context.log('[Rejected] Error:', e.message));
      log.push('✅ Written to Rejected list');

      // ── Write to Activity Log ─────────────────────────────────────────────
      const actNow = new Date();
      const dateStr = now.toLocaleDateString('en-US', { timeZone: 'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
      const timeStr = now.toLocaleTimeString('en-US', { timeZone: 'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
      await createItem(LISTS.ACTIVITY_LOG, {
        Title:        `${actNow.toISOString().split('T')[0]} ${labId}`,
        Date:         dateStr,
        Time:         timeStr,
        Client:       labId,
        ActivityType: rejType,
        Notes:        reason.trim(),
        By:           rejectedBy || 'Lab Staff',
        Qty:          0,
      }).catch(e => context.log('[ActivityLog] Error:', e.message));
      log.push('✅ Written to Activity Log');

      // 2. Update Archived Intake — change suffix to REJ and append notes
      // field_1=fullId, field_2=coaTest, field_13=notes, field_14=status
      const archived = await listItems(LISTS.ARCHIVED_INTAKE, { top: 500 });
      const matches  = archived.filter(r => (r.field_1 || '').startsWith(baseId));

      for (const item of matches) {
        const existingNotes = item.field_13 || '';
        const newNotes = existingNotes ? `${existingNotes} | ${rejNote}` : rejNote;
        // Change full lab ID suffix to REJ (e.g. "072826-003 COMP" → "072826-003 REJ")
        const newFullId = `${baseId} REJ`;
        await updateItem(LISTS.ARCHIVED_INTAKE, item._id, {
          field_1:  newFullId,
          field_2:  rejectionType,
          field_13: newNotes,
          field_14: 'Rejected',
        }).catch(e => context.log('[ArchivedIntake] Error:', e.message));
      }
      log.push(`✅ Archived Intake: updated ${matches.length} row(s) → suffix changed to REJ`);

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, labId, newLabId: `${baseId} REJ`, rejectionType, log }),
      };
    } catch(e) {
      context.log('[reject-sample] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
