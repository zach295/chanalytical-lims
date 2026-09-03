from pathlib import Path

def replace(path, old, new, label):
    p = Path(path)
    s = p.read_text()
    if old not in s:
        raise SystemExit(f'{label} not found in {path}')
    p.write_text(s.replace(old, new, 1))

# patch-report-cell: shared helper, no silent audit failures, old -> new already captured
replace('api/src/patch-report-cell.js',
"const { getToken, createItem } = require('../shared/graph');\n",
"const { getToken } = require('../shared/graph');\nconst { writeActivityLog } = require('../shared/audit');\n",
'patch-report import')
replace('api/src/patch-report-cell.js',
"""      // ── Log result edit to Activity Log ───────────────────────────────────
      const { labId: logLabId, changedBy } = body;
      if (field === 'value' && logLabId && changedBy) {
        try {
          const now     = new Date();
          const logDate = now.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
          const logTime = now.toLocaleTimeString('en-US', { timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
          await createItem('Activity Log', {
            Title:        `${logDate} ${logLabId}`,
            Client:       logLabId,
            ActivityType: 'Result Edited',
            Notes:        `${paramName} changed from \"${oldValue ?? ''}\" to \"${value}\"`,
            By:           changedBy,
            LogDate:      logDate,
            LogTime:      logTime,
            Quantity:     0,
          }).catch(() => {});
        } catch(e) { context.log('[patch-report-cell] ActivityLog (non-fatal):', e.message); }
      }

      return { status: 200, jsonBody: { success: true, paramName, field, value, newHex, hardnessUpdate } };
""",
"""      // ── Log result edit to Activity Log ───────────────────────────────────
      const { labId: logLabId, changedBy } = body;
      let auditWarning = null;
      if (field === 'value' && logLabId && changedBy) {
        const audit = await writeActivityLog({
          labId: logLabId,
          type: 'Result Edited',
          notes: `${paramName} changed from \"${oldValue ?? ''}\" to \"${value}\"`,
          by: changedBy,
          context,
        });
        if (!audit.success) auditWarning = audit.error || 'Activity Log write failed';
      }

      return { status: 200, jsonBody: { success: true, paramName, field, value, newHex, hardnessUpdate, auditWarning } };
""",
'patch-report activity block')

# send-report: shared helper + revision/file name details + visible audit warning
replace('api/src/send-report.js',
"const { getToken, createItem, listItems, LISTS } = require('../shared/graph');\n",
"const { getToken, createItem, listItems, LISTS } = require('../shared/graph');\nconst { writeActivityLog } = require('../shared/audit');\n",
'send-report import')
replace('api/src/send-report.js',
"""      // Log to Activity Log
      try {
        const now     = new Date();
        const logDate = now.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
        const logTime = now.toLocaleTimeString('en-US', { timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
        const action  = body.saveOnly ? 'Report Saved' : 'Report Emailed';
        const details = body.saveOnly
          ? `PDF saved to SharePoint Archive`
          : `Emailed to: ${toEmail} | COC attached: ${coc ? 'Yes' : 'No'} | Report type: ${body.isRadon ? 'RW' : 'COA'}`;
        await createItem('Activity Log', {
          Title:        `${logDate} ${labId}`,
          Client:       labId,
          ActivityType: action,
          Notes:        details,
          By:           body.authorizedBy || 'Lab Staff',
          LogDate:      logDate,
          LogTime:      logTime,
          Quantity:     0,
        }).catch(() => {});
      } catch(e) { context.log('[send-report] ActivityLog error:', e.message); }

      return { status: 200, jsonBody: {
        success:     true,
        sentTo:      toEmail,
        attachments: attachments.map(a => a.name),
        hasCOC:      attachments.length > 1,
      }};
""",
"""      // Log to Activity Log — include revision state and exact archived filename.
      const revisionLabel = alreadyReported ? 'Revised' : 'Original';
      const details = `Emailed to: ${toEmail} | Revision: ${revisionLabel} | PDF: ${pdfFileName} | COC attached: ${coc ? 'Yes' : 'No'} | Report type: ${body.isRadon ? 'RW' : 'COA'}`;
      const audit = await writeActivityLog({
        labId,
        type: 'Report Emailed',
        notes: details,
        by: body.authorizedBy || 'Lab Staff',
        context,
      });
      const auditWarning = audit.success ? null : (audit.error || 'Activity Log write failed');

      return { status: 200, jsonBody: {
        success:     true,
        sentTo:      toEmail,
        attachments: attachments.map(a => a.name),
        hasCOC:      attachments.length > 1,
        auditWarning,
      }};
""",
'send-report activity block')

# sample correction: preserve previous values in notes instead of only new values
replace('api/src/update-sample.js',
"""        const changes = Object.entries(updates)
          .filter(([,v]) => v !== undefined && v !== '')
          .map(([k,v]) => `${fieldLabels[k] || k} → ${v}`)
          .join('; ');
""",
"""        const firstOld = archivedItems[0] || {};
        const oldValues = {
          coaTest: [...new Set(archivedItems.map(r => r.field_2).filter(Boolean))].join(' | '),
          customer: firstOld.field_3 || '',
          dateDrawn: firstOld.field_4 || '',
          timeDrawn: firstOld.field_5 || '',
          receivedDate: firstOld.field_6 || '',
          receivedTime: firstOld.field_7 || '',
          location: firstOld.field_8 || '',
          city: firstOld.field_9 || '',
          state: firstOld.field_10 || '',
          zip: firstOld.field_11 || '',
          notes: firstOld.field_13 || '',
        };
        const changes = Object.entries(updates)
          .filter(([,v]) => v !== undefined && v !== '')
          .map(([k,v]) => `${fieldLabels[k] || k}: \"${oldValues[k] ?? ''}\" → \"${v}\"`)
          .join('; ');
""",
'update-sample change notes')

# rejection: stable ActivityType; rejection subtype stays in notes
replace('api/src/reject-sample.js',
"""        const fullNotes = [`Reason: ${reason.trim()}`, updateResults && `Updates: ${updateResults}`]
          .filter(Boolean).join('\\n');
        await createItem('Activity Log', {
          Title:        `${dateStr} ${labId}`,
          Client:       labId,
          ActivityType: rejectionType,
""",
"""        const fullNotes = [`Rejection Type: ${rejectionType}`, `Reason: ${reason.trim()}`, updateResults && `Updates: ${updateResults}`]
          .filter(Boolean).join('\\n');
        await createItem('Activity Log', {
          Title:        `${dateStr} ${labId}`,
          Client:       labId,
          ActivityType: 'Sample Rejected',
""",
'reject-sample type')

# inventory activity: server supplies date/time when caller omits them, and actor never blank
replace('api/src/inventory-write.js',
"""      if (action === 'log_activity') {
        const entry = body.payload?.entry || body.entry || body.payload || body;
        await createItem(LISTS.ACTIVITY_LOG, {
          Title:        `${entry.date || ''} ${entry.client || ''}`.trim(),
          LogDate:         entry.date   || '',
          LogTime:         entry.time   || '',
          Client:       entry.client || '',
          ActivityType: entry.type   || '',
          Quantity: Number(entry.qty) || 0,
          Notes:        entry.notes  || '',
          By:           entry.by     || '',
        });
""",
"""      if (action === 'log_activity') {
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
""",
'inventory log timestamp')
