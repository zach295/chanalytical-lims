from pathlib import Path

p = Path('api/src/send-report.js')
s = p.read_text()

old_import = "const { getToken, createItem } = require('../shared/graph');"
new_import = "const { getToken, createItem, listItems, LISTS } = require('../shared/graph');"
if old_import not in s:
    raise SystemExit('Expected graph import not found')
s = s.replace(old_import, new_import, 1)

old_block = """      // Build attachments — use pretty filename if provided
      const pdfFileName = bodyFileName || body.fileName || (isRadon ? `${labId} RW Report.pdf` : `${labId} Report.pdf`);
      const attachments = [{
"""
new_block = """      // If this sample is already marked Reported in Archived Intake, this is a revised report.
      // Use the status before this send so the original first report keeps its normal filename.
      let alreadyReported = false;
      try {
        const reportBaseId = String(labId).match(/(\\d{6}-\\d{3})/)?.[1] || String(labId).split(' ')[0].trim();
        if (reportBaseId) {
          const intakeRows = await listItems(LISTS.ARCHIVED_INTAKE, { top: 2000 });
          alreadyReported = intakeRows.some(r =>
            String(r.field_1 || '').split(' ')[0].trim() === reportBaseId &&
            String(r.field_14 || '').trim().toLowerCase() === 'reported'
          );
        }
      } catch (e) {
        context.log('[send-report] Reported-status lookup (non-fatal):', e.message);
      }

      // Build attachments — use pretty filename if provided. Re-reports get "Revised" before .pdf.
      const basePdfFileName = bodyFileName || body.fileName || (isRadon ? `${labId} RW Report.pdf` : `${labId} Report.pdf`);
      const pdfFileName = alreadyReported && !/\\brevised\\b/i.test(basePdfFileName)
        ? (basePdfFileName.toLowerCase().endsWith('.pdf')
            ? `${basePdfFileName.slice(0, -4)} Revised.pdf`
            : `${basePdfFileName} Revised.pdf`)
        : basePdfFileName;
      const attachments = [{
"""
if old_block not in s:
    raise SystemExit('Expected PDF filename block not found')
s = s.replace(old_block, new_block, 1)
p.write_text(s)
