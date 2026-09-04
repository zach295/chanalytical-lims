from pathlib import Path
p=Path('api/src/scan-folder.js')
s=p.read_text()
old="""          // Move to REVIEW immediately to prevent duplicate processing
          const moveStartedAt = Date.now();
          await moveSpFile(file.id, SCAN_REVIEW, token);
          timing.moveMs = Date.now() - moveStartedAt;

          // Download file as Buffer → base64 for Azure Doc Intel
"""
new="""          // Do NOT move the file before OCR. Some SharePoint/scanner flows can race
          // with an early move and return the file to Incoming. Duplicate processing is
          // already prevented by Review Queue FileID tracking, so process it in place first.
          timing.moveMs = 0;

          // Download file as Buffer → base64 for Azure Doc Intel
"""
if old not in s: raise SystemExit('initial move block not found')
s=s.replace(old,new,1)

old2="""          timing.reviewQueueMs = Date.now() - queueStartedAt;
          timing.totalMs = Date.now() - fileStartedAt;
"""
new2="""          timing.reviewQueueMs = Date.now() - queueStartedAt;

          // Only after the Review Queue row exists, move the physical scan to Review.
          // If SharePoint move/propagation fails, keep the queue row and continue: the
          // FileID remains valid and queuedIds prevents the Incoming copy being reprocessed.
          const moveStartedAt = Date.now();
          try {
            await moveSpFile(file.id, SCAN_REVIEW, token);
          } catch (moveErr) {
            context.log(`[scan] Review move deferred for ${file.name}: ${moveErr.message}`);
          }
          timing.moveMs = Date.now() - moveStartedAt;
          timing.totalMs = Date.now() - fileStartedAt;
"""
if old2 not in s: raise SystemExit('post queue timing block not found')
s=s.replace(old2,new2,1)
p.write_text(s)
print('Changed COC scan to process in Incoming and move only after Review Queue write')
