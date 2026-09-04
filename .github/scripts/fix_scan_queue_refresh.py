from pathlib import Path

# Backend: return the exact SharePoint file id with each successful scan result.
p=Path('api/src/scan-folder.js')
s=p.read_text()
old="""          results.push({
            fileName:     file.name,
            barcodeId:    ocr.barcodeId || '',"""
new="""          results.push({
            fileId:       file.id,
            fileName:     file.name,
            barcodeId:    ocr.barcodeId || '',"""
if old not in s: raise SystemExit('scan-folder result anchor not found')
s=s.replace(old,new,1)
p.write_text(s)

# Frontend: wait for the exact scanned file to become readable in Review Queue,
# then render from that confirmed queue response instead of relying on timing/counts.
p=Path('admin-dashboard.html')
s=p.read_text()
old="""      if (found > 0) {
        // scan-folder does not return until the Review Queue item has been written.
        // Reload immediately; the old polling logic captured its baseline AFTER the
        // new item already existed, so it waited the full retry window unnecessarily.
        showToast('✅ Scan processed — loading review queue...');
        await loadReviewQueue();
        showToast('✅ Scan added to queue');
      } else {"""
new="""      if (found > 0) {
        const expectedIds = (data.results || []).filter(r => !r.error && r.fileId).map(r => String(r.fileId));
        showToast('✅ Scan processed — loading review queue...');

        // SharePoint list reads can lag briefly after a successful create. Poll for
        // the exact file we just scanned (not a count baseline) so we neither refresh
        // too early nor sit through a fixed retry window after the item is already there.
        let confirmed = null;
        for (let attempt = 0; attempt < 10; attempt++) {
          try {
            const qRes = await fetch('/api/get-scan-queue?fresh=' + Date.now(), { cache: 'no-store' });
            const qData = await qRes.json();
            const pendingIds = new Set((qData.pending || []).map(r => String(r.fileId || '')));
            if (!expectedIds.length || expectedIds.some(id => pendingIds.has(id))) {
              confirmed = qData;
              break;
            }
          } catch(e) {}
          await new Promise(r => setTimeout(r, 1000));
        }

        // By the time the exact item is visible, the normal queue reload will render it.
        // If Graph propagation took longer than 10s, still perform one final reload.
        await loadReviewQueue();
        showToast(confirmed ? '✅ Scan added to queue' : '✅ Scan processed — queue refreshed');
      } else {"""
if old not in s: raise SystemExit('triggerScanNow refresh block not found')
s=s.replace(old,new,1)
p.write_text(s)
print('Added exact-file Review Queue refresh after COC scan')