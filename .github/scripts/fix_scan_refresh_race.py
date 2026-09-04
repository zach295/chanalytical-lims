from pathlib import Path
p=Path('admin-dashboard.html')
s=p.read_text()
old="""        // By the time the exact item is visible, the normal queue reload will render it.
        // If Graph propagation took longer than 10s, still perform one final reload.
        await loadReviewQueue();
        showToast(confirmed ? '✅ Scan added to queue' : '✅ Scan processed — queue refreshed');"""
new="""        // If another Review Queue load is already in progress, loadReviewQueue() would
        // immediately return because of the _rqLoading guard. Wait for that older load
        // to finish first, then do a guaranteed fresh reload after the scanned item is visible.
        let waitForExistingLoad = 0;
        while (_rqLoading && waitForExistingLoad < 40) {
          await new Promise(r => setTimeout(r, 250));
          waitForExistingLoad++;
        }
        await loadReviewQueue();
        showToast(confirmed ? '✅ Scan added to queue' : '✅ Scan processed — queue refreshed');"""
if old not in s: raise SystemExit('refresh anchor not found')
s=s.replace(old,new,1)
p.write_text(s)
print('Fixed Review Queue refresh race with existing load')