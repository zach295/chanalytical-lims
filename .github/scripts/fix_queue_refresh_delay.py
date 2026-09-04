from pathlib import Path
p=Path('admin-dashboard.html')
s=p.read_text()
old="""      if (found > 0) {
        showToast('✅ Processing scan — checking queue...');
        // Poll get-scan-queue directly (no card rebuilds) until new item appears
        let attempts = 0;
        let prevQueueCount = 0;
        try {
          const initRes = await fetch('/api/get-scan-queue');
          prevQueueCount = ((await initRes.json()).pending || []).length;
        } catch(e) {}

        const poll = async () => {
          attempts++;
          await new Promise(res => setTimeout(res, 2500));
          try {
            const res  = await fetch('/api/get-scan-queue');
            const data = await res.json();
            const curr = (data.pending || []).length;
            if (curr > prevQueueCount || attempts >= 12) {
              // New card detected — do ONE full reload
              await loadReviewQueue();
              showToast(curr > prevQueueCount ? '✅ Scan added to queue' : '✅ Queue refreshed');
              return;
            }
          } catch(e) {}
          return poll();
        };
        await poll();
      } else {"""
new="""      if (found > 0) {
        // scan-folder does not return until the Review Queue item has been written.
        // Reload immediately; the old polling logic captured its baseline AFTER the
        // new item already existed, so it waited the full retry window unnecessarily.
        showToast('✅ Scan processed — loading review queue...');
        await loadReviewQueue();
        showToast('✅ Scan added to queue');
      } else {"""
if old not in s:
    raise SystemExit('target polling block not found')
s=s.replace(old,new,1)
p.write_text(s)
print('Removed unnecessary post-scan queue polling delay')