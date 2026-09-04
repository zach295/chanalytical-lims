from pathlib import Path

p = Path('admin-dashboard.html')
s = p.read_text()

old = """      await refreshClientList(); // ensure emails/abbrevs are fresh before rendering cards
      const res  = await fetch('/api/get-scan-queue?fresh=' + Date.now(), { cache: 'no-store' });
      const data = await res.json();
"""
new = """      // Load the Review Queue first. Client-list refresh is helpful for dropdowns,
      // but it must never block scanned COCs from appearing for review.
      const queuePromise = fetch('/api/get-scan-queue?fresh=' + Date.now(), { cache: 'no-store' });
      const clientsPromise = Promise.resolve()
        .then(() => refreshClientList())
        .catch(e => console.warn('[queue] client refresh failed:', e?.message || e));
      const res = await queuePromise;
      const data = await res.json();
      await clientsPromise;
"""
if old not in s:
    raise SystemExit('target loadReviewQueue prefetch block not found')
s = s.replace(old, new, 1)
p.write_text(s)
