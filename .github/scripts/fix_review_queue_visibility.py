from pathlib import Path

# 1) Successful scans should explicitly carry ReviewStatus.
p = Path('api/src/scan-folder.js')
s = p.read_text()
old = """            Title:            reviewStatus,\n            LabID:            '',"""
new = """            Title:            reviewStatus,\n            ReviewStatus:     reviewStatus,\n            LabID:            '',"""
if old not in s:
    raise SystemExit('scan-folder ReviewStatus anchor not found')
s = s.replace(old, new, 1)
p.write_text(s)

# 2) SharePoint is source of truth. Remove persistent browser-side deleted-scan filtering
# and force every Review Queue load to bypass browser/SWA caches.
p = Path('admin-dashboard.html')
s = p.read_text()
old = """    const res  = await fetch('/api/get-scan-queue');\n    const data = await res.json();\n    // Filter out locally-deleted scans (prevents re-appearance after F5)\n    const deletedScans = getDeletedScans();\n    console.log('[queue] deletedScans:', [...deletedScans], '| pending fileIds:', (data.pending||[]).map(r=>r.fileId));\n    const scans = (data.pending || []).filter(r =>\n      !deletedScans.has(r.fileId) &&\n      !deletedScans.has(r.barcodeId) &&\n      !deletedScans.has(r.baseId)\n    );"""
new = """    const res  = await fetch('/api/get-scan-queue?fresh=' + Date.now(), { cache: 'no-store' });\n    const data = await res.json();\n    if (!res.ok) throw new Error(data.error || `Review Queue load failed (${res.status})`);\n    // SharePoint Review Queue is the source of truth. Do not hide server rows based\n    // on stale browser localStorage from previously discarded scans.\n    const scans = data.pending || [];"""
if old not in s:
    raise SystemExit('dashboard queue filtering anchor not found')
s = s.replace(old, new, 1)

# Avoid permanently wedging _rqLoading if the queue container is unavailable.
old2 = """  const container = document.getElementById('rq-scan-cards');\n  if (!container) return;"""
new2 = """  const container = document.getElementById('rq-scan-cards');\n  if (!container) { _rqLoading = false; return; }"""
if old2 in s:
    s = s.replace(old2, new2, 1)

p.write_text(s)
print('Fixed Review Queue visibility: explicit status, no stale local filter, no-cache loads')
