from pathlib import Path

p = Path('admin-dashboard.html')
s = p.read_text()

old = """      const res  = await fetch('/api/get-scan-queue');
      const data = await res.json();
      // Filter out locally-deleted scans (prevents re-appearance after F5)
      const deletedScans = getDeletedScans();
      console.log('[queue] deletedScans:', [...deletedScans], '| pending fileIds:', (data.pending||[]).map(r=>r.fileId));
      const scans = (data.pending || []).filter(r =>
        !deletedScans.has(r.fileId) &&
        !deletedScans.has(r.barcodeId) &&
        !deletedScans.has(r.baseId)
      );
"""
new = """      const res  = await fetch('/api/get-scan-queue?fresh=' + Date.now(), { cache: 'no-store' });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Failed to load Review Queue');
      reviewQueueData = data;
      const scans = (data.pending || []).map(r => ({
        ...r,
        services: r.services || (Array.isArray(r.tests) ? r.tests.join('; ') : ''),
        tests: Array.isArray(r.tests) ? r.tests : []
      }));
      reviewQueueData.pending = scans;
"""
if old not in s:
    raise SystemExit('loadReviewQueue block not found')
s = s.replace(old, new, 1)

old = "document.getElementById('rv-tests-list').textContent = record.services || '—';"
new = "document.getElementById('rv-tests-list').textContent = record.services || (record.tests || []).join('; ') || '—';"
if old not in s:
    raise SystemExit('review modal test display line not found')
s = s.replace(old, new, 1)

old = "tests = record.services.split(';').map(s => s.trim()).filter(Boolean);"
new = "tests = Array.isArray(record.tests) && record.tests.length ? record.tests : String(record.services || '').split(';').map(s => s.trim()).filter(Boolean);"
if old not in s:
    raise SystemExit('confirmAccession services parser not found')
s = s.replace(old, new, 1)
p.write_text(s)

p = Path('api/src/get-scan-queue.js')
s = p.read_text()
old = """        tests:            r.TestSelections ? r.TestSelections.split(',').map(t => t.trim()).filter(Boolean) : [],
        confidence:       r.OCRConfidence || 0,
"""
new = """        services:         r.TestSelections || '',
        tests:            r.TestSelections
          ? String(r.TestSelections).split(/\\s*(?:\\||;)\\s*/).map(t => t.trim()).filter(Boolean)
          : [],
        confidence:       r.OCRConfidence || 0,
"""
if old not in s:
    raise SystemExit('get-scan-queue TestSelections mapping not found')
s = s.replace(old, new, 1)
p.write_text(s)
