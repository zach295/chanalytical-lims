from pathlib import Path


def rep(path, old, new, label, count=1):
    p = Path(path)
    s = p.read_text()
    if s.count(old) < count:
        raise SystemExit(f'{label} not found enough times in {path}')
    p.write_text(s.replace(old, new, count))

# Dashboard: pass signed-in user with every import.
rep('admin-dashboard.html',
    "body: JSON.stringify({ datePrefix })",
    "body: JSON.stringify({ datePrefix, importedBy: getAdminName() })",
    'pH and bacteria payloads', 2)
rep('admin-dashboard.html',
    "body: JSON.stringify({ dateFilter }),",
    "body: JSON.stringify({ dateFilter, importedBy: getAdminName() }),",
    'radon payload')
rep('admin-dashboard.html',
    "body: JSON.stringify({ all: true, datePrefix: getImportDatePrefix() }),",
    "body: JSON.stringify({ all: true, datePrefix: getImportDatePrefix(), importedBy: getAdminName() }),",
    'generic ICPMS/control/acid payload')
rep('admin-dashboard.html',
    "body: JSON.stringify({ uploadedBy: getAdminName(), datePrefix: getImportDatePrefix() }),",
    "body: JSON.stringify({ uploadedBy: getAdminName(), importedBy: getAdminName(), datePrefix: getImportDatePrefix() }),",
    'control payload')

# Activity Log: normalize date display to MM/DD/YY, including legacy YYYY-MM-DD rows.
needle = '''  function formatActivityMilitaryTime(value) {
    const raw = String(value || '').trim();
    if (!raw) return '—';
    const m = raw.match(/^(\\d{1,2}):(\\d{2})(?::\\d{2})?\\s*(AM|PM)?$/i);
    if (!m) return raw;
    let hour = Number(m[1]);
    const minute = m[2];
    const ap = (m[3] || '').toUpperCase();
    if (ap === 'PM' && hour < 12) hour += 12;
    if (ap === 'AM' && hour === 12) hour = 0;
    if (hour === 24) hour = 0;
    return `${String(hour).padStart(2, '0')}:${minute}`;
  }
'''
replacement = needle + '''\n  function formatActivityDate(value) {
    const raw = String(value || '').trim();
    if (!raw) return '—';
    let m = raw.match(/^(\\d{4})-(\\d{1,2})-(\\d{1,2})$/);
    if (m) return `${String(m[2]).padStart(2,'0')}/${String(m[3]).padStart(2,'0')}/${m[1].slice(-2)}`;
    m = raw.match(/^(\\d{1,2})[\\/-](\\d{1,2})[\\/-](\\d{2}|\\d{4})$/);
    if (m) return `${String(m[1]).padStart(2,'0')}/${String(m[2]).padStart(2,'0')}/${String(m[3]).slice(-2)}`;
    return raw;
  }
'''
rep('admin-dashboard.html', needle, replacement, 'activity date formatter')
rep('admin-dashboard.html',
    "${r.date||'—'}",
    "${formatActivityDate(r.date)}",
    'activity date display')

# Summary import logs: use importer identity instead of System.
for path in ['api/src/import-icpms.js', 'api/src/import-control.js', 'api/src/import-acid.js']:
    p = Path(path)
    s = p.read_text()
    old = "By: 'System', LogDate: _ld, LogTime: _lt, Quantity: 0,"
    if old not in s:
        raise SystemExit(f'import actor log not found in {path}')
    s = s.replace(old, "By: body.importedBy || body.uploadedBy || 'Lab Staff', LogDate: _ld, LogTime: _lt, Quantity: 0,", 1)
    p.write_text(s)

# Inventory activity can receive YYYY-MM-DD from date inputs. Normalize it before storing.
p = Path('api/src/inventory-write.js')
s = p.read_text()
old = """        const logDate = entry.date || serverDate;\n        const logTime = entry.time || serverTime;\n"""
new = """        const rawDate = String(entry.date || serverDate).trim();\n        const isoMatch = rawDate.match(/^(\\d{4})-(\\d{2})-(\\d{2})$/);\n        const usMatch  = rawDate.match(/^(\\d{1,2})[\\/-](\\d{1,2})[\\/-](\\d{2}|\\d{4})$/);\n        const logDate = isoMatch\n          ? `${isoMatch[2]}/${isoMatch[3]}/${isoMatch[1].slice(-2)}`\n          : usMatch\n            ? `${String(usMatch[1]).padStart(2,'0')}/${String(usMatch[2]).padStart(2,'0')}/${String(usMatch[3]).slice(-2)}`\n            : serverDate;\n        const logTime = entry.time || serverTime;\n"""
if old not in s:
    raise SystemExit('inventory activity date block not found')
p.write_text(s.replace(old, new, 1))
