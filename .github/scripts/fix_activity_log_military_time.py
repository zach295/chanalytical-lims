from pathlib import Path

p = Path('admin-dashboard.html')
s = p.read_text()

needle = '''    _spLogRows = (data.items || []).sort((a, b) => {
      const diff = activityTimestamp(b) - activityTimestamp(a);
      if (diff !== 0) return diff;
      // Stable fallback for records with identical/missing timestamps.
      return String(b.title || '').localeCompare(String(a.title || ''));
    });
    filterActivityLog();
  }
'''
replacement = '''    _spLogRows = (data.items || []).sort((a, b) => {
      const diff = activityTimestamp(b) - activityTimestamp(a);
      if (diff !== 0) return diff;
      // Stable fallback for records with identical/missing timestamps.
      return String(b.title || '').localeCompare(String(a.title || ''));
    });
    filterActivityLog();
  }

  // Normalize Activity Log display to military time (HH:MM), including older
  // rows that may have been stored with AM/PM.
  function formatActivityMilitaryTime(value) {
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
if needle not in s:
    raise SystemExit('Activity Log sort block not found')
s = s.replace(needle, replacement, 1)

old = '''        <td style="padding:8px 12px;white-space:nowrap;color:var(--slate);font-size:12px;">${r.time||'—'}</td>'''
new = '''        <td style="padding:8px 12px;white-space:nowrap;color:var(--slate);font-size:12px;font-family:monospace;">${formatActivityMilitaryTime(r.time)}</td>'''
if old not in s:
    raise SystemExit('Activity Log time display cell not found')
s = s.replace(old, new, 1)

p.write_text(s)
