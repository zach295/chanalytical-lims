from pathlib import Path

p = Path('admin-dashboard.html')
s = p.read_text()
old = '''    _spLogRows = (data.items || []).sort((a, b) => {
      const da = (a.date || '') + (a.time || '');
      const db = (b.date || '') + (b.time || '');
      return db.localeCompare(da);
    });
'''
new = '''    // Sort by actual date + time, newest first. Activity Log dates are normally
    // MM/DD/YY, but accept YYYY-MM-DD too so older/newer records sort correctly.
    const activityTimestamp = (row) => {
      const dateRaw = String(row.date || '').trim();
      const timeRaw = String(row.time || '').trim();
      let year = 0, month = 1, day = 1;

      let m = dateRaw.match(/^(\\d{1,2})[\\/-](\\d{1,2})[\\/-](\\d{2}|\\d{4})$/);
      if (m) {
        month = Number(m[1]); day = Number(m[2]);
        year = Number(m[3]);
        if (year < 100) year += 2000;
      } else {
        m = dateRaw.match(/^(\\d{4})-(\\d{1,2})-(\\d{1,2})$/);
        if (m) { year = Number(m[1]); month = Number(m[2]); day = Number(m[3]); }
      }

      let hour = 0, minute = 0, second = 0;
      const tm = timeRaw.match(/^(\\d{1,2}):(\\d{2})(?::(\\d{2}))?\\s*(AM|PM)?$/i);
      if (tm) {
        hour = Number(tm[1]); minute = Number(tm[2]); second = Number(tm[3] || 0);
        const ap = (tm[4] || '').toUpperCase();
        if (ap === 'PM' && hour < 12) hour += 12;
        if (ap === 'AM' && hour === 12) hour = 0;
      }

      if (!year || month < 1 || month > 12 || day < 1 || day > 31) return 0;
      return new Date(year, month - 1, day, hour, minute, second).getTime();
    };

    _spLogRows = (data.items || []).sort((a, b) => {
      const diff = activityTimestamp(b) - activityTimestamp(a);
      if (diff !== 0) return diff;
      // Stable fallback for records with identical/missing timestamps.
      return String(b.title || '').localeCompare(String(a.title || ''));
    });
'''
if old not in s:
    raise SystemExit('Activity Log sort block not found')
p.write_text(s.replace(old, new, 1))
