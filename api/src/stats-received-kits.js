const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

function dateISO(value) {
  const raw = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw.slice(0, 10);
  const m = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2}|\d{4})(?:\s|$)/);
  if (!m) return '';
  const year = m[3].length === 2 ? '20' + m[3] : m[3];
  return year + '-' + m[1].padStart(2, '0') + '-' + m[2].padStart(2, '0');
}

app.http('stats-received-kits', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const now = new Date();
      const parts = new Intl.DateTimeFormat('en-US', {
        timeZone: 'America/New_York', year: 'numeric', month: '2-digit', day: '2-digit', weekday: 'short'
      }).formatToParts(now);
      const get = type => parts.find(p => p.type === type)?.value || '';
      const today = get('year') + '-' + get('month') + '-' + get('day');
      const weekday = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'].indexOf(get('weekday'));
      const monday = new Date(today + 'T12:00:00Z');
      monday.setUTCDate(monday.getUTCDate() - (weekday === 0 ? 6 : weekday - 1));
      const start = monday.toISOString().slice(0, 10);
      const items = await listItems(LISTS.ARCHIVED_INTAKE);
      const ids = new Set();
      for (const row of items) {
        const date = dateISO(row.field_6);
        if (!date || date < start || date > today) continue;
        const labId = String(row.field_1 || '').trim();
        const base = labId.match(/^\d{6}-\d{3}/)?.[0] || labId;
        if (base) ids.add(base);
      }
      return { status: 200, headers: { 'Cache-Control':'no-store' },
        jsonBody: { success:true, count:ids.size, weekStart:start, through:today } };
    } catch (e) {
      context.log('[stats-received-kits]', e.message);
      return { status:500, jsonBody:{ success:false, error:e.message } };
    }
  },
});
