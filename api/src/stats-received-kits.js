const { app } = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

// Accession base IDs encode their receipt date as MMDDYY-###.
function accessionDate(baseId) {
  const m = String(baseId || '').trim().match(/^(\\d{2})(\\d{2})(\\d{2})-\\d{3}(?:\\s|$)/);
  if (!m) return '';
  const year = 2000 + Number(m[3]), month = Number(m[1]), day = Number(m[2]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return '';
  return String(year) + '-' + m[1] + '-' + m[2];
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
      const items = await listItems(LISTS.ACCESSION_LOG);
      const ids = new Set();
      for (const row of items) {
        // field_1 stores the base accession ID; field_2 can include a test suffix.
        const rawId = String(row.field_1 || row.field_2 || '').trim();
        const match = rawId.match(/^(\\d{6}-\\d{3})(?:\\s|$)/);
        if (!match) continue;
        const baseId = match[1];
        const received = accessionDate(baseId);
        if (received && received >= start && received <= today) ids.add(baseId);
      }
      return { status: 200, headers: { 'Cache-Control':'no-store' },
        jsonBody: { success:true, count:ids.size, weekStart:start, through:today } };
    } catch (e) {
      context.log('[stats-received-kits]', e.message);
      return { status:500, jsonBody:{ success:false, error:e.message } };
    }
  },
});
