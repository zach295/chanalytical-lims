/**
 * sync-to-sheets.js — Azure version
 * Receives field app (index.html) sample submissions and writes to
 * SharePoint Archived Intake list. Replaces Google Sheets append.
 *
 * POST { id, dateDrawn, timeDrawn, customer, email, phone,
 *        location, city, state, zip, services, notes, clientKey }
 */
const { app }        = require('@azure/functions');
const { createItem, LISTS } = require('../shared/graph');

function formatDate(dateStr) {
  if (!dateStr) return '';
  const [y, m, d] = dateStr.split('-');
  return `${m}-${d}-${y.slice(-2)}`;
}

function formatTime(timeStr) {
  if (!timeStr) return '';
  return String(timeStr).substring(0, 5);
}

app.http('sync-to-sheets', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const submission = await request.json();

      const {
        id, dateDrawn, timeDrawn, customer, email, phone,
        location, city, state, zip, services, notes, clientKey,
      } = submission;

      if (!id) return { status: 400, jsonBody: { error: 'id required' } };

      const ts = new Date().toISOString();

      await createItem(LISTS.ARCHIVED_INTAKE, {
        Title:        id,
        Timestamp:    ts,
        FullId:       id,
        CoaTest:      services || '',
        ClientName:   customer || '',
        ClientCode:   clientKey || '',
        DateDrawn:    formatDate(dateDrawn),
        TimeDrawn:    formatTime(timeDrawn),
        ReceivedDate: '',
        ReceivedTime: '',
        Address:      location || '',
        City:         city     || '',
        State:        state    || 'ME',
        Zip:          zip      ? String(zip).padStart(5,'0') : '',
        Notes:        notes    || '',
        ReportStatus: 'Pre-registered',
        Email:        email    || '',
        Phone:        phone    || '',
      });

      context.log(`[sync-to-sheets] Registered ${id} for ${customer}`);

      return {
        status:   200,
        jsonBody: { success: true, id },
      };

    } catch (err) {
      context.log('[sync-to-sheets] Error:', err.message);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
