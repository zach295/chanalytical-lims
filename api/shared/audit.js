const { createItem, getListId, graphGet, graphPost } = require('./graph');

function easternStamp() {
  const now = new Date();
  return {
    date: now.toLocaleDateString('en-US', {
      timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: '2-digit',
    }),
    // Always store Activity Log time as 24-hour military time (HH:MM).
    // hourCycle h23 guarantees midnight is 00:xx rather than 24:xx.
    time: (() => {
      const parts = new Intl.DateTimeFormat('en-US', {
        timeZone: 'America/New_York',
        hour: '2-digit',
        minute: '2-digit',
        hourCycle: 'h23',
      }).formatToParts(now);
      const hh = parts.find(p => p.type === 'hour')?.value || '00';
      const mm = parts.find(p => p.type === 'minute')?.value || '00';
      return `${hh}:${mm}`;
    })(),
  };
}

function cleanText(value, max = 3000) {
  return String(value ?? '').trim().slice(0, max);
}

let _activitySchema = null;

async function getActivitySchema(forceRefresh = false) {
  if (_activitySchema && !forceRefresh) return _activitySchema;
  const listId = await getListId('Activity Log');
  const data = await graphGet(`/sites/${process.env.SP_SITE_ID}/lists/${listId}/columns?$select=name,displayName,choice&$top=100`);
  const cols = data.value || [];
  const names = new Set(cols.map(c => c.name));
  const display = new Map(cols.map(c => [c.displayName, c.name]));
  _activitySchema = { listId, names, display };
  return _activitySchema;
}

async function ensureExactActivityTypeColumn(context) {
  let schema = await getActivitySchema();
  if (schema.names.has('ActivityType') || schema.display.has('Activity Type')) return schema;

  try {
    await graphPost(
      `/sites/${process.env.SP_SITE_ID}/lists/${schema.listId}/columns`,
      { name: 'ActivityType', displayName: 'Activity Type', text: { allowMultipleLines: false } }
    );
    schema = await getActivitySchema(true);
    if (context) context.log('[ActivityLog] Added Activity Type column so administrative events retain their exact label.');
  } catch (e) {
    if (context) context.log('[ActivityLog] Could not add Activity Type column:', e.message);
  }
  return schema;
}

function legacyType(activityType) {
  const t = String(activityType || '').toLowerCase();
  if (t.includes('approve') || t.includes('receive') || t.includes('check')) return 'received';
  if (t.includes('reject') || t.includes('correct') || t.includes('adjust') || t.includes('delete') || t.includes('restore')) return 'adjust';
  if (t.includes('send')) return 'sent';
  if (t.includes('initial')) return 'initial';
  if (t.includes('assemble')) return 'assemble';
  return 'adjust';
}

async function writeActivityLog({ labId, type, notes = '', by = 'Lab Staff', quantity = 0, context, preserveType = false } = {}) {
  const id = cleanText(labId, 255);
  const activityType = cleanText(type, 255);
  const actor = cleanText(by, 255) || 'Lab Staff';
  if (!id) return { success: false, error: 'Activity Log requires labId/client' };
  if (!activityType) return { success: false, error: 'Activity Log requires activity type' };

  const stamp = easternStamp();
  try {
    const schema = preserveType ? await ensureExactActivityTypeColumn(context) : await getActivitySchema();
    const fields = {
      Title: `${stamp.date} ${id}`,
      Client: id,
      Notes: cleanText(notes, 3000),
      By: actor,
    };

    if (schema.names.has('ActivityType') || schema.display.has('Activity Type')) {
      fields[schema.names.has('ActivityType') ? 'ActivityType' : schema.display.get('Activity Type')] = activityType;
    } else if (schema.names.has('Type') || schema.display.has('Type')) {
      // Legacy fallback. Administrative delete/restore callers request preserveType,
      // so this path should only be used if the exact-type column could not be created.
      fields[schema.names.has('Type') ? 'Type' : schema.display.get('Type')] = legacyType(activityType);
    }

    if (schema.names.has('LogDate') || schema.display.has('Log Date')) {
      fields[schema.names.has('LogDate') ? 'LogDate' : schema.display.get('Log Date')] = stamp.date;
    } else if (schema.names.has('Date') || schema.display.has('Date')) {
      fields[schema.names.has('Date') ? 'Date' : schema.display.get('Date')] = stamp.date;
    }

    if (schema.names.has('LogTime') || schema.display.has('Log Time')) {
      fields[schema.names.has('LogTime') ? 'LogTime' : schema.display.get('Log Time')] = stamp.time;
    } else if (schema.names.has('Time') || schema.display.has('Time')) {
      fields[schema.names.has('Time') ? 'Time' : schema.display.get('Time')] = stamp.time;
    }

    if (schema.names.has('Quantity') || schema.display.has('Quantity')) {
      fields[schema.names.has('Quantity') ? 'Quantity' : schema.display.get('Quantity')] = Number(quantity) || 0;
    } else if (schema.names.has('Qty') || schema.display.has('Qty')) {
      fields[schema.names.has('Qty') ? 'Qty' : schema.display.get('Qty')] = Number(quantity) || 0;
    }

    await createItem('Activity Log', fields);
    return { success: true, date: stamp.date, time: stamp.time };
  } catch (e) {
    if (context) context.log(`[ActivityLog] ${activityType} for ${id} failed: ${e.message}`);
    return { success: false, error: e.message };
  }
}

module.exports = { writeActivityLog, easternStamp };
