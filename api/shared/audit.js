const { createItem } = require('./graph');

function easternStamp() {
  const now = new Date();
  return {
    date: now.toLocaleDateString('en-US', {
      timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: '2-digit',
    }),
    time: now.toLocaleTimeString('en-US', {
      timeZone: 'America/New_York', hour: '2-digit', minute: '2-digit', hour12: false,
    }),
  };
}

function cleanText(value, max = 3000) {
  return String(value ?? '').trim().slice(0, max);
}

async function writeActivityLog({ labId, type, notes = '', by = 'Lab Staff', quantity = 0, context } = {}) {
  const id = cleanText(labId, 255);
  const activityType = cleanText(type, 255);
  const actor = cleanText(by, 255) || 'Lab Staff';
  if (!id) return { success: false, error: 'Activity Log requires labId/client' };
  if (!activityType) return { success: false, error: 'Activity Log requires activity type' };

  const stamp = easternStamp();
  try {
    await createItem('Activity Log', {
      Title: `${stamp.date} ${id}`,
      Client: id,
      ActivityType: activityType,
      Notes: cleanText(notes, 3000),
      By: actor,
      LogDate: stamp.date,
      LogTime: stamp.time,
      Quantity: Number(quantity) || 0,
    });
    return { success: true, date: stamp.date, time: stamp.time };
  } catch (e) {
    if (context) context.log(`[ActivityLog] ${activityType} for ${id} failed: ${e.message}`);
    return { success: false, error: e.message };
  }
}

module.exports = { writeActivityLog, easternStamp };
