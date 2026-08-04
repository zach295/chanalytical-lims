const { app } = require('@azure/functions');
const { updateItem, deleteItem, getToken, LISTS } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

async function moveSpFile(itemId, destFolderPath, token) {
  const siteId = process.env.SP_SITE_ID;
  const marker = 'Shared Documents/';
  const idx    = destFolderPath.indexOf(marker);
  const rel    = idx >= 0 ? destFolderPath.slice(idx + marker.length) : destFolderPath.replace(/^\/+/,'');
  const drivePath = rel.split('/').map(s => encodeURIComponent(s)).join('/');
  try {
    const folderRes = await fetch(
      `${GRAPH}/sites/${siteId}/drive/root:/${drivePath}?$select=id`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!folderRes.ok) return;
    const destId = (await folderRes.json()).id;
    await fetch(`${GRAPH}/sites/${siteId}/drive/items/${itemId}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ parentReference: { id: destId } }),
    });
  } catch(e) { console.warn('[moveSpFile]', e.message); }
}

async function deleteSpFile(itemId, token) {
  const siteId = process.env.SP_SITE_ID;
  try {
    await fetch(`${GRAPH}/sites/${siteId}/drive/items/${itemId}`, {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${token}` },
    });
  } catch(e) { console.warn('[deleteSpFile]', e.message); }
}

app.http('mark-scan-processed', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { fileId, outcome, reviewQueueRow, rowIndex } = await request.json();
      const row = reviewQueueRow || rowIndex;
      if (!row) return { status: 400, body: JSON.stringify({ error: 'rowIndex required' }) };

      const SCAN_ARCHIVE = process.env.SP_SCAN_ARCHIVE ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Archived';

      if (outcome === 'discarded') {
        await deleteItem(LISTS.REVIEW_QUEUE, row).catch(() => {});
        if (fileId) {
          const token = await getToken();
          await deleteSpFile(fileId, token);
        }
      } else {
        await updateItem(LISTS.REVIEW_QUEUE, row, { Title: 'Processed' }).catch(() => {});
        if (fileId) {
          const token = await getToken();
          await moveSpFile(fileId, SCAN_ARCHIVE, token);
        }
      }

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, driveDeleted: !!fileId }),
      };
    } catch(e) {
      context.log('[mark-scan-processed] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
