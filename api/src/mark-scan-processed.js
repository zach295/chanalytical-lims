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
  if (!itemId) { console.warn('[deleteSpFile] No itemId provided'); return; }
  try {
    // ── ABSOLUTE PROTECTION: Never delete files in the Archive ──────────────
    const metaRes = await fetch(
      `${GRAPH}/sites/${siteId}/drive/items/${itemId}?$select=id,name,parentReference`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (metaRes.ok) {
      const meta = await metaRes.json();
      const parentPath = (meta.parentReference?.path || '').toLowerCase();
      if (parentPath.includes('archived') || parentPath.includes('archive')) {
        const err = `BLOCKED DELETE: file "${meta.name}" (${itemId}) is in Archive — deletion is forbidden`;
        console.error(`[deleteSpFile] ${err}`);
        throw new Error(err);
      }
    }
    // ── Proceed with delete only if file is NOT in Archive ──────────────────
    const res = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${itemId}`, {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok && res.status !== 404) {
      const txt = await res.text().catch(() => '');
      console.warn(`[deleteSpFile] Failed ${res.status}: ${txt.slice(0, 100)}`);
    } else {
      console.log(`[deleteSpFile] Deleted ${itemId} (status ${res.status})`);
    }
  } catch(e) { console.warn('[deleteSpFile]', e.message); throw e; }
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
        try {
          // Mark as Processed so get-scan-queue filters it out (same as approve)
          await updateItem(LISTS.REVIEW_QUEUE, row, { Title: 'Processed' });
          context.log(`[mark-scan-processed] Marked discarded item ${row} as Processed`);
        } catch(deleteErr) {
          context.log(`[mark-scan-processed] Update failed for row ${row}:`, deleteErr.message);
        }
        if (fileId) {
          const token = await getToken();
          context.log(`[mark-scan-processed] Deleting file ${fileId}`);
          await deleteSpFile(fileId, token);
        } else {
          context.log('[mark-scan-processed] No fileId to delete');
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
        body: JSON.stringify({ success: true, driveDeleted: !!fileId, row, outcome }),
      };
    } catch(e) {
      context.log('[mark-scan-processed] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
