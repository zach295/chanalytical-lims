const { app } = require('@azure/functions');
const { getToken } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';

async function deleteSpFile(siteId, itemId, token) {
  if (!itemId) return;
  await fetch(`${GRAPH}/sites/${siteId}/drive/items/${itemId}`, {
    method: 'DELETE',
    headers: { Authorization: `Bearer ${token}` },
  }).catch(e => console.warn('[deleteSpFile]', e.message));
}

async function getListId(siteId, displayName, token) {
  const res  = await fetch(`${GRAPH}/sites/${siteId}/lists?$select=id,displayName`,
    { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) return null;
  return ((await res.json()).value || []).find(l => l.displayName === displayName)?.id || null;
}

app.http('mark-scan-processed', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { fileId, outcome, reviewQueueRow, rowIndex } = await request.json();
      const row    = reviewQueueRow || rowIndex;
      const token  = await getToken();
      const siteId = process.env.SP_SITE_ID;

      if (outcome === 'discarded') {
        // 1. Delete the Review Queue list entry
        if (row) {
          const rqListId = await getListId(siteId, 'Review Queue', token);
          if (rqListId) {
            await fetch(`${GRAPH}/sites/${siteId}/lists/${rqListId}/items/${row}`, {
              method: 'DELETE',
              headers: { Authorization: `Bearer ${token}` },
            }).catch(e => context.log('[mark-scan] RQ delete failed:', e.message));
          }
        }

        // 2. Delete the PDF file from SharePoint Drive
        if (fileId) {
          await deleteSpFile(siteId, fileId, token);
          context.log('[mark-scan] Deleted PDF:', fileId);
        }

        // 3. If file is still in Incoming (not yet moved), find and delete it there too
        // (handles case where scan-folder moved it before we got the ID)
        return { status: 200, jsonBody: { success: true, driveDeleted: !!fileId } };

      } else {
        // Mark as processed in Review Queue list
        if (row) {
          const rqListId = await getListId(siteId, 'Review Queue', token);
          if (rqListId) {
            await fetch(`${GRAPH}/sites/${siteId}/lists/${rqListId}/items/${row}/fields`, {
              method: 'PATCH',
              headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
              body: JSON.stringify({ Title: 'Processed', ReviewStatus: 'Processed' }),
            }).catch(e => context.log('[mark-scan] RQ update failed:', e.message));
          }
        }
        return { status: 200, jsonBody: { success: true } };
      }

    } catch(e) {
      context.log('[mark-scan-processed] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});
