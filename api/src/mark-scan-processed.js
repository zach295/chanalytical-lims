const { app } = require('@azure/functions');
const { updateItem, deleteItem, findItem, listItems, getToken, LISTS } = require('../shared/graph');
const { writeActivityLog } = require('../shared/audit');

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

  // ══════════════════════════════════════════════════════════════════════════
  // ABSOLUTE ARCHIVE PROTECTION — HARDCODED — NEVER REMOVE OR BYPASS
  // Files in the Archive folder CANNOT be deleted under ANY circumstances.
  // This check runs BEFORE any delete attempt and throws if violated.
  // ══════════════════════════════════════════════════════════════════════════
  try {
    const metaRes = await fetch(
      `${GRAPH}/sites/${siteId}/drive/items/${itemId}?$select=id,name,parentReference`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (metaRes.ok) {
      const meta       = await metaRes.json();
      const parentPath = (meta.parentReference?.path || '').toLowerCase();
      const itemName   = (meta.name || '');

      // HARD CHECK: anything already in Archive is immutable.
      // A COC that is still in the Review folder MUST remain deletable from the
      // Review Queue's "Delete Scan" action; filename alone is not an archive signal.
      const archiveKeywords = ['archived', 'archive', 'lab scans/arch'];
      const inArchive = archiveKeywords.some(kw => parentPath.includes(kw));

      if (inArchive) {
        const msg = `HARD BLOCK: Cannot delete "${itemName}" — it is in the Archive. Path: ${parentPath}`;
        console.error(`[deleteSpFile] ${msg}`);
        throw new Error(msg);
      }
    }
  } catch(guardErr) {
    // Re-throw if it's our own guard error
    if (guardErr.message.startsWith('HARD BLOCK')) throw guardErr;
    // If metadata fetch failed, REFUSE to delete (fail safe)
    console.error(`[deleteSpFile] Could not verify file location for ${itemId} — delete REFUSED for safety`);
    throw new Error(`Delete refused: could not verify file ${itemId} is not in Archive`);
  }
  // ══════════════════════════════════════════════════════════════════════════

  try {
    const res = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${itemId}`, {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok && res.status !== 404) {
      const txt = await res.text().catch(() => '');
      throw new Error(`File delete failed (${res.status}): ${txt.slice(0, 180)}`);
    }
    console.log(`[deleteSpFile] Deleted/already absent ${itemId} (status ${res.status})`);
    return true;
  } catch(e) { console.warn('[deleteSpFile]', e.message); throw e; }
}

app.http('mark-scan-processed', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { fileId, outcome, reviewQueueRow, rowIndex, processedBy, labId, fileName } = await request.json();
      const row = reviewQueueRow || rowIndex;
      if (!row) return { status: 400, body: JSON.stringify({ error: 'rowIndex required' }) };

      const SCAN_ARCHIVE = process.env.SP_SCAN_ARCHIVE ||
        '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Archived';

      let queueDeleted = false;
      let driveDeleted = false;
      if (outcome === 'discarded') {
        // A FileID can have more than one Review Queue row after retries/races.
        // Delete/mark ALL matching rows so an older duplicate cannot reappear on refresh.
        let matchingRows = [];
        if (fileId) {
          try {
            // Do not rely on a SharePoint OData filter here. FileID filtering has
            // intermittently failed because of internal-column naming/propagation.
            // Read the queue and match the stable drive item ID in JavaScript.
            const allQueueRows = await listItems(LISTS.REVIEW_QUEUE);
            matchingRows = allQueueRows.filter(item =>
              String(item.FileID || item.FileId || '').trim() === String(fileId).trim()
            );
            context.log(`[mark-scan-processed] Found ${matchingRows.length} queue row(s) for FileID ${fileId}`);
          } catch (lookupErr) {
            context.log(`[mark-scan-processed] FileID lookup failed: ${lookupErr.message}`);
          }
        }
        // rowIndex is only a fallback. It can be stale after a live refresh, so a
        // 404 on this row must be treated as already removed rather than fatal.
        if (!matchingRows.length && row) matchingRows = [{ _id: row }];

        let queueWarning = null;
        let cleanedRows = 0;
        for (const match of matchingRows) {
          const targetRow = match._id;
          if (!targetRow) continue;

          try {
            await updateItem(LISTS.REVIEW_QUEUE, targetRow, {
              ReviewStatus: 'Discarded',
              Title: 'Discarded',
            });
            context.log(`[mark-scan-processed] Marked Review Queue item ${targetRow} discarded`);
          } catch (statusErr) {
            context.log(`[mark-scan-processed] Status update for ${targetRow}: ${statusErr.message}`);
          }

          try {
            await deleteItem(LISTS.REVIEW_QUEUE, targetRow);
            cleanedRows++;
            context.log(`[mark-scan-processed] Deleted Review Queue item ${targetRow}`);
          } catch (deleteErr) {
            if (/404/.test(deleteErr.message || '')) {
              cleanedRows++;
              context.log(`[mark-scan-processed] Review Queue item ${targetRow} already absent`);
            } else {
              queueWarning = deleteErr.message;
              context.log(`[mark-scan-processed] Review Queue delete deferred for ${targetRow}: ${deleteErr.message}`);
            }
          }
        }

        // Re-query by FileID and make sure no live duplicate can render again.
        if (fileId) {
          try {
            const allRemaining = await listItems(LISTS.REVIEW_QUEUE);
            const remaining = allRemaining.filter(item =>
              String(item.FileID || item.FileId || '').trim() === String(fileId).trim()
            );
            for (const rem of remaining) {
              const status = String(rem.ReviewStatus || rem.Title || '').toLowerCase();
              if (status !== 'discarded') {
                await updateItem(LISTS.REVIEW_QUEUE, rem._id, {
                  ReviewStatus: 'Discarded',
                  Title: 'Discarded',
                }).catch(e => context.log(`[mark-scan-processed] Final discard mark ${rem._id}: ${e.message}`));
              }
            }
            const verifyRows = await listItems(LISTS.REVIEW_QUEUE).catch(() => allRemaining);
            const finalRows = verifyRows.filter(item =>
              String(item.FileID || item.FileId || '').trim() === String(fileId).trim()
            );
            queueDeleted = !finalRows.length || finalRows.every(rem =>
              String(rem.ReviewStatus || rem.Title || '').toLowerCase() === 'discarded'
            );
          } catch (verifyErr) {
            context.log(`[mark-scan-processed] Queue cleanup verification: ${verifyErr.message}`);
            queueDeleted = cleanedRows > 0 && !queueWarning;
          }
        } else {
          queueDeleted = cleanedRows > 0 && !queueWarning;
        }

        // PDF cleanup is independent from list-row cleanup.
        if (fileId) {
          try {
            const token = await getToken();
            context.log(`[mark-scan-processed] Deleting file ${fileId}`);
            driveDeleted = await deleteSpFile(fileId, token);
          } catch (fileErr) {
            context.log(`[mark-scan-processed] PDF delete failed: ${fileErr.message}`);
            queueWarning = queueWarning || `PDF delete failed: ${fileErr.message}`;
          }
        } else {
          context.log('[mark-scan-processed] No fileId to delete');
        }

        if (queueWarning) context.log(`[mark-scan-processed] Cleanup warning: ${queueWarning}`);
      } else {
        await deleteItem(LISTS.REVIEW_QUEUE, row).catch(() => {});
        if (fileId) {
          const token = await getToken();
          await moveSpFile(fileId, SCAN_ARCHIVE, token);
        }
      }

      let auditWarning = null;
      if (outcome === 'discarded') {
        const audit = await writeActivityLog({
          labId: labId || fileName || `Scan ${row}`,
          type: 'Scan Discarded',
          notes: `Review Queue row ${row} discarded${fileName ? ` | File: ${fileName}` : ''}${fileId ? ` | File ID: ${fileId}` : ''}${fileId ? ' | Underlying file deletion requested' : ''}`,
          by: processedBy || 'Lab Staff',
          context,
        });
        if (!audit.success) auditWarning = audit.error;
      }

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, queueDeleted, driveDeleted, row, outcome, auditWarning }),
      };
    } catch(e) {
      context.log('[mark-scan-processed] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
