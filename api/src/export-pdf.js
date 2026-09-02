/**
 * export-pdf.js — Phase 3
 * Converts an existing filled Excel file (created by prepare-report.js)
 * to PDF and returns it as base64.
 *
 * POST body:
 *   { tempId, labId, isRadon, keepTemp }
 *   tempId    — SharePoint Drive item ID of the filled Excel file
 *   labId     — used to name the PDF file
 *   isRadon   — true for radon reports
 *   keepTemp  — if true, keep the Excel file after export (default: false = delete)
 */

const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');
const GRAPH        = 'https://graph.microsoft.com/v1.0';

app.http('export-pdf', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {
    try {
      const body = await request.json().catch(() => ({}));
      const { tempId, labId, isRadon, keepTemp = false } = body;

      if (!tempId) return { status: 400, jsonBody: { error: 'tempId required' } };

      // cleanupOnly: just delete the temp file, no PDF conversion
      if (body.cleanupOnly) {
        const siteId2 = process.env.SP_SITE_ID;
        const token2  = await getToken();
        // Just delete on cleanup — orphaned sessions don't need archiving
        await fetch(`${GRAPH}/sites/${siteId2}/drive/items/${tempId}`,
          { method: 'DELETE', headers: { Authorization: `Bearer ${token2}` } }).catch(() => {});
        context.log('[export-pdf] Cleanup only — deleted tempId:', tempId);
        return { status: 200, jsonBody: { success: true, deleted: true } };
      }

      if (!labId)  return { status: 400, jsonBody: { error: 'labId required'  } };

      const siteId = process.env.SP_SITE_ID;
      const token  = await getToken();

      context.log(`[export-pdf] Exporting ${labId} from tempId ${tempId}`);

      // ── Export to PDF ────────────────────────────────────────────────────
      // Small wait to ensure any pending writes are flushed before export
      await new Promise(r => setTimeout(r, 1000));

      const pr = await fetch(
        `${GRAPH}/sites/${siteId}/drive/items/${tempId}/content?format=pdf`,
        { headers: { Authorization: `Bearer ${token}` } }
      );

      if (!pr.ok) {
        const errText = await pr.text().catch(() => '');
        context.log(`[export-pdf] PDF export failed: ${pr.status} ${errText.slice(0, 200)}`);
        return { status: 500, jsonBody: { error: `PDF export failed (${pr.status})` } };
      }

      const pdfBase64 = Buffer.from(await pr.arrayBuffer()).toString('base64');
      context.log(`[export-pdf] PDF size: ${pdfBase64.length} chars`);

      // ── Delete temp Excel and save PDF copy to Archive ───────────────────────
      if (!keepTemp) {
        // Delete the temp Excel file
        await fetch(`${GRAPH}/sites/${siteId}/drive/items/${tempId}`,
          { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } }
        ).catch(e => context.log('[export-pdf] Delete failed (non-fatal):', e.message));
        context.log('[export-pdf] Temp Excel deleted');

        // Save PDF copy to Archive folder
        try {
          const SCAN_ARCHIVE = process.env.SP_SCAN_ARCHIVE ||
            '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Archived';
          const archiveMarker = 'Shared Documents/';
          const archiveIdx    = SCAN_ARCHIVE.indexOf(archiveMarker);
          const archiveRel    = archiveIdx >= 0
            ? SCAN_ARCHIVE.slice(archiveIdx + archiveMarker.length)
            : SCAN_ARCHIVE.replace(/^\/+/, '');
          const pdfName    = isRadon ? `${labId} RW Report.pdf` : `${labId} Report.pdf`;
          const uploadPath = `${archiveRel}/${encodeURIComponent(pdfName)}`;
          const pdfBytes   = Buffer.from(pdfBase64, 'base64');
          await fetch(
            `${GRAPH}/sites/${siteId}/drive/root:/${uploadPath}:/content`,
            { method: 'PUT', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/pdf' },
              body: pdfBytes }
          ).catch(e => context.log('[export-pdf] PDF archive save failed (non-fatal):', e.message));
          context.log('[export-pdf] PDF copy saved to Archive:', pdfName);
        } catch(e) { context.log('[export-pdf] PDF archive error (non-fatal):', e.message); }
      } else {
        context.log('[export-pdf] Temp file kept (keepTemp=true)');
      }

      const fileName = isRadon ? `${labId} RW Report.pdf` : `${labId} Report.pdf`;

      return {
        status:   200,
        jsonBody: { success: true, pdfBase64, fileName },
      };

    } catch(e) {
      context.log('[export-pdf] Error:', e.message);
      return { status: 500, jsonBody: { error: e.message } };
    }
  },
});
