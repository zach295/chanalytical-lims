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

      // ── Delete temp file unless keepTemp is set ──────────────────────────
      if (!keepTemp) {
        await fetch(
          `${GRAPH}/sites/${siteId}/drive/items/${tempId}`,
          { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } }
        ).catch(e => context.log('[export-pdf] Delete failed (non-fatal):', e.message));
        context.log('[export-pdf] Temp file deleted');
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
