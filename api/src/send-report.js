/**
 * send-report.js — Azure version
 * Emails COA/RW PDF report to client.
 * Uses nodemailer with SMTP (same as before — just removes Google Sheets dependency).
 * Email lookup now comes from SharePoint Archived Intake list.
 *
 * Required env vars: SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS (or GMAIL_USER/GMAIL_APP_PASSWORD)
 *
 * POST { labId, pdfPages: [base64...], reportType, authorizedBy, overrideEmail }
 */
const { app }   = require('@azure/functions');
const nodemailer = require('nodemailer');
const { listItems, updateItem, LISTS } = require('../shared/graph');

function getTransporter() {
  // Support both SMTP env vars and legacy Gmail env vars
  if (process.env.SMTP_HOST) {
    return nodemailer.createTransport({
      host:   process.env.SMTP_HOST,
      port:   parseInt(process.env.SMTP_PORT) || 587,
      secure: process.env.SMTP_SECURE === 'true',
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS,
      },
    });
  }
  // Legacy Gmail fallback
  return nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: process.env.GMAIL_USER,
      pass: process.env.GMAIL_APP_PASSWORD,
    },
  });
}

function getFromAddress() {
  return `"Chanalytical Laboratories" <${process.env.SMTP_USER || process.env.GMAIL_USER}>`;
}

app.http('send-report', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { labId, pdfPages, reportType, authorizedBy, overrideEmail } = await request.json();

      if (!labId || !pdfPages?.length) {
        return { status: 400, jsonBody: { error: 'labId and pdfPages are required.' } };
      }

      // ── Get client email from SharePoint Archived Intake ─────────────────────
      let clientEmail      = overrideEmail || '';
      let propertyAddress  = '';
      let clientName       = '';

      if (!clientEmail) {
        const baseId = String(labId).match(/(\d{6}-\d{3})/)?.[1] || labId;
        const items  = await listItems(LISTS.ARCHIVED_INTAKE, {
          filter: `startswith(fields/FullId,'${baseId}')`,
          top: 1,
        }).catch(() => []);

        if (items.length) {
          const f     = items[0];
          clientEmail = f.Email      || '';
          clientName  = f.ClientName || '';
          propertyAddress = [f.Address, f.City, f.State].filter(Boolean).join(', ');
        }

        // Fallback: look up email from Clients list by customer name
        if (!clientEmail && clientName) {
          const { listItems: li } = require('../shared/graph');
          const clients = await li('Clients', {
            filter: `fields/Title eq '${clientName.replace(/'/g,"''")}'`,
            top: 1,
          }).catch(() => []);
          if (clients.length) clientEmail = clients[0].Email || '';
        }
      }

      if (!clientEmail || !clientEmail.includes('@')) {
        return { status: 400, jsonBody: { error: `No valid email found for Lab ID ${labId}.` } };
      }

      // ── Build email ──────────────────────────────────────────────────────────
      const isRadon    = reportType === 'RW' || labId.toUpperCase().includes(' RW');
      const reportLabel = isRadon ? 'Radon Water Lab Report' : 'Certificate of Analysis';
      const subject     = propertyAddress
        ? `${propertyAddress} — ${reportLabel} (${labId})`
        : `${reportLabel} — Lab ID ${labId}`;

      const bodyHtml = `<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;font-size:14px;color:#222;">
<p>Please find your water quality report attached.</p><br/>
<p>Respectfully,</p>
<p><strong>Chanalytical Laboratories, Inc.</strong><br/>
347 Main St., Unit 1B<br/>Gorham, ME 04038<br/>(207) 747-1815<br/>
<a href="mailto:Labs@chanalytical.com">Labs@chanalytical.com</a></p><br/>
<div style="font-size:11px;color:#888;line-height:1.5;">
<strong>CONFIDENTIALITY NOTICE:</strong> The contents of this email are intended solely for the
addressee(s). If you are not the intended recipient, please notify the sender immediately and
delete this message.
</div></body></html>`;

      // Build attachments from base64 PDF pages
      const attachments = pdfPages.map((b64, i) => ({
        filename: pdfPages.length === 1
          ? `${labId.replace(/\s+/g,'_')}_${isRadon?'RW':'COA'}.pdf`
          : `${labId.replace(/\s+/g,'_')}_${isRadon?'RW':'COA'}_page${i+1}.pdf`,
        content:     Buffer.from(b64, 'base64'),
        contentType: 'application/pdf',
      }));

      const transporter = getTransporter();
      await transporter.sendMail({
        from:    getFromAddress(),
        to:      clientEmail,
        subject,
        html:    bodyHtml,
        attachments,
      });

      context.log(`[send-report] Sent ${labId} to ${clientEmail}`);

      // ── Mark as Reported in Archived Intake ──────────────────────────────────
      try {
        const baseId = String(labId).match(/(\d{6}-\d{3})/)?.[1] || labId;
        const items  = await listItems(LISTS.ARCHIVED_INTAKE, {
          filter: `startswith(fields/FullId,'${baseId}')`,
          top: 10,
        }).catch(() => []);

        const now   = new Date();
        const stamp = `${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}-${String(now.getFullYear()).slice(-2)}`;
        for (const item of items) {
          await updateItem(LISTS.ARCHIVED_INTAKE, item._id, {
            ReportStatus: 'Sent',
            ReportSentBy: authorizedBy || 'Lab Staff',
            ReportSentDate: stamp,
          }).catch(() => {});
        }
      } catch(e) { context.log('[send-report] Mark sent failed (non-fatal):', e.message); }

      return {
        status:   200,
        jsonBody: { success: true, sentTo: clientEmail, subject },
      };

    } catch (err) {
      context.log('[send-report] Error:', err.message);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
