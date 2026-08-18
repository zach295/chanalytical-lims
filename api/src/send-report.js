/**
 * send-report.js — Azure Function
 * Sends COA PDF to client via labs@chanalytical.com using Gmail API
 * with Google Service Account domain-wide delegation
 */
const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');
const crypto       = require('crypto');

const GRAPH      = 'https://graph.microsoft.com/v1.0';
const FROM_EMAIL = 'Labs@chanalytical.com';
const FROM_NAME  = 'Chanalytical Laboratories';
const PHONE      = '207-747-1815';

// ── Get Gmail access token via service account JWT ────────────────────────────
async function getGmailToken() {
  const sa = JSON.parse(process.env.GMAIL_SERVICE_ACCOUNT || '{}');
  if (!sa.private_key) throw new Error('GMAIL_SERVICE_ACCOUNT env var missing or invalid');

  const now    = Math.floor(Date.now() / 1000);
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const claims = Buffer.from(JSON.stringify({
    iss:   sa.client_email,
    sub:   FROM_EMAIL,
    scope: 'https://www.googleapis.com/auth/gmail.send',
    aud:   'https://oauth2.googleapis.com/token',
    exp:   now + 3600,
    iat:   now,
  })).toString('base64url');

  const sign = crypto.createSign('RSA-SHA256');
  sign.update(`${header}.${claims}`);
  const sig = sign.sign(sa.private_key, 'base64url');
  const jwt = `${header}.${claims}.${sig}`;

  const res  = await fetch('https://oauth2.googleapis.com/token', {
    method:  'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body:    `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
  });
  const data = await res.json();
  if (!data.access_token) throw new Error(`Token error: ${JSON.stringify(data)}`);
  return data.access_token;
}

// ── Build MIME email with PDF attachment ──────────────────────────────────────
function buildMimeMessage(toEmail, clientName, labId, htmlBody, attachments, location, isRadon) {
  const boundary = 'coa_boundary_' + Date.now();
  const lines = [
    `From: ${FROM_NAME} <${FROM_EMAIL}>`,
    `To: ${clientName ? `${clientName} <${toEmail}>` : toEmail}`,
    `CC: ${FROM_EMAIL}`,
    `Subject: ${location ? `${location} ${isRadon ? 'RW ' : ''}Lab Report` : `Certificate of Analysis - ${labId}`}`,
    `MIME-Version: 1.0`,
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    '',
    `--${boundary}`,
    `Content-Type: text/html; charset=utf-8`,
    '',
    htmlBody,
    '',
  ];

  for (const att of attachments) {
    lines.push(
      `--${boundary}`,
      `Content-Type: ${att.contentType}; name="${att.name}"`,
      `Content-Disposition: attachment; filename="${att.name}"`,
      `Content-Transfer-Encoding: base64`,
      '',
      att.contentBytes,
      '',
    );
  }
  lines.push(`--${boundary}--`);
  return Buffer.from(lines.join('\r\n')).toString('base64url');
}

// ── Build HTML email body ─────────────────────────────────────────────────────
function buildEmailBody(clientName, labId) {
  return `<!DOCTYPE html>
<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px;margin:0 auto;">
  <div style="background:#1F3864;padding:20px;text-align:center;">
    <h2 style="color:white;margin:0;">Chanalytical Laboratories, Inc.</h2>
    <p style="color:#cde;margin:4px 0 0;">Certificate of Analysis</p>
  </div>
  <div style="padding:24px;">
    <p>Please find the attached Radon Water report.</p>
    <p style="margin:0;">Respectfully,</p>
    <br>
    <p style="margin:0;"><strong>Chanalytical Laboratories, Inc.</strong></p>
    <p style="margin:0;">347 Main St., Unit 1B</p>
    <p style="margin:0;">Gorham, ME 04038</p>
    <p style="margin:0;">(207) 747-1815</p>
  </div>
  <div style="padding:16px 24px;border-top:1px solid #eee;">
    <p style="font-size:10px;color:#aaa;line-height:1.5;margin:0;">
      CONFIDENTIALITY NOTICE: The contents of this email message and any attachments are intended solely for the addressee(s) and may contain confidential and/or privileged information and may be legally protected from disclosure. If you are not the intended recipient of this message or their agent, or if this message has been addressed to you in error, please immediately alert the sender by reply email and then delete this message and any attachments. If you are not the intended recipient, you are hereby notified that any use, dissemination, copying, or storage of this message or its attachments is strictly prohibited.
    </p>
  </div>
</body></html>`;
}

// ── Find COC scan in SharePoint ───────────────────────────────────────────────
async function findCOCScan(siteId, labId, token) {
  const SCAN_ARCHIVE = process.env.SP_SCAN_ARCHIVE ||
    '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Archived';
  try {
    const searchRes = await fetch(
      `${GRAPH}/sites/${siteId}/drive/root/search(q='${encodeURIComponent(labId)}')?$select=id,name&$top=20`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!searchRes.ok) return null;
    const { value } = await searchRes.json();
    return (value || []).find(f =>
      f.name?.toLowerCase().includes(labId.toLowerCase()) &&
      f.name?.toLowerCase().endsWith('.pdf') &&
      !f.name?.toLowerCase().includes('report') &&
      !f.name?.toLowerCase().includes('temp_')
    ) || null;
  } catch (e) { return null; }
}

// ── Main handler ──────────────────────────────────────────────────────────────
app.http('send-report', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body       = await request.json().catch(() => ({}));
      const { labId, pdfBase64, toEmail: overrideEmail2, overrideEmail, clientName: overrideName, location, isRadon } = body;
      const resolvedEmail = overrideEmail || overrideEmail2 || '';
      if (!labId)     return { status: 400, jsonBody: { error: 'labId required' } };
      if (!pdfBase64) return { status: 400, jsonBody: { error: 'pdfBase64 required' } };

      const siteId = process.env.SP_SITE_ID;
      const token  = await getToken();

      // Get client email from Clients list if not overridden
      let toEmail    = resolvedEmail || '';
      let clientName = overrideName  || '';
      if (!toEmail) {
        const clientsRes = await fetch(
          `${GRAPH}/sites/${siteId}/lists/Clients/items?$expand=fields&$top=500`,
          { headers: { Authorization: `Bearer ${token}` } }
        );
        if (clientsRes.ok) {
          const { value } = await clientsRes.json();
          const baseId    = String(labId).match(/(\d{6}-\d{3})/)?.[1] || labId;
          const match     = (value || []).find(i =>
            String(i.fields?.LabID || '').split(' ')[0].trim() === baseId
          );
          if (match) {
            toEmail    = match.fields?.ReportEmail || match.fields?.Email || '';
            clientName = match.fields?.ClientName  || match.fields?.Title || '';
          }
        }
      }

      if (!toEmail) return { status: 400, jsonBody: {
        error: 'No client email address on file. Add an email in Clients & Codes or use the Override Email field.'
      }};

      // Build attachments
      const pdfFileName = `${labId} Report.pdf`;
      const attachments = [{
        name:         pdfFileName,
        contentType:  'application/pdf',
        contentBytes: pdfBase64,
      }];

      // Try to attach COC scan
      const coc = await findCOCScan(siteId, labId, token);
      if (coc) {
        try {
          const cocRes = await fetch(
            `${GRAPH}/sites/${siteId}/drive/items/${coc.id}/content`,
            { headers: { Authorization: `Bearer ${token}` } }
          );
          if (cocRes.ok) {
            const buf = await cocRes.arrayBuffer();
            attachments.push({
              name:         coc.name,
              contentType:  'application/pdf',
              contentBytes: Buffer.from(buf).toString('base64'),
            });
          }
        } catch (e) { context.log('[send-report] COC attach failed:', e.message); }
      }

      // Get Gmail access token
      const gmailToken = await getGmailToken();

      // Build and send MIME message
      const htmlBody = buildEmailBody(clientName || 'Valued Client', labId);
      const raw      = buildMimeMessage(toEmail, clientName, labId, htmlBody, attachments, location, isRadon);

      const sendRes = await fetch(
        `https://gmail.googleapis.com/gmail/v1/users/${encodeURIComponent(FROM_EMAIL)}/messages/send`,
        {
          method:  'POST',
          headers: { Authorization: `Bearer ${gmailToken}`, 'Content-Type': 'application/json' },
          body:    JSON.stringify({ raw }),
        }
      );

      if (!sendRes.ok) {
        const errText = await sendRes.text().catch(() => '');
        context.log('[send-report] Gmail send failed:', sendRes.status, errText.slice(0, 200));
        return { status: 500, jsonBody: { error: `Gmail send failed (${sendRes.status}): ${errText.slice(0, 100)}` } };
      }

      context.log(`[send-report] Sent to ${toEmail} with ${attachments.length} attachment(s)`);
      return { status: 200, jsonBody: {
        success:     true,
        sentTo:      toEmail,
        attachments: attachments.map(a => a.name),
        hasCOC:      attachments.length > 1,
      }};

    } catch (err) {
      context.log('[send-report] Error:', err.message);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
