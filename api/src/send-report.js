/**
 * send-report.js — Azure Function
 * Sends the COA PDF + COC scan to the client via Labs@chanalytical.com
 * Requires: Mail.Send application permission on the Azure app registration
 */
const { app }      = require('@azure/functions');
const { getToken } = require('../shared/graph');

const GRAPH       = 'https://graph.microsoft.com/v1.0';
const FROM_EMAIL  = 'Labs@chanalytical.com';
const FROM_NAME   = 'Chanalytical Laboratories';
const PHONE       = '207-747-1815';

// ── Find COC scan in SharePoint Archive (organized by Month/Day) ─────────────
async function findCOCScan(siteId, labId, token) {
  const SCAN_ARCHIVE = process.env.SP_SCAN_ARCHIVE ||
    '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Archived';

  // Helper to search within a specific folder by path
  async function searchInFolder(folderPath) {
    const marker = 'Shared Documents/';
    const idx = folderPath.indexOf(marker);
    const rel = idx >= 0 ? folderPath.slice(idx + marker.length) : folderPath.replace(/^\/+/, '');
    const dp  = rel.split('/').map(s => encodeURIComponent(s)).join('/');
    const res = await fetch(
      `${GRAPH}/sites/${siteId}/drive/root:/${dp}:/children?$select=id,name&$top=100`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) return null;
    const data = await res.json();
    return (data.value || []).find(f =>
      f.name?.toLowerCase().includes(labId.toLowerCase()) &&
      f.name?.toLowerCase().endsWith('.pdf') &&
      !f.name?.toLowerCase().includes('report') &&
      !f.name?.toLowerCase().includes('temp_')
    ) || null;
  }

  try {
    // 1. Look in today's month/day subfolder first (most likely location)
    const now = new Date();
    const monthFolder = now.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    const dayFolder   = String(now.getDate());
    const todayPath   = `${SCAN_ARCHIVE}/${monthFolder}/${dayFolder}`;
    let coc = await searchInFolder(todayPath).catch(() => null);
    if (coc) return coc;

    // 2. Search the full Archive via SharePoint search (finds files in any subfolder)
    const searchRes = await fetch(
      `${GRAPH}/sites/${siteId}/drive/root/search(q='${encodeURIComponent(labId)}')?$select=id,name&$top=20`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!searchRes.ok) return null;
    const results = await searchRes.json();
    return (results.value || []).find(f =>
      f.name?.toLowerCase().includes(labId.toLowerCase()) &&
      f.name?.toLowerCase().endsWith('.pdf') &&
      !f.name?.toLowerCase().includes('report') &&
      !f.name?.toLowerCase().includes('temp_')
    ) || null;
  } catch(e) {
    return null;
  }
}

// ── Download file as base64 ──────────────────────────────────────────────────
async function downloadBase64(siteId, fileId, token) {
  const res = await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${fileId}/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`Download failed: ${res.status}`);
  return Buffer.from(await res.arrayBuffer()).toString('base64');
}

// ── Build email HTML body ────────────────────────────────────────────────────
function buildEmailBody(clientName, labId) {
  return `
<div style="font-family:Arial,sans-serif;font-size:14px;color:#333;line-height:1.6;">
  <p>Dear ${clientName},</p>
  <p>Please find attached your <strong>Certificate of Analysis</strong> and <strong>Chain of Custody</strong>
  for Lab ID <strong>${labId}</strong>.</p>
  <p>If you have any questions regarding your results, please don't hesitate to contact us.</p>
  <br>
  <p style="margin:0;font-weight:bold;">Chanalytical Laboratories, Inc.</p>
  <p style="margin:0;">347 Main St., Unit 1B &nbsp;|&nbsp; Gorham, ME 04038</p>
  <p style="margin:0;">Phone: ${PHONE} &nbsp;|&nbsp; Email: <a href="mailto:${FROM_EMAIL}">${FROM_EMAIL}</a></p>
  <br>
  <p style="font-size:11px;color:#666;">
    Analytical results are generated at the request of and for the exclusive use of the person named on this report.
    Results apply only to samples as submitted.
  </p>
</div>`.trim();
}

// ─────────────────────────────────────────────────────────────────────────────
app.http('send-report', {
  methods:   ['POST'],
  authLevel: 'anonymous',
  handler:   async (request, context) => {
    const siteId = process.env.SP_SITE_ID;
    const body   = await request.json().catch(() => null);
    if (!body) return { status: 400, jsonBody: { error: 'Request body required' } };

    const { pdfBase64, fileName, labId, clientEmail, clientName, overrideEmail } = body;

    if (!pdfBase64)    return { status: 400, jsonBody: { error: 'pdfBase64 required' } };
    if (!labId)        return { status: 400, jsonBody: { error: 'labId required' } };

    const toEmail = overrideEmail || clientEmail;
    if (!toEmail) return { status: 400, jsonBody: { error: 'No client email address on file. Add an email in Clients & Codes or use the Override Email field.' } };

    let token;
    try { token = await getToken(); }
    catch(e) { return { status: 500, jsonBody: { error: 'Auth failed: ' + e.message } }; }

    // Build attachments list
    const attachments = [
      {
        '@odata.type':  '#microsoft.graph.fileAttachment',
        name:           fileName || `${labId} Report.pdf`,
        contentType:    'application/pdf',
        contentBytes:   pdfBase64,
      }
    ];

    // Try to find and attach the COC scan
    const cocFile = await findCOCScan(siteId, labId, token);
    if (cocFile) {
      try {
        const cocB64 = await downloadBase64(siteId, cocFile.id, token);
        attachments.push({
          '@odata.type': '#microsoft.graph.fileAttachment',
          name:          cocFile.name,
          contentType:   'application/pdf',
          contentBytes:  cocB64,
        });
        context.log(`[send-report] COC attached: ${cocFile.name}`);
      } catch(e) {
        context.log(`[send-report] COC attach failed: ${e.message}`);
        // Continue without COC — don't block the send
      }
    } else {
      context.log(`[send-report] No COC found for ${labId}`);
    }

    // Send email via Graph API from Labs@chanalytical.com
    const emailPayload = {
      message: {
        subject: `Certificate of Analysis — ${labId}`,
        body: {
          contentType: 'HTML',
          content:     buildEmailBody(clientName || 'Valued Client', labId),
        },
        from: {
          emailAddress: { address: FROM_EMAIL, name: FROM_NAME },
        },
        toRecipients: [
          { emailAddress: { address: toEmail, name: clientName || toEmail } }
        ],
        ccRecipients: [
          { emailAddress: { address: FROM_EMAIL, name: FROM_NAME } }
        ],
        attachments,
      },
      saveToSentItems: true,
    };

    const sendRes = await fetch(
      `${GRAPH}/users/${FROM_EMAIL}/sendMail`,
      {
        method:  'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body:    JSON.stringify(emailPayload),
      }
    );

    if (!sendRes.ok) {
      const errText = await sendRes.text().catch(() => '');
      context.log('[send-report] Send failed:', sendRes.status, errText.slice(0, 200));
      return { status: 500, jsonBody: {
        error: sendRes.status === 403
          ? 'Permission denied — the Azure app registration needs Mail.Send permission granted for Labs@chanalytical.com'
          : `Email send failed (${sendRes.status})`,
      }};
    }

    context.log(`[send-report] Sent to ${toEmail} with ${attachments.length} attachment(s)`);
    return { status: 200, jsonBody: {
      success:     true,
      sentTo:      toEmail,
      attachments: attachments.map(a => a.name),
      hasCOC:      attachments.length > 1,
    }};
  }
});
