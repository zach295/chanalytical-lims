/**
 * send-report.js — Azure Function
 * Sends COA PDF to client via labs@chanalytical.com using Gmail API
 * with Google Service Account domain-wide delegation
 */
const { app }      = require('@azure/functions');
const { getToken, createItem } = require('../shared/graph');
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
  return `<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;line-height:1.6;">
    <p>Please find the attached water analysis report.</p>
    <p>Respectfully,<br>
    <strong>Chanalytical Laboratories, Inc.</strong><br>
    347 Main St., Unit 1B<br>
    Gorham, ME 04038<br>
    (207) 747-1815</p>
    <p style="font-size:11px;color:#888;border-top:1px solid #ddd;padding-top:10px;margin-top:20px;">
      CONFIDENTIALITY NOTICE: The contents of this email message and any attachments are intended solely for the addressee(s) and may contain confidential and/or privileged information and may be legally protected from disclosure. If you are not the intended recipient of this message or their agent, or if this message has been addressed to you in error, please immediately alert the sender by reply email and then delete this message and any attachments. If you are not the intended recipient, you are hereby notified that any use, dissemination, copying, or storage of this message or its attachments is strictly prohibited.
    </p>
  </div>`;
}
// ── Find COC scan in SharePoint ───────────────────────────────────────────────
async function findCOCScan(siteId, labId, token) {
  // Search ONLY within the Archive folder — never search globally to avoid
  // accidentally finding files mid-move or in processing folders.
  // We only READ bytes — the original archived file is never modified or deleted.
  const SCAN_ARCHIVE = process.env.SP_SCAN_ARCHIVE ||
    '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Archived';
  try {
    const marker     = 'Shared Documents/';
    const mi         = SCAN_ARCHIVE.indexOf(marker);
    const folderPath = mi >= 0 ? SCAN_ARCHIVE.slice(mi + marker.length) : SCAN_ARCHIVE.replace(/^\/+/, '');

    const isCOC = f =>
      f.name?.toLowerCase().endsWith('.pdf') &&
      !f.name?.toLowerCase().includes('report') &&
      !f.name?.toLowerCase().includes('temp_');

    // Try search terms in order: full labId, then base ID without suffix
    const baseId     = labId.split(' ')[0].trim();
    const searchTerms = [labId, baseId].filter((v, i, a) => a.indexOf(v) === i);

    for (const term of searchTerms) {
      const searchRes = await fetch(
        `${GRAPH}/sites/${siteId}/drive/root:/${encodeURIComponent(folderPath)}:/search(q='${encodeURIComponent(term)}')?$select=id,name,parentReference&$top=20`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!searchRes.ok) continue;
      const { value } = await searchRes.json();
      const match = (value || []).find(f =>
        f.name?.toLowerCase().includes(term.toLowerCase()) && isCOC(f)
      );
      if (match) return match;
    }
    return null;
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

      // Skip email if saveOnly
      if (body.saveOnly) {
        // Just log and return
        const now2     = new Date();
        const logDate2 = now2.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
        const logTime2 = now2.toLocaleTimeString('en-US', { timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
        await createItem('Activity Log', {
          Title: `${logDate2} ${labId}`, Client: labId,
          ActivityType: 'Report Saved',
          Notes: 'PDF downloaded by lab staff',
          By: authorizedBy || 'Lab Staff',
          LogDate: logDate2, LogTime: logTime2, Quantity: 0,
        }).catch(() => {});
        return { status: 200, jsonBody: { success: true, logged: true } };
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

      // Clear Results Cache entry for this lab ID now that report has been sent
      try {
        const graphToken = await getToken();
        const siteId2    = process.env.SP_SITE_ID;
        const baseId2    = String(labId).split(' ')[0].trim();
        const gHdr       = { Authorization: `Bearer ${graphToken}` };
        const GURL       = 'https://graph.microsoft.com/v1.0';

        const rcListRes  = await fetch(`${GURL}/sites/${siteId2}/lists?$select=id,displayName`, { headers: gHdr });
        const rcListId   = ((await rcListRes.json()).value || []).find(l => l.displayName === 'Results Cache')?.id;
        if (rcListId) {
          const rcItemsRes = await fetch(
            `${GURL}/sites/${siteId2}/lists/${rcListId}/items?$expand=fields($select=LabID)&$top=500`,
            { headers: gHdr }
          );
          const rcItem = ((await rcItemsRes.json()).value || [])
            .find(i => String(i.fields?.LabID || '').split(' ')[0].trim() === baseId2);
          if (rcItem) {
            await fetch(
              `${GURL}/sites/${siteId2}/lists/${rcListId}/items/${rcItem.id}`,
              { method: 'DELETE', headers: gHdr }
            );
            context.log(`[send-report] Cleared Results Cache for ${baseId2}`);
          } else {
            context.log(`[send-report] No Results Cache entry found for ${baseId2}`);
          }
        }
      } catch(cacheErr) {
        context.log('[send-report] Results Cache clear error (non-fatal):', cacheErr.message);
      }

      // Log to Activity Log
      try {
        const now     = new Date();
        const logDate = now.toLocaleDateString('en-US', { timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
        const logTime = now.toLocaleTimeString('en-US', { timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
        const action  = body.saveOnly ? 'Report Saved' : 'Report Emailed';
        const details = body.saveOnly
          ? `PDF saved to SharePoint Archive`
          : `Emailed to: ${toEmail} | COC attached: ${coc ? 'Yes' : 'No'} | Report type: ${reportData?.isRW ? 'RW' : reportData?.isFHA ? 'FHA' : 'Standard'}`;
        await createItem('Activity Log', {
          Title:        `${logDate} ${labId}`,
          Client:       labId,
          ActivityType: action,
          Notes:        details,
          By:           body.authorizedBy || 'Lab Staff',
          LogDate:      logDate,
          LogTime:      logTime,
          Quantity:     0,
        }).catch(() => {});
      } catch(e) { context.log('[send-report] ActivityLog error:', e.message); }

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
