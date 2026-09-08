from pathlib import Path

p = Path('api/src/send-report.js')
s = p.read_text()

helper_marker = "// ── Build MIME email with PDF attachment ──────────────────────────────────────\n"
helper = r'''// ── Get Google Sheets access token via service account JWT ─────────────────────
async function getSheetsToken() {
  const sa = JSON.parse(process.env.GMAIL_SERVICE_ACCOUNT || '{}');
  if (!sa.private_key) throw new Error('GMAIL_SERVICE_ACCOUNT env var missing or invalid');

  const now    = Math.floor(Date.now() / 1000);
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const claims = Buffer.from(JSON.stringify({
    iss:   sa.client_email,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud:   'https://oauth2.googleapis.com/token',
    exp:   now + 3600,
    iat:   now,
  })).toString('base64url');

  const sign = crypto.createSign('RSA-SHA256');
  sign.update(`${header}.${claims}`);
  const sig = sign.sign(sa.private_key, 'base64url');
  const jwt = `${header}.${claims}.${sig}`;

  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
  });
  const data = await res.json();
  if (!data.access_token) throw new Error(`Sheets token error: ${JSON.stringify(data)}`);
  return data.access_token;
}

'''
if 'async function getSheetsToken()' not in s:
    if helper_marker not in s:
        raise SystemExit('helper marker not found')
    s = s.replace(helper_marker, helper + helper_marker, 1)

insert_marker = "      // Results Cache kept as permanent record for report regeneration\n"
coa_block = r'''      // Write the actual sent Report Date to COA/Form Responses column G.
      // Column H stores the base Lab ID; update every matching row because one sample
      // can have multiple COA rows (package + separately ordered elements).
      let coaReportDateWarning = null;
      try {
        const COA_SHEETS_ID = '15403E6ZaZFQuKNTtgJcb6-2jmlLnk02eq04BGPATqTw';
        const COA_TAB = 'Form Responses';
        const reportBaseId5 = String(labId).match(/(\d{6}-\d{3})/)?.[1] || String(labId).split(' ')[0].trim();
        const now5 = new Date();
        const etNow5 = new Date(now5.toLocaleString('en-US', { timeZone: 'America/New_York' }));
        const pad5 = n => String(n).padStart(2, '0');
        const reportDate5 = `${pad5(etNow5.getMonth()+1)}/${pad5(etNow5.getDate())}/${String(etNow5.getFullYear()).slice(-2)}`;
        const sheetsToken5 = await getSheetsToken();

        // Read H (Lab ID) so we can identify every COA row for this sample.
        const labRange5 = encodeURIComponent(`${COA_TAB}!H1:H`);
        const labRes5 = await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${COA_SHEETS_ID}/values/${labRange5}`,
          { headers: { Authorization: `Bearer ${sheetsToken5}` } }
        );
        if (!labRes5.ok) {
          const t5 = await labRes5.text().catch(()=> '');
          throw new Error(`COA Lab ID lookup failed (${labRes5.status}): ${t5.slice(0,160)}`);
        }
        const labVals5 = (await labRes5.json()).values || [];
        const matchRows5 = [];
        for (let i = 0; i < labVals5.length; i++) {
          const rowLab = String(labVals5[i]?.[0] || '').match(/(\d{6}-\d{3})/)?.[1] || String(labVals5[i]?.[0] || '').split(' ')[0].trim();
          if (rowLab === reportBaseId5) matchRows5.push(i + 1);
        }
        context.log(`[COA Report Date] ${reportBaseId5} — found ${matchRows5.length} Form Responses row(s)`);
        if (!matchRows5.length) throw new Error(`No COA/Form Responses rows found for ${reportBaseId5}`);

        // Batch update column G on every matching row.
        const updateData5 = matchRows5.map(row => ({ range: `${COA_TAB}!G${row}`, values: [[reportDate5]] }));
        const updateRes5 = await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${COA_SHEETS_ID}/values:batchUpdate`,
          {
            method: 'POST',
            headers: { Authorization: `Bearer ${sheetsToken5}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data: updateData5 }),
          }
        );
        if (!updateRes5.ok) {
          const t5 = await updateRes5.text().catch(()=> '');
          throw new Error(`COA Report Date write failed (${updateRes5.status}): ${t5.slice(0,160)}`);
        }

        // Verify G:H for the affected row span instead of assuming the batch write stuck.
        const firstRow5 = Math.min(...matchRows5);
        const lastRow5  = Math.max(...matchRows5);
        const verifyRange5 = encodeURIComponent(`${COA_TAB}!G${firstRow5}:H${lastRow5}`);
        const verifyRes5 = await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${COA_SHEETS_ID}/values/${verifyRange5}`,
          { headers: { Authorization: `Bearer ${sheetsToken5}` } }
        );
        if (!verifyRes5.ok) throw new Error(`COA Report Date verification read failed (${verifyRes5.status})`);
        const verifyVals5 = (await verifyRes5.json()).values || [];
        const missingRows5 = [];
        for (const row of matchRows5) {
          const rel = row - firstRow5;
          const gVal = String(verifyVals5[rel]?.[0] || '').trim();
          const hVal = String(verifyVals5[rel]?.[1] || '').trim();
          const hBase = hVal.match(/(\d{6}-\d{3})/)?.[1] || hVal.split(' ')[0].trim();
          if (hBase !== reportBaseId5 || gVal !== reportDate5) missingRows5.push(row);
        }
        if (missingRows5.length) throw new Error(`COA Report Date verification failed on row(s): ${missingRows5.join(', ')}`);

        context.log(`[COA Report Date] Updated and verified column G on row(s) ${matchRows5.join(', ')} to ${reportDate5}`);
        await writeActivityLog({
          labId: reportBaseId5,
          type: 'COA Report Date Updated',
          notes: `Report Date ${reportDate5} written to Form Responses column G on ${matchRows5.length} row(s): ${matchRows5.join(', ')}`,
          by: body.authorizedBy || 'Lab Staff',
          context,
        }).catch(() => {});
      } catch(e) {
        coaReportDateWarning = e.message;
        context.log('[COA Report Date] Warning:', e.message);
        await writeActivityLog({
          labId,
          type: 'COA Report Date Warning',
          notes: `Report sent, but COA/Form Responses column G Report Date update failed: ${e.message}`,
          by: body.authorizedBy || 'Lab Staff',
          context,
        }).catch(() => {});
      }

'''
if 'COA Report Date Updated' not in s:
    if insert_marker not in s:
        raise SystemExit('insert marker not found')
    s = s.replace(insert_marker, coa_block + insert_marker, 1)

old_return = "        billingDateWarning,\n      }};\n"
new_return = "        billingDateWarning,\n        coaReportDateWarning,\n      }};\n"
if 'coaReportDateWarning,' not in s:
    if old_return not in s:
        raise SystemExit('return marker not found')
    s = s.replace(old_return, new_return, 1)

p.write_text(s)
