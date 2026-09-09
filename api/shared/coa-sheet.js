const SHEETS_ID = '15403E6ZaZFQuKNTtgJcb6-2jmlLnk02eq04BGPATqTw';
const SHEETS_TAB = 'Form Responses';

async function getSheetsToken() {
  const sa = JSON.parse(process.env.GMAIL_SERVICE_ACCOUNT || '{}');
  if (!sa.private_key) throw new Error('GMAIL_SERVICE_ACCOUNT missing');
  const crypto = require('crypto');
  const now = Math.floor(Date.now() / 1000);
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify({
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600,
  })).toString('base64url');
  const sign = crypto.createSign('RSA-SHA256');
  sign.update(`${header}.${payload}`);
  const sig = sign.sign(sa.private_key, 'base64url');
  const jwt = `${header}.${payload}.${sig}`;
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
  });
  const data = await res.json();
  if (!data.access_token) throw new Error('Sheets token failed: ' + JSON.stringify(data));
  return data.access_token;
}

function toCoaDate(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const pad = n => String(n).padStart(2, '0');
  let m = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return `${pad(m[2])}-${pad(m[3])}-${m[1].slice(-2)}`;
  m = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (m) return `${pad(m[1])}-${pad(m[2])}-${m[3].slice(-2)}`;
  return raw;
}

function nextBusinessDayCoa(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  let d;
  let m = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  else {
    m = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
    if (!m) return '';
    const year = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
    d = new Date(year, Number(m[1]) - 1, Number(m[2]));
  }
  if (Number.isNaN(d.getTime())) return '';
  d.setDate(d.getDate() + 1);
  while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1);
  return `${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}-${String(d.getFullYear()).slice(-2)}`;
}

function displayTest(testName) {
  const name = String(testName || '').trim();
  if (name === 'Basic Safety (FHA)') return 'Basic Safety';
  if (name === 'Expanded Safety (Mortgage Test)') return 'Expanded Safety (Mtg Test)';
  return name;
}

function rowIsRejected(row) {
  const status = String(row?.field_14 || '').trim().toLowerCase();
  const test = String(row?.field_2 || '').trim().toLowerCase();
  const fullId = String(row?.field_1 || '').trim();
  return status === 'rejected' || /^rejected|^wq.*reject|^rw.*reject/i.test(test) || /\bREJ\b/i.test(fullId);
}

async function syncOnce(baseId, archivedRows, context) {
  const token = await getSheetsToken();
  const metaRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SHEETS_ID}?fields=sheets(properties(title,gridProperties(rowCount)))`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!metaRes.ok) throw new Error(`COA metadata lookup failed (${metaRes.status})`);
  const meta = await metaRes.json();
  const sheetMeta = (meta.sheets || []).find(sh => sh.properties?.title === SHEETS_TAB);
  const rowCount = Math.max(Number(sheetMeta?.properties?.gridProperties?.rowCount || 1000), 2);
  const range = `${SHEETS_TAB}!A1:N${rowCount}`;
  const rangeEnc = encodeURIComponent(range);

  const gridRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SHEETS_ID}/values/${rangeEnc}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!gridRes.ok) throw new Error(`COA read failed (${gridRes.status})`);
  const grid = (await gridRes.json()).values || [];

  const existing = [];
  for (let rowNum = 2; rowNum <= rowCount; rowNum++) {
    const vals = grid[rowNum - 1] || [];
    if (String(vals[7] || '').trim() === baseId) existing.push({ rowNum, vals });
  }

  const desired = [];
  for (const row of archivedRows || []) {
    const tests = String(row.field_2 || '')
      .split(/[|;]/)
      .map(displayTest)
      .map(v => v.trim())
      .filter(Boolean);
    const rowTests = tests.length ? tests : [''];
    for (const testName of rowTests) desired.push({ source: row, testName });
  }
  if (!desired.length) return { updated: 0, cleared: 0, rows: [] };

  const usedExisting = new Set();
  const assignments = [];
  const firstClientCode = String(existing.find(r => String(r.vals[5] || '').trim())?.vals?.[5] || '').trim();
  const firstReportDate = String(existing.find(r => String(r.vals[6] || '').trim())?.vals?.[6] || '').trim();

  for (const d of desired) {
    let match = existing.find(r => !usedExisting.has(r.rowNum) && String(r.vals[12] || '').trim() === d.testName);
    if (!match) match = existing.find(r => !usedExisting.has(r.rowNum));
    if (match) usedExisting.add(match.rowNum);
    assignments.push({ ...d, match });
  }

  const needBlank = assignments.filter(a => !a.match).length;
  const blankRows = [];
  if (needBlank) {
    for (let rowNum = 2; rowNum <= rowCount && blankRows.length < needBlank; rowNum++) {
      if (existing.some(r => r.rowNum === rowNum)) continue;
      const vals = grid[rowNum - 1] || [];
      const isBlank = Array.from({ length: 14 }, (_, c) => String(vals[c] || '').trim()).every(v => !v);
      if (isBlank) blankRows.push(rowNum);
    }
    if (blankRows.length < needBlank) {
      throw new Error(`COA sheet does not have enough existing empty rows (${blankRows.length}/${needBlank})`);
    }
  }

  let blankIdx = 0;
  const writeData = [];
  for (const a of assignments) {
    const rowNum = a.match?.rowNum || blankRows[blankIdx++];
    const oldVals = a.match?.vals || [];
    const src = a.source;
    const status = String(src.field_14 || '').trim().toLowerCase();
    const reportDate = rowIsRejected(src)
      ? nextBusinessDayCoa(src.field_6)
      : status === 'reported'
        ? String(oldVals[6] || firstReportDate || '').trim()
        : '';
    const clientCode = String(oldVals[5] || firstClientCode || '').trim();
    const zip = String(src.field_11 || '').replace(/[^0-9]/g, '').padStart(5, '0');
    const qty = oldVals[13] === undefined || oldVals[13] === '' ? 1 : oldVals[13];
    const values = [[
      toCoaDate(src.field_6),
      src.field_7 || '',
      toCoaDate(src.field_4),
      src.field_5 || '',
      src.field_3 || '',
      clientCode,
      reportDate,
      baseId,
      src.field_8 || '',
      src.field_9 || '',
      src.field_10 || 'ME',
      zip,
      a.testName,
      qty,
    ]];
    writeData.push({ range: `${SHEETS_TAB}!A${rowNum}:N${rowNum}`, values });
  }

  const surplus = existing.filter(r => !usedExisting.has(r.rowNum));
  for (const r of surplus) {
    writeData.push({ range: `${SHEETS_TAB}!A${r.rowNum}:N${r.rowNum}`, values: [Array(14).fill('')] });
  }

  const writeRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SHEETS_ID}/values:batchUpdate`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ valueInputOption: 'RAW', data: writeData }),
    }
  );
  if (!writeRes.ok) {
    const err = await writeRes.text().catch(() => '');
    throw new Error(`COA sync write failed (${writeRes.status}): ${err.slice(0, 250)}`);
  }

  const verifyRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SHEETS_ID}/values/${rangeEnc}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!verifyRes.ok) throw new Error(`COA verification read failed (${verifyRes.status})`);
  const verifyGrid = (await verifyRes.json()).values || [];
  const actualTests = [];
  for (let r = 1; r < verifyGrid.length; r++) {
    if (String(verifyGrid[r]?.[7] || '').trim() === baseId) actualTests.push(String(verifyGrid[r]?.[12] || '').trim());
  }
  const expectedTests = desired.map(d => d.testName).sort();
  actualTests.sort();
  if (actualTests.length !== expectedTests.length || actualTests.some((v, i) => v !== expectedTests[i])) {
    throw new Error(`COA verification mismatch for ${baseId}: expected [${expectedTests.join(', ')}], found [${actualTests.join(', ')}]`);
  }

  if (context) context.log(`[COA sync] ${baseId}: wrote ${assignments.length}, cleared ${surplus.length}`);
  return { updated: assignments.length, cleared: surplus.length, rows: expectedTests };
}

async function syncCoaFromArchivedRows(baseId, archivedRows, context) {
  let lastError;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const result = await syncOnce(baseId, archivedRows, context);
      return { ...result, attempt };
    } catch (e) {
      lastError = e;
      if (context) context.log(`[COA sync] ${baseId} attempt ${attempt}/3 failed: ${e.message}`);
      if (attempt < 3) await new Promise(resolve => setTimeout(resolve, attempt * 1500));
    }
  }
  throw lastError || new Error('COA sync failed');
}

module.exports = { syncCoaFromArchivedRows, toCoaDate, nextBusinessDayCoa, displayTest };
