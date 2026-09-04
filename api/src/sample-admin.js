const { app } = require('@azure/functions');
const crypto = require('crypto');
const { getToken, listItems, createItem, deleteItem, LISTS } = require('../shared/graph');
const { writeActivityLog } = require('../shared/audit');

const GRAPH = 'https://graph.microsoft.com/v1.0';
const TZ = 'America/New_York';
const RECOVERY_ROOT = 'Documents/LIMS Recovery/Deleted Samples';
const TOKEN_TTL_MS = 15 * 60 * 1000;
const RECOVERY_DAYS = 30;

function hashPassword(pw) {
  let h = 0;
  for (let i = 0; i < String(pw || '').length; i++) {
    h = ((h << 5) - h) + String(pw).charCodeAt(i);
    h = h & h;
  }
  const s = String(pw || '');
  return h.toString(36) + s.length;
}
const WELCOME_HASH = '-tlew818';

function tokenSecret() {
  return process.env.ADMIN_REAUTH_SECRET || process.env.MS_CLIENT_SECRET || 'chanalytical-admin-reauth';
}
function b64url(v) { return Buffer.from(v).toString('base64url'); }
function signToken(payload) {
  const body = b64url(JSON.stringify(payload));
  const sig = crypto.createHmac('sha256', tokenSecret()).update(body).digest('base64url');
  return `${body}.${sig}`;
}
function verifyToken(token) {
  if (!token || !token.includes('.')) return null;
  const [body, sig] = token.split('.');
  const expected = crypto.createHmac('sha256', tokenSecret()).update(body).digest('base64url');
  const a = Buffer.from(sig || '');
  const b = Buffer.from(expected || '');
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  let payload;
  try { payload = JSON.parse(Buffer.from(body, 'base64url').toString('utf8')); } catch { return null; }
  if (!payload.exp || Date.now() > payload.exp || payload.role !== 'admin') return null;
  return payload;
}

function baseOf(v) { return String(v || '').trim().split(' ')[0].trim(); }
function unique(arr) { return [...new Set(arr.filter(Boolean))]; }
function isRadonAccession(r) {
  return /radon/i.test(String(r.field_3 || '')) || /(^|\W)RW(\W|$)/i.test(String(r.field_4 || '')) || /\sRW\b/i.test(String(r.field_2 || ''));
}
function etStamp(d = new Date()) {
  return d.toLocaleString('en-US', { timeZone: TZ, hour12: false });
}
function safeFilePart(v) { return String(v || '').replace(/[^A-Za-z0-9_.-]/g, '_'); }
function encPath(path) { return path.split('/').map(encodeURIComponent).join('/'); }

async function findUserByEmail(email) {
  const all = await listItems(LISTS.USERS, { top: 500 });
  const key = String(email || '').toLowerCase().trim();
  return all.find(r => String(r.Title || '').toLowerCase().trim() === key) || null;
}
async function verifyAdminPassword(email, password) {
  const user = await findUserByEmail(email);
  if (!user) return { ok:false, error:'Admin account not found.' };
  const role = String(user.field_2 || '').toLowerCase();
  const active = user.field_9 !== false && user.field_9 !== 'FALSE' && role !== 'deactivated';
  if (!active || !role.split(',').map(v=>v.trim()).includes('admin')) return { ok:false, error:'This account does not have active Admin access.' };
  const storedHash = user.field_8 || WELCOME_HASH;
  const inputHash = hashPassword(password || '');
  if (inputHash !== storedHash) return { ok:false, error:'Incorrect password.' };
  return { ok:true, name:user.field_1 || email, email:String(user.Title || email).toLowerCase().trim() };
}

async function getListIdByName(name, token) {
  const siteId = process.env.SP_SITE_ID;
  const r = await fetch(`${GRAPH}/sites/${siteId}/lists?$select=id,displayName&$top=100`, { headers:{ Authorization:`Bearer ${token}` } });
  if (!r.ok) throw new Error(`Unable to enumerate SharePoint lists (${r.status})`);
  return ((await r.json()).value || []).find(x => x.displayName === name)?.id || null;
}
async function readListRaw(name, token) {
  const siteId = process.env.SP_SITE_ID;
  const listId = await getListIdByName(name, token);
  if (!listId) return [];
  const out = [];
  let url = `${GRAPH}/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`;
  while (url) {
    const r = await fetch(url, { headers:{ Authorization:`Bearer ${token}` } });
    if (!r.ok) throw new Error(`${name} read failed (${r.status})`);
    const d = await r.json();
    out.push(...(d.value || []).map(i => ({ _id:i.id, ...(i.fields || {}) })));
    url = d['@odata.nextLink'] || null;
  }
  return out;
}
async function deleteListRaw(name, ids, token) {
  const siteId = process.env.SP_SITE_ID;
  const listId = await getListIdByName(name, token);
  if (!listId || !ids.length) return 0;
  let n = 0;
  for (const id of ids) {
    const r = await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items/${id}`, { method:'DELETE', headers:{ Authorization:`Bearer ${token}` } });
    if (!r.ok && r.status !== 404) throw new Error(`${name} delete ${id} failed (${r.status})`);
    if (r.ok) n++;
  }
  return n;
}

const SKIP_FIELDS = new Set([
  '_id','id','ContentType','Modified','Created','AuthorLookupId','EditorLookupId','Attachments',
  '_UIVersionString','Edit','LinkTitle','LinkTitleNoMenu','ItemChildCount','FolderChildCount',
  '_ComplianceFlags','_ComplianceTag','FileLeafRef','FileDirRef','FileRef'
]);
function cleanFields(row) {
  const out = {};
  for (const [k,v] of Object.entries(row || {})) {
    if (SKIP_FIELDS.has(k) || k.startsWith('@odata') || k.startsWith('_')) continue;
    if (v === undefined) continue;
    out[k] = v;
  }
  return out;
}
async function restoreListRaw(name, rows) {
  let n = 0;
  for (const row of rows || []) {
    await createItem(name, cleanFields(row));
    n++;
  }
  return n;
}

function monthInfo(datePrefix) {
  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const mm = parseInt(datePrefix.slice(0,2),10);
  const yy = datePrefix.slice(4,6);
  return { month:months[mm-1] || datePrefix.slice(0,2), year:`20${yy}` };
}
function controlRoot() {
  const raw = process.env.SP_CONTROL_FOLDER || '/sites/Laboratory/Shared Documents/Documents/Lab Scans/Test C';
  const marker = 'Shared Documents/';
  const i = raw.indexOf(marker);
  return i >= 0 ? raw.slice(i + marker.length) : raw.replace(/^\/+/, '');
}
async function driveItemByPath(relPath, token) {
  const siteId = process.env.SP_SITE_ID;
  const r = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encPath(relPath)}?$select=id,name,parentReference`, { headers:{ Authorization:`Bearer ${token}` } });
  if (!r.ok) return null;
  return r.json();
}
async function findControlFile(baseId, radon, token) {
  const prefix = baseId.slice(0,6);
  const { month, year } = monthInfo(prefix);
  const root = controlRoot();
  const candidates = radon
    ? [`${root}/${month} Radon ${year}/RCS_${prefix}.xlsx`]
    : [`${root}/${month} ${year}/C_${prefix}.xlsx`, `${root}/C_${prefix}.xlsx`];
  for (const path of candidates) {
    const item = await driveItemByPath(path, token);
    if (item) return { ...item, path };
  }
  return null;
}
async function openWorkbook(fileId, token) {
  const siteId = process.env.SP_SITE_ID;
  const auth = { Authorization:`Bearer ${token}`, 'Content-Type':'application/json' };
  const s = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook/createSession`, { method:'POST', headers:auth, body:JSON.stringify({persistChanges:true}) });
  const sd = await s.json().catch(()=>({}));
  if (!s.ok || !sd.id) throw new Error(`Workbook session failed (${s.status})`);
  const headers = { ...auth, 'workbook-session-id':sd.id };
  const base = `${GRAPH}/sites/${siteId}/drive/items/${fileId}/workbook`;
  const w = await fetch(`${base}/worksheets?$select=id,name`, { headers });
  const wd = await w.json();
  const ws = (wd.value || [])[0];
  if (!ws) throw new Error('Workbook has no worksheet');
  return { base, headers, ws, close:()=>fetch(`${base}/closeSession`, { method:'POST', headers }).catch(()=>{}) };
}
async function readSheetRow(file, baseId, radon, token) {
  if (!file) return null;
  const wb = await openWorkbook(file.id, token);
  try {
    const a = await fetch(`${wb.base}/worksheets/${wb.ws.id}/range(address='A1:A250')?$select=values`, { headers:wb.headers });
    if (!a.ok) throw new Error(`Control sheet row lookup failed (${a.status})`);
    const vals = (await a.json()).values || [];
    let row = -1;
    for (let i=0;i<vals.length;i++) {
      const b = baseOf(vals[i]?.[0]);
      if (b === baseId) { row = i+1; break; }
    }
    if (row < 0) return null;
    const endCol = radon ? 'G' : 'AE';
    const rr = await fetch(`${wb.base}/worksheets/${wb.ws.id}/range(address='A${row}:${endCol}${row}')?$select=values`, { headers:wb.headers });
    if (!rr.ok) throw new Error(`Control sheet snapshot failed (${rr.status})`);
    const values = (await rr.json()).values || [[]];
    return { kind:radon?'radon':'control', fileId:file.id, filePath:file.path, worksheetId:wb.ws.id, worksheetName:wb.ws.name, row, endCol, values:values[0] || [] };
  } finally { await wb.close(); }
}
async function clearSheetSnapshot(snap, token) {
  const siteId = process.env.SP_SITE_ID;
  const wb = await openWorkbook(snap.fileId, token);
  try {
    const width = snap.kind === 'radon' ? 7 : 31;
    const empty = Array(width).fill('');
    const r = await fetch(`${wb.base}/worksheets/${wb.ws.id}/range(address='A${snap.row}:${snap.endCol}${snap.row}')`, {
      method:'PATCH', headers:wb.headers, body:JSON.stringify({ values:[empty] })
    });
    if (!r.ok) throw new Error(`${snap.kind} control sheet clear failed (${r.status})`);
  } finally { await wb.close(); }
}
async function restoreSheetSnapshot(snap, token) {
  const file = await driveItemByPath(snap.filePath, token);
  if (!file) throw new Error(`Recovery workbook missing: ${snap.filePath}`);
  const wb = await openWorkbook(file.id, token);
  try {
    const r = await fetch(`${wb.base}/worksheets/${wb.ws.id}/range(address='A${snap.row}:${snap.endCol}${snap.row}')`, {
      method:'PATCH', headers:wb.headers, body:JSON.stringify({ values:[snap.values || []] })
    });
    if (!r.ok) throw new Error(`${snap.kind} control sheet restore failed (${r.status})`);
  } finally { await wb.close(); }
}

async function ensureFolder(path, token) {
  const siteId = process.env.SP_SITE_ID;
  const parts = path.split('/').filter(Boolean);
  let parentId = 'root';
  let currentPath = '';
  for (const part of parts) {
    currentPath = currentPath ? `${currentPath}/${part}` : part;
    const existing = await driveItemByPath(currentPath, token);
    if (existing) { parentId = existing.id; continue; }
    const url = parentId === 'root'
      ? `${GRAPH}/sites/${siteId}/drive/root/children`
      : `${GRAPH}/sites/${siteId}/drive/items/${parentId}/children`;
    const r = await fetch(url, { method:'POST', headers:{ Authorization:`Bearer ${token}`, 'Content-Type':'application/json' }, body:JSON.stringify({ name:part, folder:{}, '@microsoft.graph.conflictBehavior':'fail' }) });
    if (!r.ok) {
      const retry = await driveItemByPath(currentPath, token);
      if (!retry) throw new Error(`Unable to create recovery folder ${currentPath} (${r.status})`);
      parentId = retry.id;
    } else parentId = (await r.json()).id;
  }
  return parentId;
}
async function saveSnapshot(snapshot, token) {
  await ensureFolder(RECOVERY_ROOT, token);
  const siteId = process.env.SP_SITE_ID;
  const name = `${safeFilePart(snapshot.deletionId)}.json`;
  const path = `${RECOVERY_ROOT}/${name}`;
  const r = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encPath(path)}:/content`, {
    method:'PUT', headers:{ Authorization:`Bearer ${token}`, 'Content-Type':'application/json' }, body:JSON.stringify(snapshot)
  });
  if (!r.ok) throw new Error(`Recovery snapshot write failed (${r.status})`);
  return { ...(await r.json()), path };
}
async function loadSnapshot(path, token) {
  const siteId = process.env.SP_SITE_ID;
  const r = await fetch(`${GRAPH}/sites/${siteId}/drive/root:/${encPath(path)}:/content`, { headers:{ Authorization:`Bearer ${token}` } });
  if (!r.ok) throw new Error('Recovery snapshot not found or expired.');
  return r.json();
}
async function listRecoveryFiles(token) {
  const siteId = process.env.SP_SITE_ID;
  const folder = await driveItemByPath(RECOVERY_ROOT, token);
  if (!folder) return [];
  const r = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${folder.id}/children?$select=id,name,createdDateTime,lastModifiedDateTime&$top=500`, { headers:{ Authorization:`Bearer ${token}` } });
  if (!r.ok) return [];
  return (await r.json()).value || [];
}
async function deleteDriveItem(id, token) {
  const siteId = process.env.SP_SITE_ID;
  const r = await fetch(`${GRAPH}/sites/${siteId}/drive/items/${id}`, { method:'DELETE', headers:{ Authorization:`Bearer ${token}` } });
  if (!r.ok && r.status !== 404) throw new Error(`Recovery file delete failed (${r.status})`);
}
async function purgeExpired(token) {
  const files = await listRecoveryFiles(token);
  const cutoff = Date.now() - RECOVERY_DAYS*86400000;
  let purged = 0;
  for (const f of files) {
    const t = new Date(f.createdDateTime || f.lastModifiedDateTime || 0).getTime();
    if (t && t < cutoff) { await deleteDriveItem(f.id, token); purged++; }
  }
  return purged;
}

async function buildGroup(baseId, token) {
  const accessions = await readListRaw('Accession Log', token);
  const selected = accessions.filter(r => baseOf(r.field_1 || r.field_2) === baseId);
  if (!selected.length) throw new Error(`Lab ID ${baseId} was not found in the Accession Log.`);
  const createdAt = String(selected[0].Title || selected[0].Created || '').trim();
  if (!createdAt) throw new Error(`Accession Log row for ${baseId} does not have a CreatedAt/Title linkage value.`);
  const group = accessions.filter(r => String(r.Title || r.Created || '').trim() === createdAt);
  const baseIds = unique(group.map(r => baseOf(r.field_1 || r.field_2)));
  const radonBaseIds = unique(group.filter(isRadonAccession).map(r => baseOf(r.field_1 || r.field_2)));

  const [archived, billing, cache, rejected] = await Promise.all([
    readListRaw('Archived Intake', token),
    readListRaw('Reports to be Billed', token),
    readListRaw('Results Cache', token),
    readListRaw('Rejected', token),
  ]);
  const inSet = v => baseIds.includes(baseOf(v));
  const lists = {
    'Accession Log': group,
    'Archived Intake': archived.filter(r => inSet(r.field_1)),
    'Reports to be Billed': billing.filter(r => inSet(r.Title)),
    'Results Cache': cache.filter(r => inSet(r.LabID)),
    'Rejected': rejected.filter(r => inSet(r.field_1)),
  };

  const controlSheets = [];
  for (const id of baseIds) {
    const f = await findControlFile(id, false, token);
    const snap = await readSheetRow(f, id, false, token);
    if (snap) controlSheets.push({ baseId:id, ...snap });
  }
  for (const id of radonBaseIds) {
    const f = await findControlFile(id, true, token);
    const snap = await readSheetRow(f, id, true, token);
    if (snap) controlSheets.push({ baseId:id, ...snap });
  }

  return { requestedBaseId:baseId, createdAt, baseIds, radonBaseIds, lists, controlSheets };
}

function groupSummary(group) {
  return {
    requestedBaseId: group.requestedBaseId,
    createdAt: group.createdAt,
    baseIds: group.baseIds,
    radonBaseIds: group.radonBaseIds,
    tests: group.lists['Accession Log'].map(r => ({ baseId:baseOf(r.field_1||r.field_2), fullId:r.field_2||'', test:r.field_3||'', suffix:r.field_4||'' })),
    counts: Object.fromEntries(Object.entries(group.lists).map(([k,v])=>[k,v.length])),
    controlSheets: group.controlSheets.map(s => ({ kind:s.kind, baseId:s.baseId, file:s.filePath, row:s.row })),
  };
}

async function doDelete(group, reason, admin, token, context) {
  const deletionId = `${Date.now()}-${group.requestedBaseId}`;
  const deletedAt = new Date();
  const expiresAt = new Date(deletedAt.getTime() + RECOVERY_DAYS*86400000);
  const snapshot = {
    version:1, deletionId, deletedAt:deletedAt.toISOString(), expiresAt:expiresAt.toISOString(),
    deletedBy:admin.name || admin.email, deletedByEmail:admin.email, reason,
    requestedBaseId:group.requestedBaseId, createdAt:group.createdAt, baseIds:group.baseIds, radonBaseIds:group.radonBaseIds,
    lists:Object.fromEntries(Object.entries(group.lists).map(([k,v])=>[k,v.map(cleanFields)])),
    controlSheets:group.controlSheets.map(s => ({ ...s, fileId:undefined, worksheetId:undefined })),
  };
  const saved = await saveSnapshot(snapshot, token);

  const results = {};
  for (const [name, rows] of Object.entries(group.lists)) {
    results[name] = await deleteListRaw(name, rows.map(r=>r._id), token);
  }
  for (const s of group.controlSheets) await clearSheetSnapshot(s, token);
  results['Control Sheets'] = group.controlSheets.length;

  await writeActivityLog({
    labId:group.requestedBaseId,
    type:'Admin Sample Deletion',
    by:admin.name || admin.email,
    notes:[
      `Reason: ${reason}`,
      `CreatedAt group: ${group.createdAt}`,
      `Affected Lab IDs: ${group.baseIds.join(', ')}`,
      `Radon Lab IDs: ${group.radonBaseIds.join(', ') || 'None'}`,
      `Recovery ID: ${deletionId}`,
      `Recovery expires: ${expiresAt.toISOString()}`,
      `Removed: ${Object.entries(results).map(([k,v])=>`${k}=${v}`).join(', ')}`,
    ].join('\n'),
    context,
  });
  return { deletionId, expiresAt:expiresAt.toISOString(), recoveryPath:saved.path, results };
}

async function doRestore(snapshot, admin, token, context) {
  if (Date.now() > new Date(snapshot.expiresAt).getTime()) throw new Error('This recovery snapshot has expired.');
  const results = {};
  // Restore lists first so the sample exists operationally before sheet rows return.
  for (const [name, rows] of Object.entries(snapshot.lists || {})) {
    results[name] = await restoreListRaw(name, rows);
  }
  for (const s of snapshot.controlSheets || []) await restoreSheetSnapshot(s, token);
  results['Control Sheets'] = (snapshot.controlSheets || []).length;

  await writeActivityLog({
    labId:snapshot.requestedBaseId,
    type:'Sample Deletion Reversed',
    by:admin.name || admin.email,
    notes:[
      `Recovery ID: ${snapshot.deletionId}`,
      `Original deletion reason: ${snapshot.reason}`,
      `Original deleted by: ${snapshot.deletedBy}`,
      `CreatedAt group: ${snapshot.createdAt}`,
      `Restored Lab IDs: ${(snapshot.baseIds || []).join(', ')}`,
      `Restored: ${Object.entries(results).map(([k,v])=>`${k}=${v}`).join(', ')}`,
    ].join('\n'),
    context,
  });
  return results;
}

app.http('sample-admin', {
  methods:['POST'], authLevel:'anonymous',
  handler:async (request, context) => {
    try {
      const body = await request.json().catch(()=>({}));
      const action = body.action || '';

      if (action === 'verify') {
        const verified = await verifyAdminPassword(body.email, body.password);
        if (!verified.ok) return { status:401, jsonBody:{ success:false, error:verified.error } };
        const token = signToken({ email:verified.email, name:verified.name, role:'admin', exp:Date.now()+TOKEN_TTL_MS });
        return { status:200, jsonBody:{ success:true, token, expiresInMinutes:15, admin:{ email:verified.email, name:verified.name } } };
      }

      const admin = verifyToken(body.adminToken);
      if (!admin) return { status:401, jsonBody:{ error:'Admin verification expired. Re-enter your password.' } };
      const token = await getToken();
      await purgeExpired(token).catch(e => context.log('[sample-admin] purge warning:', e.message));

      if (action === 'preview') {
        const baseId = baseOf(body.baseId);
        if (!/^\d{6}-\d{3}$/.test(baseId)) return { status:400, jsonBody:{ error:'Enter a Lab ID like MMDDYY-###.' } };
        const group = await buildGroup(baseId, token);
        return { status:200, jsonBody:{ success:true, group:groupSummary(group) } };
      }

      if (action === 'delete') {
        const baseId = baseOf(body.baseId);
        const reason = String(body.reason || '').trim();
        if (reason.length < 3) return { status:400, jsonBody:{ error:'A deletion reason is required.' } };
        const group = await buildGroup(baseId, token);
        // Client sends the CreatedAt from preview. Refuse if the group changed between preview and deletion.
        if (!body.createdAt || String(body.createdAt) !== String(group.createdAt)) {
          return { status:409, jsonBody:{ error:'Accession group changed since preview. Find the sample again before deleting.' } };
        }
        const result = await doDelete(group, reason, admin, token, context);
        return { status:200, jsonBody:{ success:true, ...result } };
      }

      if (action === 'recoveries') {
        const files = await listRecoveryFiles(token);
        const items = [];
        for (const f of files.sort((a,b)=>String(b.createdDateTime).localeCompare(String(a.createdDateTime)))) {
          if (!/\.json$/i.test(f.name)) continue;
          try {
            const path = `${RECOVERY_ROOT}/${f.name}`;
            const s = await loadSnapshot(path, token);
            if (Date.now() <= new Date(s.expiresAt).getTime()) items.push({
              deletionId:s.deletionId, requestedBaseId:s.requestedBaseId, baseIds:s.baseIds || [], radonBaseIds:s.radonBaseIds || [],
              deletedAt:s.deletedAt, expiresAt:s.expiresAt, deletedBy:s.deletedBy, reason:s.reason, createdAt:s.createdAt, path,
            });
          } catch(e) { context.log('[sample-admin] recovery read warning:', e.message); }
        }
        return { status:200, jsonBody:{ success:true, recoveries:items } };
      }

      if (action === 'restore-preview') {
        const snap = await loadSnapshot(body.path, token);
        return { status:200, jsonBody:{ success:true, snapshot:{
          deletionId:snap.deletionId, requestedBaseId:snap.requestedBaseId, baseIds:snap.baseIds, radonBaseIds:snap.radonBaseIds,
          deletedAt:snap.deletedAt, expiresAt:snap.expiresAt, deletedBy:snap.deletedBy, reason:snap.reason, createdAt:snap.createdAt,
          counts:Object.fromEntries(Object.entries(snap.lists || {}).map(([k,v])=>[k,v.length])),
          controlSheets:(snap.controlSheets || []).map(s=>({kind:s.kind,baseId:s.baseId,file:s.filePath,row:s.row})), path:body.path,
        } } };
      }

      if (action === 'restore') {
        const snap = await loadSnapshot(body.path, token);
        const results = await doRestore(snap, admin, token, context);
        const item = await driveItemByPath(body.path, token);
        if (item) await deleteDriveItem(item.id, token); // restored snapshots no longer count as deleted/recoverable
        return { status:200, jsonBody:{ success:true, results } };
      }

      return { status:400, jsonBody:{ error:`Unknown action: ${action}` } };
    } catch(e) {
      context.log('[sample-admin] Error:', e.stack || e.message);
      return { status:500, jsonBody:{ error:e.message } };
    }
  }
});

// Daily hard purge so deleted-sample recovery data is never retained beyond 30 days.
app.timer('sample-admin-recovery-purge', {
  schedule:'0 17 3 * * *',
  handler:async (_timer, context) => {
    try {
      const token = await getToken();
      const purged = await purgeExpired(token);
      context.log(`[sample-admin] purged ${purged} recovery snapshot(s) older than ${RECOVERY_DAYS} days`);
    } catch(e) {
      context.log('[sample-admin] purge failed:', e.message);
    }
  }
});
