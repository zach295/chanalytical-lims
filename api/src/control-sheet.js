/**
 * control-sheet.js
 * Module-level cache for fileId + wsId survives warm starts — cuts API calls from 5 to 2.
 * usedRange(valuesOnly=true) returns actual data rows only (not formatted empty cells).
 *
 * Supports two sheet types via the optional `type` body param:
 *   type: 'wq' (default) → C_MMDDYY.xlsx  from Master Control Sheet.xlsx
 *   type: 'rw'           → RW_MMDDYY.xlsx from Master RW Control Sheet.xlsx
 */

const { app }       = require('@azure/functions');
const GRAPH         = 'https://graph.microsoft.com/v1.0';
const SITE_ID       = process.env.SP_SITE_ID;
const CTRL_FOLDER   = process.env.SP_CONTROL_FOLDER || '/sites/Laboratory/Shared Documents/Documents/Control Sheets';
const CTRL_BASE     = '/sites/Laboratory/Shared Documents/Documents/Control Sheets';
const MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];

// ── Module-level caches (survive warm function starts) ────────────────────────
let _token = null, _tokenExpiry = 0;
const _fileCache = {}; // { [fileName]: { fileId, wsId } }

async function getToken() {
  if (_token && Date.now() < _tokenExpiry) return _token;
  const { MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET } = process.env;
  const res = await fetch(
    `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`,
    { method:'POST', headers:{'Content-Type':'application/x-www-form-urlencoded'},
      body: new URLSearchParams({ grant_type:'client_credentials',
        client_id:MS_CLIENT_ID, client_secret:MS_CLIENT_SECRET,
        scope:'https://graph.microsoft.com/.default' }) }
  );
  if (!res.ok) { const b = await res.text(); throw new Error(`Auth ${res.status}: ${b.slice(0,150)}`); }
  const { access_token, expires_in } = await res.json();
  _token = access_token; _tokenExpiry = Date.now() + (expires_in - 60) * 1000;
  return _token;
}

async function graphGet(path, token) {
  const res = await fetch(`${GRAPH}${path}`, { headers:{Authorization:`Bearer ${token}`} });
  if (!res.ok) throw new Error(`GET ${path.slice(0,80)} → ${res.status}`);
  return res.json();
}
async function graphPost(path, body, token) {
  const res = await fetch(`${GRAPH}${path}`, {
    method:'POST', headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body: JSON.stringify(body),
  });
  if (!res.ok && res.status !== 202) throw new Error(`POST ${path.slice(0,80)} → ${res.status}: ${(await res.text()).slice(0,100)}`);
  return res.json().catch(()=>({}));
}
async function graphPatch(path, body, token) {
  const res = await fetch(`${GRAPH}${path}`, {
    method:'PATCH', headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`PATCH ${path.slice(0,80)} → ${res.status}: ${(await res.text()).slice(0,100)}`);
  return res.json().catch(()=>({}));
}

function toDrivePath(p) {
  const i = p.indexOf('Shared Documents/');
  return i >= 0 ? p.slice(i + 17) : p.replace(/^\/+/, '');
}
function dateInfo(mmddyy) {
  const mm = parseInt(mmddyy.slice(0,2)) - 1, yy = parseInt('20'+mmddyy.slice(4,6));
  return { monthFolder:`${MONTHS[mm]} ${yy}` };
}
function todayMMDDYY() {
  const p = {};
  new Intl.DateTimeFormat('en-US',{timeZone:'America/New_York',
    month:'2-digit',day:'2-digit',year:'2-digit'}).formatToParts(new Date()).forEach(({type,value})=>p[type]=value);
  return `${p.month}${p.day}${p.year}`;
}

/** Return file name prefix and template path for the given sheet type. */
function sheetConfig(sheetType, ctrlDrivePath) {
  if (sheetType === 'rw') {
    return {
      filePrefix:   'RCS_',
      templatePath: toDrivePath(process.env.SP_RADON_TEMPLATE || '/sites/Laboratory/Shared Documents/Documents/Control Sheets/Master Radon Control Sheet.xlsx'),
      radonFolder:   true,
    };
  }
  // default: wq
  return {
    filePrefix:   'C_',
    templatePath: `${ctrlDrivePath}/Master Control Sheet.xlsx`,
  };
}

// Get (or cache) fileId + wsId for a control sheet
async function getSheetIds(token, destFilePath, fileName) {
  if (_fileCache[fileName]) {
    console.log(`[CS] Using cached IDs for ${fileName}`);
    return _fileCache[fileName];
  }
  const file = await graphGet(`/sites/${SITE_ID}/drive/root:/${destFilePath}:?$select=id`, token);
  const wsRes = await graphGet(`/sites/${SITE_ID}/drive/items/${file.id}/workbook/worksheets?$select=id,name`, token);
  const wsId = wsRes.value?.[0]?.id;
  if (!wsId) throw new Error('No worksheet found');
  _fileCache[fileName] = { fileId:file.id, wsId };
  console.log(`[CS] Cached IDs for ${fileName}: file=${file.id.slice(0,8)}... ws=${wsId.slice(0,8)}...`);
  return _fileCache[fileName];
}

async function ensureMonthFolder(ctrlDrivePath, monthFolder, token) {
  try { await graphGet(`/sites/${SITE_ID}/drive/root:/${ctrlDrivePath}/${monthFolder}:?$select=id`, token); }
  catch {
    const p = await graphGet(`/sites/${SITE_ID}/drive/root:/${ctrlDrivePath}:?$select=id`, token);
    await graphPost(`/sites/${SITE_ID}/drive/items/${p.id}/children`,
      { name:monthFolder, folder:{}, '@microsoft.graph.conflictBehavior':'replace' }, token);
  }
}

app.http('control-sheet', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
  try {
    const { action, labIds, date, type } = await request.json().catch(() => ({}));
    const sheetType     = (type || 'wq').toLowerCase(); // 'wq' or 'rw'
    const mmddyy        = date || todayMMDDYY();
    const { monthFolder } = dateInfo(mmddyy);
    const ctrlDrivePath = toDrivePath(CTRL_FOLDER);
    const cfg           = sheetConfig(sheetType, toDrivePath(CTRL_BASE));
    const { filePrefix, templatePath } = cfg;
    const fileName      = `${filePrefix}${mmddyy}.xlsx`;
    const targetFolder  = cfg.radonFolder ? `${MONTHS[parseInt(mmddyy.slice(0,2),10)-1]} Radon 20${mmddyy.slice(4,6)}` : monthFolder;
    const destFilePath  = `${ctrlDrivePath}/${targetFolder}/${fileName}`;

    const token = await getToken();
    console.log(`[CS] action=${action} type=${sheetType} file=${fileName}`);

    // Clear stale cache entries for other days
    Object.keys(_fileCache).forEach(k => { if (k !== fileName) delete _fileCache[k]; });

    // ── CREATE ───────────────────────────────────────────────────────────────
    if (action === 'create') {
      try {
        await graphGet(`/sites/${SITE_ID}/drive/root:/${destFilePath}:?$select=id`, token);
        return { status:200, jsonBody:{ success:true, alreadyExists:true, fileName, monthFolder:targetFolder,
            message:`${fileName} already exists in Control Sheets/${targetFolder}` } };
      } catch {}
      await ensureMonthFolder(ctrlDrivePath, targetFolder, token);
      const tpl  = await graphGet(`/sites/${SITE_ID}/drive/root:/${templatePath}:?$select=id`, token);
      const dest = await graphGet(`/sites/${SITE_ID}/drive/root:/${ctrlDrivePath}/${targetFolder}:?$select=id`, token);
      await graphPost(`/sites/${SITE_ID}/drive/items/${tpl.id}/copy`,
        { parentReference:{id:dest.id}, name:fileName }, token);
      // Clear cache so next addLabIds call fetches fresh IDs
      delete _fileCache[fileName];
      return { status:200, jsonBody:{ success:true, fileName, monthFolder:targetFolder,
          message:`Created ${fileName} in Control Sheets/${targetFolder}` } };
    }

    // ── ADD LAB IDS ──────────────────────────────────────────────────────────
    if (action === 'addLabIds') {
      if (!labIds?.length) return { status:400, jsonBody:{error:'labIds required'} };

      // Get cached (or fresh) file + worksheet IDs
      let fileId, wsId;
      try {
        ({ fileId, wsId } = await getSheetIds(token, destFilePath, fileName));
      } catch(e) {
        return { status:404, jsonBody:{ success:false, error:`${fileName} not found — ${e.message}` } };
      }

      const wb = `/sites/${SITE_ID}/drive/items/${fileId}/workbook`;

      // Read column A once — use for both dedup check AND finding next empty row
      let nextRow = 2;
      const existing = new Set();
      try {
        const rangeRes = await graphGet(
          `${wb}/worksheets/${wsId}/range(address='A1:A500')?$select=values`, token);
        const vals = rangeRes.values || [];
        nextRow = 2;
        for (let i = 0; i < vals.length; i++) {
          const cell = vals[i] ? String(vals[i][0] || '').trim() : '';
          if (cell) {
            existing.add(cell); // build dedup set from ALL existing values
            if (i > 0) nextRow = i + 2; // track last filled row (skip header at i=0)
          } else if (i > 0) {
            break; // first empty row found — stop scanning
          }
        }
        console.log(`[CS] Scanned: ${existing.size} entries, nextRow=${nextRow}`);
      } catch(e) { console.warn(`[CS] Row scan failed: ${e.message}, using row 2`); }

      // Filter out any lab IDs already in the sheet (prevents duplicates from retried calls)
      const newIds = labIds.filter(id => !existing.has(id.trim()));
      if (!newIds.length) {
        console.log(`[CS] All lab IDs already in sheet — skipping write`);
        return { status:200, jsonBody:{ success:true, written:0, fileName, startRow:nextRow, message:'Already written' } };
      }

      const endRow = nextRow + newIds.length - 1;
      try {
        await graphPatch(`${wb}/worksheets/${wsId}/range(address='A${nextRow}:A${endRow}')`,
          { values: newIds.map(id=>[id]) }, token);
      } catch(patchErr) {
        // Stale cache (e.g. file was recreated) — clear and retry with fresh IDs
        console.warn(`[CS] PATCH failed (${patchErr.message}) — clearing cache and retrying`);
        delete _fileCache[fileName];
        const fresh = await getSheetIds(token, destFilePath, fileName);
        const wb2 = `/sites/${SITE_ID}/drive/items/${fresh.fileId}/workbook`;
        await graphPatch(`${wb2}/worksheets/${fresh.wsId}/range(address='A${nextRow}:A${endRow}')`,
          { values: newIds.map(id=>[id]) }, token);
      }

      console.log(`[CS] ✅ Wrote ${newIds.length} IDs to ${fileName} rows ${nextRow}-${endRow}`);
      return { status:200, jsonBody:{ success:true, written:newIds.length, fileName, startRow:nextRow } };
    }

    return { status:400, jsonBody:{error:'Unknown action: '+action} };
  } catch(e) {
    console.error('[control-sheet]', e.message);
    return { status:500, jsonBody:{error:e.message} };
  }
  }
});
