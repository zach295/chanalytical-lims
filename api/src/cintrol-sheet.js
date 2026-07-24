/**
 * control-sheet.js — Azure version
 * Creates and populates the daily control sheet Excel file in SharePoint.
 * Module-level cache survives warm starts — cuts API calls from 5 to 2.
 *
 * Actions:
 *   create    — copy Master Control Sheet.xlsx to today's month folder
 *   addLabIds — write lab IDs to column A of today's control sheet
 */
const { app }    = require('@azure/functions');
const { getToken } = require('../shared/graph');

const GRAPH = 'https://graph.microsoft.com/v1.0';
const MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];

// ── Module-level caches (survive warm starts) ──────────────────────────────────
let _token = null, _tokenExpiry = 0;
const _fileCache = {}; // { [fileName]: { fileId, wsId } }

async function getCachedToken() {
  if (_token && Date.now() < _tokenExpiry) return _token;
  _token = await getToken();
  _tokenExpiry = Date.now() + 55 * 60 * 1000;
  return _token;
}

async function graphGet(path, token) {
  const res = await fetch(`${GRAPH}${path}`, { headers:{ Authorization:`Bearer ${token}` } });
  if (!res.ok) throw new Error(`GET ${path.slice(0,80)} → ${res.status}`);
  return res.json();
}

async function graphPost(path, body, token) {
  const res = await fetch(`${GRAPH}${path}`, {
    method:'POST',
    headers:{ Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok && res.status !== 202) throw new Error(`POST ${path.slice(0,80)} → ${res.status}: ${(await res.text()).slice(0,100)}`);
  return res.json().catch(()=>({}));
}

async function graphPatch(path, body, token) {
  const res = await fetch(`${GRAPH}${path}`, {
    method:'PATCH',
    headers:{ Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
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
  const mm = parseInt(mmddyy.slice(0,2)) - 1;
  const yy = parseInt('20' + mmddyy.slice(4,6));
  return { monthFolder: `${MONTHS[mm]} ${yy}` };
}

function todayMMDDYY() {
  const p = {};
  new Intl.DateTimeFormat('en-US', {
    timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit',
  }).formatToParts(new Date()).forEach(({type,value}) => p[type]=value);
  return `${p.month}${p.day}${p.year}`;
}

// Get (or cache) fileId + wsId for a control sheet
async function getSheetIds(token, destFilePath, fileName) {
  if (_fileCache[fileName]) {
    console.log(`[CS] Using cached IDs for ${fileName}`);
    return _fileCache[fileName];
  }
  const siteId = process.env.SP_SITE_ID;
  const file  = await graphGet(`/sites/${siteId}/drive/root:/${destFilePath}:?$select=id`, token);
  const wsRes = await graphGet(`/sites/${siteId}/drive/items/${file.id}/workbook/worksheets?$select=id,name`, token);
  const wsId  = wsRes.value?.[0]?.id;
  if (!wsId) throw new Error('No worksheet found');
  _fileCache[fileName] = { fileId: file.id, wsId };
  return _fileCache[fileName];
}

async function ensureMonthFolder(ctrlDrivePath, monthFolder, token) {
  const siteId = process.env.SP_SITE_ID;
  try {
    await graphGet(`/sites/${siteId}/drive/root:/${ctrlDrivePath}/${monthFolder}:?$select=id`, token);
  } catch {
    const parent = await graphGet(`/sites/${siteId}/drive/root:/${ctrlDrivePath}:?$select=id`, token);
    await graphPost(`/sites/${siteId}/drive/items/${parent.id}/children`,
      { name:monthFolder, folder:{}, '@microsoft.graph.conflictBehavior':'replace' }, token);
  }
}

app.http('control-sheet', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const body = await request.json().catch(() => ({}));
      const { action, labIds, date } = body;

      const CTRL_FOLDER    = process.env.SP_CONTROL_FOLDER || '/sites/Laboratory/Shared Documents/Documents/Test C';
      const TEMPLATE_PATH  = process.env.SP_CONTROL_TEMPLATE || '/sites/Laboratory/Shared Documents/Documents/Test C/Master Control Sheet.xlsx';
      const siteId         = process.env.SP_SITE_ID;

      const mmddyy         = date || todayMMDDYY();
      const { monthFolder } = dateInfo(mmddyy);
      const fileName       = `C_${mmddyy}.xlsx`;
      const ctrlDrivePath  = toDrivePath(CTRL_FOLDER);
      const destFilePath   = `${ctrlDrivePath}/${monthFolder}/${fileName}`;

      const token = await getCachedToken();
      context.log(`[CS] action=${action} file=${fileName}`);

      // ── CREATE ─────────────────────────────────────────────────────────────
      if (action === 'create') {
        try {
          await graphGet(`/sites/${siteId}/drive/root:/${destFilePath}:?$select=id`, token);
          return {
            status: 200,
            jsonBody: { success:true, alreadyExists:true, fileName, monthFolder,
              message: `${fileName} already exists in Control Sheets/${monthFolder}` },
          };
        } catch {}

        await ensureMonthFolder(ctrlDrivePath, monthFolder, token);
        const tpl  = await graphGet(`/sites/${siteId}/drive/root:/${toDrivePath(TEMPLATE_PATH)}:?$select=id`, token);
        const dest = await graphGet(`/sites/${siteId}/drive/root:/${ctrlDrivePath}/${monthFolder}:?$select=id`, token);
        await graphPost(`/sites/${siteId}/drive/items/${tpl.id}/copy`,
          { parentReference:{ id:dest.id }, name:fileName }, token);

        delete _fileCache[fileName];
        return {
          status: 200,
          jsonBody: { success:true, fileName, monthFolder,
            message: `Created ${fileName} in Control Sheets/${monthFolder}` },
        };
      }

      // ── ADD LAB IDS ────────────────────────────────────────────────────────
      if (action === 'addLabIds') {
        if (!labIds?.length) return { status:400, jsonBody:{ error:'labIds required' } };

        let fileId, wsId;
        try {
          ({ fileId, wsId } = await getSheetIds(token, destFilePath, fileName));
        } catch(e) {
          return {
            status: 404,
            jsonBody: { success:false, error:`${fileName} not found — ${e.message}. Run action=create first.` },
          };
        }

        const wb = `/sites/${siteId}/drive/items/${fileId}/workbook`;

        // Read column A to find next empty row and check for duplicates
        let nextRow = 2;
        const existing = new Set();
        try {
          const rangeRes = await graphGet(
            `${wb}/worksheets/${wsId}/range(address='A1:A500')?$select=values`, token);
          const vals = rangeRes.values || [];
          for (let i = 0; i < vals.length; i++) {
            const cell = vals[i] ? String(vals[i][0]||'').trim() : '';
            if (cell) {
              existing.add(cell);
              if (i > 0) nextRow = i + 2;
            } else if (i > 0) { break; }
          }
        } catch(e) { context.log(`[CS] Row scan failed: ${e.message}, using row 2`); }

        const newIds = labIds.filter(id => !existing.has(id.trim()));
        if (!newIds.length) {
          return {
            status: 200,
            jsonBody: { success:true, written:0, fileName, startRow:nextRow, message:'Already written' },
          };
        }

        const endRow = nextRow + newIds.length - 1;
        try {
          await graphPatch(
            `${wb}/worksheets/${wsId}/range(address='A${nextRow}:A${endRow}')`,
            { values: newIds.map(id=>[id]) }, token);
        } catch(patchErr) {
          context.log(`[CS] PATCH failed (${patchErr.message}) — clearing cache and retrying`);
          delete _fileCache[fileName];
          const fresh = await getSheetIds(token, destFilePath, fileName);
          const wb2   = `/sites/${siteId}/drive/items/${fresh.fileId}/workbook`;
          await graphPatch(
            `${wb2}/worksheets/${fresh.wsId}/range(address='A${nextRow}:A${endRow}')`,
            { values: newIds.map(id=>[id]) }, token);
        }

        context.log(`[CS] ✅ Wrote ${newIds.length} IDs to ${fileName} rows ${nextRow}-${endRow}`);
        return {
          status: 200,
          jsonBody: { success:true, written:newIds.length, fileName, startRow:nextRow },
        };
      }

      return { status:400, jsonBody:{ error:'Unknown action: ' + action } };

    } catch(e) {
      context.log('[control-sheet] Error:', e.message);
      return { status:500, jsonBody:{ error:e.message } };
    }
  },
});
