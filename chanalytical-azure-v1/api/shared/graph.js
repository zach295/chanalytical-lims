/**
 * graph.js — Microsoft Graph API client for Azure Functions
 *
 * Token is cached module-level (55 min) — survives warm starts.
 * All SharePoint List CRUD is here. Functions just call these helpers.
 *
 * Environment variables needed:
 *   MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET
 *   SP_SITE_ID  — your SharePoint site ID
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';
const SITE_ID = process.env.SP_SITE_ID;

// ── Token cache ────────────────────────────────────────────────────────────────
let _token = null, _tokenExpiry = 0;

async function getToken() {
  if (_token && Date.now() < _tokenExpiry) return _token;
  const { MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET } = process.env;
  if (!MS_TENANT_ID) throw new Error('MS_TENANT_ID not set');
  const res = await fetch(
    `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: MS_CLIENT_ID,
        client_secret: MS_CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
      }),
    }
  );
  if (!res.ok) throw new Error(`Auth failed: ${await res.text()}`);
  const { access_token, expires_in } = await res.json();
  _token = access_token;
  _tokenExpiry = Date.now() + (expires_in - 60) * 1000;
  return _token;
}

// ── List ID cache ──────────────────────────────────────────────────────────────
const _listIds = {};

async function getListId(listName) {
  if (_listIds[listName]) return _listIds[listName];
  const token = await getToken();
  const res = await fetch(
    `${GRAPH}/sites/${SITE_ID}/lists?$filter=displayName eq '${listName}'&$select=id,displayName`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`Failed to find list "${listName}": ${res.status}`);
  const data = await res.json();
  if (!data.value?.length) throw new Error(`List "${listName}" not found in SharePoint`);
  _listIds[listName] = data.value[0].id;
  return _listIds[listName];
}

// ── Core Graph helpers ─────────────────────────────────────────────────────────
async function graphGet(path) {
  const token = await getToken();
  const res = await fetch(`${GRAPH}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph GET ${path} → ${res.status}: ${await res.text()}`);
  return res.json();
}

async function graphPost(path, body) {
  const token = await getToken();
  const res = await fetch(`${GRAPH}${path}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Graph POST ${path} → ${res.status}: ${await res.text()}`);
  return res.json();
}

async function graphPatch(path, body) {
  const token = await getToken();
  const res = await fetch(`${GRAPH}${path}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Graph PATCH ${path} → ${res.status}: ${await res.text()}`);
  return res.status === 204 ? {} : res.json();
}

async function graphDelete(path) {
  const token = await getToken();
  const res = await fetch(`${GRAPH}${path}`, {
    method: 'DELETE',
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph DELETE ${path} → ${res.status}`);
}

// ── SharePoint List operations ─────────────────────────────────────────────────

/**
 * Get all items from a list, optionally filtered and ordered.
 * Returns array of { _id, ...fields }
 */
async function listItems(listName, { filter, orderby, top, select } = {}) {
  const lid = await getListId(listName);
  let url = `/sites/${SITE_ID}/lists/${lid}/items?$expand=fields`;
  if (filter)  url += `&$filter=${encodeURIComponent(filter)}`;
  if (orderby) url += `&$orderby=${encodeURIComponent(orderby)}`;
  if (top)     url += `&$top=${top}`;
  if (select)  url += `&$select=${select}`;

  const results = [];
  let nextUrl = url;
  while (nextUrl) {
    const page = await graphGet(nextUrl.replace(GRAPH, ''));
    for (const item of (page.value || [])) {
      results.push({ _id: item.id, ...item.fields });
    }
    nextUrl = page['@odata.nextLink'] || null;
  }
  return results;
}

/**
 * Find a single item by field value.
 * Returns { _id, ...fields } or null.
 */
async function findItem(listName, fieldName, fieldValue) {
  const lid = await getListId(listName);
  const escaped = String(fieldValue).replace(/'/g, "''");
  const url = `/sites/${SITE_ID}/lists/${lid}/items?$expand=fields&$filter=fields/${fieldName} eq '${escaped}'&$top=1`;
  const data = await graphGet(url);
  const item = data.value?.[0];
  return item ? { _id: item.id, ...item.fields } : null;
}

/**
 * Create a new list item. Returns { _id, ...fields }.
 */
async function createItem(listName, fields) {
  const lid = await getListId(listName);
  const item = await graphPost(`/sites/${SITE_ID}/lists/${lid}/items`, { fields });
  return { _id: item.id, ...item.fields };
}

/**
 * Update an existing list item by its numeric ID.
 */
async function updateItem(listName, itemId, fields) {
  const lid = await getListId(listName);
  await graphPatch(`/sites/${SITE_ID}/lists/${lid}/items/${itemId}/fields`, fields);
}

/**
 * Delete a list item by its numeric ID.
 */
async function deleteItem(listName, itemId) {
  const lid = await getListId(listName);
  await graphDelete(`/sites/${SITE_ID}/lists/${lid}/items/${itemId}`);
}

/**
 * Upsert — update if item with matching field exists, otherwise create.
 * Returns { _id, created, ...fields }
 */
async function upsertItem(listName, matchField, matchValue, fields) {
  const existing = await findItem(listName, matchField, matchValue);
  if (existing) {
    await updateItem(listName, existing._id, fields);
    return { ...existing, ...fields, created: false };
  }
  const created = await createItem(listName, { ...fields, [matchField]: matchValue });
  return { ...created, created: true };
}

// ── SharePoint file operations (for Excel files) ───────────────────────────────
async function downloadFile(fileId) {
  const token = await getToken();
  const res = await fetch(
    `${GRAPH}/sites/${SITE_ID}/drive/items/${fileId}/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`Download failed: ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

async function listFolder(folderPath) {
  const encoded = folderPath.split('/').map(p => encodeURIComponent(p)).join('/');
  const token = await getToken();
  const res = await fetch(
    `${GRAPH}/sites/${SITE_ID}/drive/root:/${folderPath}:/children?$select=id,name`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) return [];
  const data = await res.json();
  return data.value || [];
}

// ── List names (matches SharePoint list display names) ─────────────────────────
const LISTS = {
  ARCHIVED_INTAKE: 'Archived Intake',
  REVIEW_QUEUE:    'Review Queue',
  CLIENTS:         'Clients',
  USERS:           'Users',
  REJECTED:        'Rejected',
  ACCESSION_LOG:   'Accession Log',
  TEST_TYPES:      'Test Types',
  ELEMENTS:        'Elements',
  ACTIVITY_LOG:    'Activity Log',
  INVENTORY:       'Client Inventory',
};

module.exports = {
  getToken, getListId, listItems, findItem,
  createItem, updateItem, deleteItem, upsertItem,
  downloadFile, listFolder, graphGet, graphPost, graphPatch, graphDelete,
  LISTS,
};
