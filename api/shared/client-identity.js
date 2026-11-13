'use strict';

/**
 * Phase 1: read-only client identity / DBA presentation resolver.
 * The official ClientName remains the account identity for billing and storage.
 * The current Clients list uses the internal "Phone" field for DBA Name.
 * Never automatically choose between duplicate names or DBA names.
 */
const { listItems, LISTS } = require('./graph');

const clean = value => String(value ?? '').trim().replace(/\s+/g, ' ');
const key = value => clean(value).toLocaleLowerCase('en-US');
const isPublicPricing = value => key(value) === 'public pricing';

function publicPersonalKey(value) {
  const raw = clean(value).replace(/^public-\s*/i, '').trim();
  if (!raw) return '';
  if (raw.includes(',')) {
    const pos = raw.indexOf(',');
    const last = clean(raw.slice(0, pos)), first = clean(raw.slice(pos + 1));
    return last && first ? key(last + ', ' + first) : key(raw);
  }
  const parts = raw.split(/\s+/);
  return parts.length > 1
    ? key(parts[parts.length - 1] + ', ' + parts.slice(0, -1).join(' '))
    : key(raw);
}

function toClient(row) {
  const officialName = clean(row.clientName ?? row.ClientName ?? row.Title);
  const dbaName = clean(row.dbaName ?? row.Phone);
  return {
    clientId: String(row.clientId ?? row._id ?? row.id ?? ''),
    officialName,
    dbaName,
    displayName: dbaName || officialName,
    pricingCategory: clean(row.pricingCategory ?? row.PricingCategory),
    abbrev: clean(row.abbrev ?? row.Abbrev),
  };
}

function buildClientDirectory(rows) {
  if (!Array.isArray(rows)) throw new TypeError('Client records must be an array');
  const clients = rows.map(toClient).filter(c => c.officialName);
  const byId = new Map(), byName = new Map(), byPublicName = new Map();
  function add(index, value, client) {
    const normalized = key(value);
    if (!normalized) return;
    if (!index.has(normalized)) index.set(normalized, new Set());
    index.get(normalized).add(client);
  }
  for (const client of clients) {
    if (client.clientId) byId.set(client.clientId, client);
    add(byName, client.officialName, client);
    add(byName, client.dbaName, client);
    if (isPublicPricing(client.pricingCategory)) {
      add(byPublicName, publicPersonalKey(client.officialName), client);
      add(byPublicName, publicPersonalKey(client.dbaName), client);
    }
  }
  const conflicts = [];
  for (const index of [byName, byPublicName]) {
    for (const [name, matches] of index) {
      if (matches.size > 1 && !conflicts.some(c => c.name === name)) {
        conflicts.push({ name, clientIds: [...matches].map(c => c.clientId) });
      }
    }
  }
  return { clients, byId, byName, byPublicName, conflicts };
}

function resolveClient(directory, input, { clientId = '' } = {}) {
  if (clientId) {
    const client = directory.byId.get(String(clientId));
    return client ? { status: 'matched', matchedBy: 'id', ...client } :
      { status: 'unmatched', query: clean(input), clientId: String(clientId) };
  }
  const normalized = key(input);
  if (!normalized) return { status: 'unmatched', query: '' };
  let matches = directory.byName.get(normalized), matchedBy = 'name';
  if (!matches || !matches.size) {
    matches = directory.byPublicName.get(publicPersonalKey(input));
    matchedBy = 'legacy-public-name';
  }
  if (!matches || !matches.size) return { status: 'unmatched', query: clean(input) };
  if (matches.size > 1) return {
    status: 'ambiguous', query: clean(input),
    candidates: [...matches].map(c => ({
      clientId: c.clientId, officialName: c.officialName, dbaName: c.dbaName,
    })),
  };
  return { status: 'matched', matchedBy, ...[...matches][0] };
}

// Presentation-only resolver. Unknown and ambiguous labels are left intact,
// so an operator must choose the correct account rather than autoassigning it.
function displayNameForInput(directory, input) {
  const match = resolveClient(directory, input);
  return match.status === 'matched' ? match.displayName : clean(input);
}

async function loadClientDirectory() {
  // listItems paginates via @odata.nextLink, unlike older 500-row reads.
  return buildClientDirectory(await listItems(LISTS.CLIENTS));
}

module.exports = { clean, isPublicPricing, toClient, buildClientDirectory,
  resolveClient, displayNameForInput, loadClientDirectory };
