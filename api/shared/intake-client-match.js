'use strict';

/**
 * Phase 4: identity-safe client recognition for COC intake and Review Queue.
 * Never select an arbitrary account when a DBA, abbreviation or client code
 * matches multiple records. "Aliases" is currently a REPORT EMAIL field in
 * SharePoint; it must not be used for customer-name matching.
 */
const { buildClientDirectory, resolveClient, clean } = require('./client-identity');

const norm = v => clean(v).toLowerCase();
const punctuationKey = v => norm(v).replace(/[.,'"()\-]/g, '').replace(/\s+/g, ' ').trim();

function matchIntakeClient(input, rows, { clientId = '', hardAliases = {} } = {}) {
  const directory = buildClientDirectory(rows);
  const query = clean(input);
  const found = resolveClient(directory, query, { clientId });
  if (found.status === 'matched' || found.status === 'ambiguous' || clientId) return found;
  if (!query) return found;
  const mapped = hardAliases[norm(query)] || hardAliases[punctuationKey(query)];
  if (mapped) {
    const aliasHit = resolveClient(directory, mapped);
    if (aliasHit.status !== 'unmatched') return aliasHit;
  }

  // Match exact client codes, abbreviations or punctuation-normalized legal
  // names/DBAs. Do not use contains/first surname/fuzzy matching for identity.
  const q = punctuationKey(query);
  const candidates = directory.clients.filter(c => {
    const r = rows.find(row => String(row.clientId ?? row._id ?? row.id ?? '') === c.clientId &&
      norm(row.clientName ?? row.ClientName ?? row.Title) === norm(c.officialName));
    const labels = [c.officialName, c.dbaName, r?.ClientCode, r?.clientCode, r?.Abbrev, r?.abbrev];
    return labels.some(x => x && punctuationKey(x) === q);
  });
  if (candidates.length === 1) return { status: 'matched', matchedBy: 'exact-identifier', ...candidates[0] };
  if (candidates.length > 1) return { status: 'ambiguous', query,
    candidates: candidates.map(c => ({clientId:c.clientId,officialName:c.officialName,dbaName:c.dbaName})) };
  return { status: 'unmatched', query };
}

module.exports = { matchIntakeClient };
