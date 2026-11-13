'use strict';

const { resolveClient, clean } = require('./client-identity');
const { matchIntakeClient } = require('./intake-client-match');

/**
 * Resolve presentation labels for Reports > Pending Lab IDs.
 * The customer stored in Archived Intake is never modified. Unknown or
 * ambiguous accounts return null so the caller can retain its existing label.
 */
function createReportLabIdDisplayResolver(directory) {
  const rows = directory.clients.map(client => ({
    _id: client.clientId,
    ClientName: client.officialName,
    Phone: client.dbaName,
    PricingCategory: client.pricingCategory,
    Abbrev: client.abbrev,
  }));
  const cache = new Map();
  return customer => {
    const raw = clean(customer);
    if (!raw) return null;
    if (cache.has(raw)) return cache.get(raw);

    // Prefer the canonical identity resolver. Only try punctuation-normalized
    // exact matching when no canonical name/DBA was found.
    let match = resolveClient(directory, raw);
    if (match.status === 'unmatched') match = matchIntakeClient(raw, rows);
    const label = match.status === 'matched' ? match.displayName : null;
    cache.set(raw, label);
    return label;
  };
}

module.exports = { createReportLabIdDisplayResolver };
