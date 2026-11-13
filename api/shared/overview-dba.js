'use strict';

/**
 * Phase 2: DBA display projection for Overview statistics only.
 * Does not update Archived Intake or change the original customer field.
 */
const { resolveClient, clean } = require('./client-identity');

function projectOverviewRecords(intakeRows, clientDirectory) {
  return intakeRows.map(row => {
    const rawCustomer = clean(row.field_3);
    const resolution = resolveClient(clientDirectory, rawCustomer);
    const matched = resolution.status === 'matched';
    const displayName = matched ? resolution.displayName : (rawCustomer || 'Unknown');
    // Distinct official client IDs must never merge just because DBA labels match.
    // Unresolved or ambiguous records retain their exact historical identity.
    const clientKey = matched && resolution.clientId
      ? 'id:' + resolution.clientId
      : 'raw:' + rawCustomer.toLocaleLowerCase('en-US');

    return {
      labId: clean(row.field_1),
      testType: clean(row.field_2),
      customer: displayName,
      officialCustomer: matched ? resolution.officialName : rawCustomer,
      clientKey,
      clientMatchStatus: resolution.status,
      dateReceived: clean(row.field_6),
      status: clean(row.field_14) || 'Pending',
    };
  }).filter(record => record.labId);
}

function countOverviewClients(records) {
  const byKey = new Map();
  for (const record of records) {
    const key = record.clientKey;
    const row = byKey.get(key) || { name: record.customer, officialCustomer: record.officialCustomer, total: 0 };
    row.total++;
    byKey.set(key, row);
  }
  return [...byKey.values()];
}

module.exports = { projectOverviewRecords, countOverviewClients };
