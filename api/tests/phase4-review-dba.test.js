'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { buildClientDirectory, displayNameForInput } = require('../shared/client-identity');

const rows = [
  { _id:'1', ClientName:'Critical Plumbing LLC', Phone:'Ward Water', PricingCategory:'Inspector Pricing' },
  { _id:'2', ClientName:'Another Plumbing LLC', Phone:'Shared DBA', PricingCategory:'Inspector Pricing' },
  { _id:'3', ClientName:'Third Plumbing LLC', Phone:'Shared DBA', PricingCategory:'Inspector Pricing' },
  { _id:'4', ClientName:'Public-Smith, Jane', Phone:'Jane Smith', PricingCategory:'Public Pricing' },
];
const directory = buildClientDirectory(rows);

test('review card displays DBA for an unambiguous official client name', () => {
  assert.equal(displayNameForInput(directory, 'Critical Plumbing LLC'), 'Ward Water');
});
test('review card preserves the DBA if used in the saved queue row', () => {
  assert.equal(displayNameForInput(directory, 'Ward Water'), 'Ward Water');
});
test('ambiguous DBA is preserved and never assigned to one of multiple accounts', () => {
  assert.equal(displayNameForInput(directory, 'Shared DBA'), 'Shared DBA');
});
test('unknown OCR customer is preserved rather than reassigned', () => {
  assert.equal(displayNameForInput(directory, 'Unrecognized Customer'), 'Unrecognized Customer');
});
test('legacy public client displays DBA but retains the underlying official identity', () => {
  assert.equal(displayNameForInput(directory, 'Public-Smith, Jane'), 'Jane Smith');
  assert.equal(directory.clients.find(c => c.clientId === '4').officialName, 'Public-Smith, Jane');
});
