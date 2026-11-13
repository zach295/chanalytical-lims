'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { buildClientDirectory } = require('../shared/client-identity');
const { createReportLabIdDisplayResolver } = require('../shared/report-labid-display');

const directory = buildClientDirectory([
  { _id: '1', ClientName: 'Critical Plumbing Inc. a/k/a Ward Water', Phone: 'Ward Water', PricingCategory: 'Inspector Pricing' },
  { _id: '2', ClientName: 'Smith, Jane', Phone: 'Jane Smith', PricingCategory: 'Public Pricing' },
  { _id: '3', ClientName: 'Client One LLC', Phone: 'Shared DBA', PricingCategory: 'Inspector Pricing' },
  { _id: '4', ClientName: 'Client Two LLC', Phone: 'Shared DBA', PricingCategory: 'Inspector Pricing' },
]);
const display = createReportLabIdDisplayResolver(directory);

test('report Lab ID dropdown shows DBA for archived official client', () => {
  assert.equal(display('Critical Plumbing Inc. a/k/a Ward Water'), 'Ward Water');
});
test('report Lab ID dropdown resolves an archived punctuation variation', () => {
  assert.equal(display('Critical Plumbing Inc a/k/a Ward Water'), 'Ward Water');
});
test('public client displays DBA rather than last-first legal label', () => {
  assert.equal(display('Smith, Jane'), 'Jane Smith');
});
test('unknown and ambiguous names are not silently assigned to another account', () => {
  assert.equal(display('Unknown Company'), null);
  assert.equal(display('Shared DBA'), null);
});
test('same client association retains official name in the directory', () => {
  assert.equal(directory.byId.get('1').officialName, 'Critical Plumbing Inc. a/k/a Ward Water');
});
