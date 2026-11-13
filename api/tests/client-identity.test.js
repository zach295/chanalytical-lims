'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { buildClientDirectory, resolveClient, isPublicPricing } = require('../shared/client-identity');

const clients = [
  { _id: '1', ClientName: 'Critical Plumbing Inc. a/k/a Ward Water', Phone: 'Ward Water', PricingCategory: 'Inspector Pricing' },
  { _id: '2', ClientName: 'Smith, John', Phone: 'John Smith', PricingCategory: 'Public Pricing' },
  { _id: '3', ClientName: 'Public-Jones, Jane', Phone: 'Jane Jones', PricingCategory: 'Public Pricing' },
  { _id: '4', ClientName: 'No DBA LLC', Phone: '', PricingCategory: 'Inspector Pricing' },
];
test('maps DBA from SharePoint Phone while preserving official account', () => {
  const r = resolveClient(buildClientDirectory(clients), 'Ward Water');
  assert.equal(r.status, 'matched');
  assert.equal(r.officialName, 'Critical Plumbing Inc. a/k/a Ward Water');
  assert.equal(r.displayName, 'Ward Water');
  assert.equal(r.clientId, '1');
});
test('official client name resolves to the same account', () => {
  assert.equal(resolveClient(buildClientDirectory(clients), 'Critical Plumbing Inc. a/k/a Ward Water').clientId, '1');
});
test('blank DBA falls back to official name', () => {
  assert.equal(resolveClient(buildClientDirectory(clients), 'No DBA LLC').displayName, 'No DBA LLC');
});
test('new public clients resolve using DBA or Last, First', () => {
  const r = resolveClient(buildClientDirectory(clients), 'John Smith');
  assert.equal(r.officialName, 'Smith, John');
  assert.ok(isPublicPricing(r.pricingCategory));
});
test('legacy public prefix resolves without changing stored name', () => {
  const r = resolveClient(buildClientDirectory(clients), 'Jones, Jane');
  assert.equal(r.officialName, 'Public-Jones, Jane');
  assert.equal(resolveClient(buildClientDirectory(clients), 'Jane Jones').clientId, '3');
});
test('duplicate DBA results are ambiguous, but stable ID resolves', () => {
  const d = buildClientDirectory([...clients, {
    _id: '5', ClientName: 'Other Plumbing', Phone: 'Ward Water', PricingCategory: 'Inspector Pricing',
  }]);
  assert.equal(resolveClient(d, 'Ward Water').status, 'ambiguous');
  assert.ok(d.conflicts.some(c => c.name === 'ward water'));
  assert.equal(resolveClient(d, 'Ward Water', { clientId: '1' }).clientId, '1');
});
test('unknown names never match by substring', () => {
  assert.equal(resolveClient(buildClientDirectory(clients), 'Ward').status, 'unmatched');
});
test('duplicate legacy public names are not guessed', () => {
  const d = buildClientDirectory([
    { _id: '3', ClientName: 'Public-Jones, Jane', PricingCategory: 'Public Pricing' },
    { _id: '6', ClientName: 'Jones, Jane', PricingCategory: 'Public Pricing' },
  ]);
  assert.equal(resolveClient(d, 'Jane Jones').status, 'ambiguous');
});
