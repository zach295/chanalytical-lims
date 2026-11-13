'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { matchIntakeClient } = require('../shared/intake-client-match');

const rows = [
  { _id:'1', ClientName:'Critical Plumbing Inc. a/k/a Ward Water', Phone:'Ward Water', ClientCode:'WARD', Abbrev:'WW', PricingCategory:'Inspector Pricing', Aliases:'reports@example.test' },
  { _id:'2', ClientName:'Smith, John', Phone:'John Smith', Abbrev:'PUBLIC', PricingCategory:'Public Pricing' },
  { _id:'3', ClientName:'Public-Jones, Jane', Phone:'Jane Jones', Abbrev:'PUBLIC', PricingCategory:'Public Pricing' },
  { _id:'4', ClientName:'Another Water LLC', Phone:'Other Water', Abbrev:'OW', PricingCategory:'Inspector Pricing' },
];
test('DBA matches official business without changing identity', () => {
  const match = matchIntakeClient('Ward Water',rows);
  assert.equal(match.status,'matched');
  assert.equal(match.officialName,'Critical Plumbing Inc. a/k/a Ward Water');
});
test('client codes and abbreviations match exactly', () => {
  assert.equal(matchIntakeClient('WARD',rows).clientId,'1');
  assert.equal(matchIntakeClient('WW',rows).clientId,'1');
});
test('public clients matched by DBA/legacy name, not shared PUBLIC abbreviation', () => {
  assert.equal(matchIntakeClient('John Smith',rows).officialName,'Smith, John');
  assert.equal(matchIntakeClient('Jones, Jane',rows).officialName,'Public-Jones, Jane');
  assert.equal(matchIntakeClient('PUBLIC',rows).status,'ambiguous');
});
test('report-email column is not a name alias', () => {
  assert.equal(matchIntakeClient('reports@example.test',rows).status,'unmatched');
});
test('duplicate DBAs never choose one account', () => {
  const both = [...rows,{_id:'5',ClientName:'Different Client',Phone:'Ward Water'}];
  assert.equal(matchIntakeClient('Ward Water',both).status,'ambiguous');
  assert.equal(matchIntakeClient('Ward Water',both,{clientId:'1'}).clientId,'1');
});
test('historical hard aliases resolve official names only', () => {
  assert.equal(matchIntakeClient('old ward',rows,{hardAliases:{'old ward':'Critical Plumbing Inc. a/k/a Ward Water'}}).clientId,'1');
});
test('no substring/surname-only autoassignment', () => {
  assert.equal(matchIntakeClient('Plumb',rows).status,'unmatched');
});
