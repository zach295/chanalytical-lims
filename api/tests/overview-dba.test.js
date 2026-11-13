'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { buildClientDirectory } = require('../shared/client-identity');
const { projectOverviewRecords, countOverviewClients } = require('../shared/overview-dba');

const dir = buildClientDirectory([
  { _id:'1', ClientName:'Critical Plumbing Inc. a/k/a Ward Water', Phone:'Ward Water' },
  { _id:'2', ClientName:'Second Company', Phone:'Ward Water' },
  { _id:'3', ClientName:'Smith, John', Phone:'John Smith', PricingCategory:'Public Pricing' },
  { _id:'4', ClientName:'No DBA LLC', Phone:'' },
  { _id:'5', ClientName:'Public-Jones, Jane', Phone:'Jane Jones', PricingCategory:'Public Pricing' },
]);

test('DBA display leaves intake fields and test totals unchanged', () => {
  const raw = [{ field_1:'111126-001 COMP', field_2:'Comprehensive', field_3:'No DBA LLC', field_6:'11/11/26', field_14:'Pending' }];
  const records = projectOverviewRecords(raw, dir);
  assert.equal(records.length, raw.length);
  assert.equal(records[0].customer, 'No DBA LLC');
  assert.equal(records[0].labId, raw[0].field_1);
  assert.equal(raw[0].field_3, 'No DBA LLC');
});

test('existing business gets DBA and official identity', () => {
  const records = projectOverviewRecords([{field_1:'111126-001',field_3:'Critical Plumbing Inc. a/k/a Ward Water'}],dir);
  assert.equal(records[0].customer, 'Ward Water');
  assert.equal(records[0].officialCustomer, 'Critical Plumbing Inc. a/k/a Ward Water');
  assert.equal(records[0].clientKey, 'id:1');
});

test('same DBA on two accounts does not mix records', () => {
  const rows = [
    {field_1:'111126-001',field_3:'Critical Plumbing Inc. a/k/a Ward Water'},
    {field_1:'111126-002',field_3:'Second Company'},
  ];
  const recs = projectOverviewRecords(rows,dir);
  assert.equal(countOverviewClients(recs).length, 2);
  assert.equal(recs[0].customer, recs[1].customer);
});

test('legacy public names map to DBA, but remain unchanged in records', () => {
  const raw = [{field_1:'111126-001',field_3:'Public-Jones, Jane'}];
  const records = projectOverviewRecords(raw,dir);
  assert.equal(records[0].customer, 'Jane Jones');
  assert.equal(records[0].officialCustomer, 'Public-Jones, Jane');
  assert.equal(raw[0].field_3, 'Public-Jones, Jane');
});

test('new public names map to DBA', () => {
  const records = projectOverviewRecords([{field_1:'111126-001',field_3:'Smith, John'}],dir);
  assert.equal(records[0].customer, 'John Smith');
});

test('unknown and ambiguous names remain unmodified', () => {
  const records = projectOverviewRecords([
    {field_1:'111126-001',field_3:'Unlisted Person'},
    {field_1:'111126-002',field_3:'Ward Water'},
  ],dir);
  assert.equal(records[0].customer, 'Unlisted Person');
  assert.equal(records[1].customer, 'Ward Water');
  assert.equal(records[1].clientMatchStatus, 'ambiguous');
});

test('counts retain rejected and nonrejected records for existing overview client table', () => {
  const records = projectOverviewRecords([
    {field_1:'111126-001',field_3:'No DBA LLC',field_14:'Rejected'},
    {field_1:'111126-002',field_3:'No DBA LLC',field_14:'Pending'},
  ],dir);
  assert.equal(countOverviewClients(records)[0].total, 2);
});
