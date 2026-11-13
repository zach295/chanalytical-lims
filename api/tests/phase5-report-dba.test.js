'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { reportAttentionName } = require('../shared/report-client-display');
const { buildClientDirectory, resolveClient } = require('../shared/client-identity');

test('business report prints DBA in Attention but retains legal account metadata', () => {
  const meta = { customer:'Critical Plumbing Inc.',clientName:'Critical Plumbing Inc.',dbaName:'Ward Water' };
  assert.equal(reportAttentionName(meta),'Ward Water');
  assert.equal(meta.clientName,'Critical Plumbing Inc.');
});
test('blank DBA falls back to official client name', () => {
  assert.equal(reportAttentionName({clientName:'Acme Laboratories',dbaName:'   '}),'Acme Laboratories');
});
test('public fallback strips Public prefix and preserves first-last display', () => {
  assert.equal(reportAttentionName({clientName:'Public-Smith, Jane'}),'Jane Smith');
  assert.equal(reportAttentionName({clientName:'Public-Smith, Jane',dbaName:'Jane Smith'}),'Jane Smith');
});
test('unknown archived client remains readable without substituting a different account', () => {
  assert.equal(reportAttentionName({customer:'Original Archived Name'}),'Original Archived Name');
});
test('duplicate DBAs cannot be used for report client identity resolution', () => {
  const d=buildClientDirectory([
    {_id:'1',ClientName:'First Legal LLC',Phone:'Shared DBA'},
    {_id:'2',ClientName:'Second Legal LLC',Phone:'Shared DBA'}
  ]);
  assert.equal(resolveClient(d,'Shared DBA').status,'ambiguous');
  assert.equal(resolveClient(d,'First Legal LLC').status,'matched');
});
