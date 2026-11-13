'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { buildCorrectionClientOptions } = require('../shared/sample-correction-clients');
const { buildClientDirectory } = require('../shared/client-identity');
const { createReportLabIdDisplayResolver } = require('../shared/report-labid-display');

const clients = [
  { clientId:'1', officialName:'Critical Plumbing Inc. a/k/a Ward Water',dbaName:'Ward Water',displayName:'Ward Water'},
  { clientId:'2', officialName:'First LLC',dbaName:'Shared DBA',displayName:'Shared DBA'},
  { clientId:'3', officialName:'Second LLC',dbaName:'Shared DBA',displayName:'Shared DBA'},
  { clientId:'4', officialName:'No DBA LLC',dbaName:'',displayName:'No DBA LLC'},
];
test('searchable correction picker shows DBA but retains legal name and client ID', () => {
  const options = buildCorrectionClientOptions(clients);
  const ward = options.find(c => c.clientId === '1');
  assert.equal(ward.label,'Ward Water');
  assert.equal(ward.officialName,'Critical Plumbing Inc. a/k/a Ward Water');
  assert.equal(ward.selectable,true);
});
test('duplicate DBAs show disambiguation with official name and ID', () => {
  const options = buildCorrectionClientOptions(clients).filter(c => c.displayName === 'Shared DBA');
  assert.equal(options.length,2);
  assert.notEqual(options[0].label,options[1].label);
  assert.ok(options.every(c=>c.label.includes(c.officialName)&&c.label.includes('ID ')));
});
test('missing DBA falls back to official name', () => {
  const options = buildCorrectionClientOptions(clients);
  assert.equal(options.find(c=>c.clientId==='4').label,'No DBA LLC');
});
test('duplicate official identities cannot be selected as uniquely writable accounts', () => {
  const options = buildCorrectionClientOptions([
    {clientId:'5',officialName:'Same LLC',displayName:'DBA One'},
    {clientId:'6',officialName:'Same LLC',displayName:'DBA Two'},
  ]);
  assert.ok(options.every(c=>c.selectable===false));
});
test('rejection/correction lookup only displays DBA for resolved archived client', () => {
  const directory=buildClientDirectory(clients.map(c=>({
    _id:c.clientId,ClientName:c.officialName,Phone:c.dbaName,
  })));
  const display=createReportLabIdDisplayResolver(directory);
  assert.equal(display('Critical Plumbing Inc a/k/a Ward Water'),'Ward Water');
  assert.equal(display('Unknown archived name'),null);
});
