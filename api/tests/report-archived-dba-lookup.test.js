'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { matchIntakeClient } = require('../shared/intake-client-match');
const { reportAttentionName } = require('../shared/report-client-display');

// Regression: Archived Intake may have older punctuation, while Clients has
// the current legal name and the DBA stored in internal SharePoint Phone.
const clients = [
  { _id:'42', ClientName:'Critical Plumbing Inc. a/k/a Ward Water', Phone:'Ward Water', PricingCategory:'Inspector Pricing' },
  { _id:'43', ClientName:'Another Client LLC', Phone:'Another Client', PricingCategory:'Inspector Pricing' }
];
test('today-approved archived names resolve despite punctuation variations', () => {
  const match = matchIntakeClient('Critical Plumbing Inc a/k/a Ward Water', clients);
  assert.equal(match.status,'matched');
  assert.equal(match.clientId,'42');
  assert.equal(match.dbaName,'Ward Water');
  assert.equal(reportAttentionName({clientName:match.officialName,dbaName:match.dbaName}),'Ward Water');
});
test('unknown archived names never inherit another company DBA', () => {
  assert.equal(matchIntakeClient('Critical Plumbing',clients).status,'unmatched');
});
test('ambiguous DBA cannot auto-assign to a client for a report', () => {
  const duplicate=[...clients,{_id:'44',ClientName:'Unrelated Legal LLC',Phone:'Ward Water'}];
  assert.equal(matchIntakeClient('Ward Water',duplicate).status,'ambiguous');
});
