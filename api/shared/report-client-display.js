'use strict';
const { clean } = require('./client-identity');

// Only the report Attention line uses the DBA. Customer/account identity,
// COA billing rows, emails, audit data, and historical records stay unchanged.
function reportAttentionName(meta = {}) {
  const official = clean(meta.clientName || meta.customer);
  const dba = clean(meta.dbaName);
  if (dba) return dba;
  const stripped = official.replace(/^Public-/i, '').trim();
  return /^Public-/i.test(official) && stripped.includes(',')
    ? stripped.replace(/^([^,]+),\s*(.+)$/, '$2 $1')
    : stripped;
}

module.exports = { reportAttentionName };
