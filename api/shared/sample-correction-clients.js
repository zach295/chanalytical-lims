'use strict';
const { clean } = require('./client-identity');

function buildCorrectionClientOptions(clients) {
  const list = (clients || []).filter(c => clean(c.officialName));
  const labelCounts = new Map(), officialCounts = new Map();
  const key = value => clean(value).toLocaleLowerCase('en-US');
  for (const c of list) {
    const display = clean(c.displayName || c.dbaName || c.officialName);
    labelCounts.set(key(display), (labelCounts.get(key(display)) || 0) + 1);
    officialCounts.set(key(c.officialName), (officialCounts.get(key(c.officialName)) || 0) + 1);
  }
  return list.map(c => {
    const officialName = clean(c.officialName);
    const displayName = clean(c.displayName || c.dbaName || officialName);
    const duplicateDBA = labelCounts.get(key(displayName)) > 1;
    const ambiguousOfficial = officialCounts.get(key(officialName)) > 1;
    return {
      clientId: String(c.clientId || ''),
      officialName,
      displayName,
      label: duplicateDBA
        ? displayName + ' (' + officialName + ' — ID ' + String(c.clientId || '?') + ')'
        : displayName,
      selectable: !ambiguousOfficial,
    };
  }).sort((a,b) => a.label.localeCompare(b.label));
}

module.exports = { buildCorrectionClientOptions };
