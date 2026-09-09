const fs = require('fs');

function patchFile(path, edits) {
  let s = fs.readFileSync(path, 'utf8');
  for (const [oldText, newText, label] of edits) {
    if (!s.includes(oldText)) throw new Error(`${path}: anchor not found for ${label}`);
    s = s.replace(oldText, newText);
  }
  fs.writeFileSync(path, s);
}

patchFile('api/src/approve-scan.js', [
  [
`        const sheetRows = labItems
          .filter(l => !l.isRejected)
          .flatMap(l => {`,
`        const sheetRows = labItems
          .flatMap(l => {`,
'include rejected samples in COA write'
  ],
  [
`              '', // Report Date is written only when the report is actually sent
              l.baseId,`,
`              (l.isRejection || l.isRejected) ? toCoaDate(nextBusinessDay(dateRec)) : '', // rejected samples use next business day; normal samples are dated when sent
              l.baseId,`,
'rejected COA report date'
  ],
]);

patchFile('api/src/update-sample.js', [
  [
`const { listItems, updateItem, createItem, getToken, LISTS } = require('../shared/graph');`,
`const { listItems, updateItem, createItem, getToken, LISTS } = require('../shared/graph');
const { syncCoaFromArchivedRows } = require('../shared/coa-sheet');`,
'COA helper import'
  ],
  [
`      // Results Cache: lab ID stores base ID only — no suffix — no update needed on correction

      // ── Write to Activity Log ───────────────────────────────────────────────`,
`      // Results Cache: lab ID stores base ID only — no suffix — no update needed on correction

      // ── Synchronize COA / Form Responses ───────────────────────────────────
      // Re-read Archived Intake after all correction writes so the COA sheet mirrors
      // the final authoritative values rather than the pre-correction snapshot.
      try {
        const refreshed = await listItems(LISTS.ARCHIVED_INTAKE, { top: 2000 });
        const correctedRows = refreshed.filter(r => (r.field_1 || '').split(' ')[0].trim() === baseId);
        const coaSync = await syncCoaFromArchivedRows(baseId, correctedRows, context);
        log.push(\`✅ COA sheet synchronized (\${coaSync.updated} row(s)\${coaSync.cleared ? \`, \${coaSync.cleared} old row(s) cleared\` : ''})\`);
      } catch(e) {
        log.push(\`⚠️ COA sheet: \${e.message}\`);
      }

      // ── Write to Activity Log ───────────────────────────────────────────────`,
'correction COA sync'
  ],
]);

patchFile('api/src/reject-sample.js', [
  [
`const { createItem, listItems, updateItem, LISTS } = require('../shared/graph');`,
`const { createItem, listItems, updateItem, LISTS } = require('../shared/graph');
const { syncCoaFromArchivedRows } = require('../shared/coa-sheet');`,
'COA helper import'
  ],
  [
`      log.push(\`✅ Archived Intake: updated \${matches.length} row(s) → suffix changed to REJ\`);

      return {`,
`      log.push(\`✅ Archived Intake: updated \${matches.length} row(s) → suffix changed to REJ\`);

      // ── Synchronize COA / Form Responses ─────────────────────────────────
      // This updates the existing COA row(s) to the rejection type and assigns the
      // next-business-day Report Date. If an older sample never had a COA row, the
      // helper uses an existing blank sheet row rather than inserting duplicates.
      try {
        const refreshed = await listItems(LISTS.ARCHIVED_INTAKE, { top: 2000 });
        const rejectedRows = refreshed.filter(r => (r.field_1 || '').split(' ')[0].trim() === baseId);
        const coaSync = await syncCoaFromArchivedRows(baseId, rejectedRows, context);
        log.push(\`✅ COA sheet synchronized to rejection (\${coaSync.updated} row(s))\`);
      } catch(e) {
        log.push(\`⚠️ COA sheet: \${e.message}\`);
      }

      return {`,
'rejection COA sync'
  ],
]);

console.log('Patched COA synchronization for corrections and rejections');
