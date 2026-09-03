from pathlib import Path


def replace(path, old, new, label):
    p = Path(path)
    s = p.read_text()
    if old not in s:
        raise SystemExit(f'{label} not found in {path}')
    p.write_text(s.replace(old, new, 1))

# USERS: use shared audit helper and log deactivation/password actions.
replace('api/src/users-manage.js',
"const { listItems, createItem, updateItem, LISTS } = require('../shared/graph');\n",
"const { listItems, createItem, updateItem, LISTS } = require('../shared/graph');\nconst { writeActivityLog } = require('../shared/audit');\n",
'users audit import')

old_func = '''async function logUserActivity(action, email, details, by) {
  try {
    const { createItem } = require('../shared/graph');
    const now     = new Date();
    const logDate = now.toLocaleDateString('en-US',{ timeZone:'America/New_York', month:'2-digit', day:'2-digit', year:'2-digit' });
    const logTime = now.toLocaleTimeString('en-US',{ timeZone:'America/New_York', hour:'2-digit', minute:'2-digit', hour12:false });
    await createItem('Activity Log', {
      Title: `${logDate} ${email}`, Client: email,
      ActivityType: action, Notes: (details || '').slice(0, 3000),
      By: by || 'Admin', LogDate: logDate, LogTime: logTime, Quantity: 0,
    });
  } catch(e) {}
}
'''
new_func = '''async function logUserActivity(action, email, details, by, context) {
  return writeActivityLog({
    labId: email,
    type: action,
    notes: details || '',
    by: by || 'Admin',
    context,
  });
}
'''
replace('api/src/users-manage.js', old_func, new_func, 'users log helper')

# Existing calls gain context so failures are visible in function logs.
p = Path('api/src/users-manage.js')
s = p.read_text()
s = s.replace("await logUserActivity('User Created', email, `New account created in Users list.\\nRole: ${role || 'lab'} | Created by: ${createdBy || 'Admin'}`, createdBy);",
              "await logUserActivity('User Created', email, `New account created in Users list.\\nRole: ${role || 'lab'} | Created by: ${createdBy || 'Admin'}`, createdBy, context);")
s = s.replace("await logUserActivity('User Edited', email, `User account updated in Users list.\\nNew email: ${newEmail || email} | Name: ${name || '—'} | Role: ${role || '—'}`, 'Admin');",
              "await logUserActivity('User Edited', email, `User account updated in Users list.\\nNew email: ${newEmail || email} | Name: ${name || '—'} | Role: ${role || '—'}`, body.updatedBy || 'Admin', context);")
s = s.replace("await logUserActivity('Role Changed', srEmail, `Role updated in Users list.\\nNew role: ${srRole || '—'}`, 'Admin');",
              "await logUserActivity('Role Changed', srEmail, `Role: \"${user.field_2 || ''}\" → \"${srRole || '—'}\"`, body.updatedBy || 'Admin', context);")

old_deactivate = '''        await updateItem(LISTS.USERS, user._id, { field_2: 'deactivated', field_9: false });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
'''
new_deactivate = '''        await updateItem(LISTS.USERS, user._id, { field_2: 'deactivated', field_9: false });
        const audit = await logUserActivity('User Deactivated', email, `Role: \"${user.field_2 || ''}\" → \"deactivated\" | Active: true → false`, body.updatedBy || 'Admin', context);
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true, auditWarning: audit.success ? null : audit.error }) };
'''
if old_deactivate not in s: raise SystemExit('users deactivate block not found')
s = s.replace(old_deactivate, new_deactivate, 1)

old_setpw = '''        const hashed = hashPassword(newPw || '');
        await updateItem(LISTS.USERS, user._id, { field_8: hashed, field_7: false, field_9: true });
        const mapped = mapUser(user);
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({
          success: true,
          user: { email: mapped.email, name: mapped.name, role: mapped.role, clientKey: mapped.clientKey },
        })};
'''
new_setpw = '''        const hashed = hashPassword(newPw || '');
        await updateItem(LISTS.USERS, user._id, { field_8: hashed, field_7: false, field_9: true });
        const mapped = mapUser(user);
        const audit = await logUserActivity('Password Changed', pwEmail, 'Password changed. Password value/hash intentionally not recorded.', body.updatedBy || mapped.name || 'User', context);
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({
          success: true,
          auditWarning: audit.success ? null : audit.error,
          user: { email: mapped.email, name: mapped.name, role: mapped.role, clientKey: mapped.clientKey },
        })};
'''
if old_setpw not in s: raise SystemExit('users setpw block not found')
s = s.replace(old_setpw, new_setpw, 1)

old_reset = '''        await updateItem(LISTS.USERS, user._id, { field_8: WELCOME_HASH, field_7: true, field_9: true });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
'''
new_reset = '''        await updateItem(LISTS.USERS, user._id, { field_8: WELCOME_HASH, field_7: true, field_9: true });
        const audit = await logUserActivity('Password Reset', pwEmail, 'Password reset to temporary welcome credential; user must change it on next login. Password value/hash intentionally not recorded.', body.updatedBy || 'Admin', context);
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true, auditWarning: audit.success ? null : audit.error }) };
'''
if old_reset not in s: raise SystemExit('users resetpw block not found')
s = s.replace(old_reset, new_reset, 1)
p.write_text(s)

# MARK SCAN: discarded records/files are destructive and must be audited.
replace('api/src/mark-scan-processed.js',
"const { updateItem, deleteItem, getToken, LISTS } = require('../shared/graph');\n",
"const { updateItem, deleteItem, getToken, LISTS } = require('../shared/graph');\nconst { writeActivityLog } = require('../shared/audit');\n",
'mark scan audit import')
replace('api/src/mark-scan-processed.js',
"      const { fileId, outcome, reviewQueueRow, rowIndex } = await request.json();\n",
"      const { fileId, outcome, reviewQueueRow, rowIndex, processedBy, labId, fileName } = await request.json();\n",
'mark scan body')
replace('api/src/mark-scan-processed.js',
'''      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, driveDeleted: !!fileId, row, outcome }),
      };
''',
'''      let auditWarning = null;
      if (outcome === 'discarded') {
        const audit = await writeActivityLog({
          labId: labId || fileName || `Scan ${row}`,
          type: 'Scan Discarded',
          notes: `Review Queue row ${row} discarded${fileName ? ` | File: ${fileName}` : ''}${fileId ? ` | File ID: ${fileId}` : ''}${fileId ? ' | Underlying file deletion requested' : ''}`,
          by: processedBy || 'Lab Staff',
          context,
        });
        if (!audit.success) auditWarning = audit.error;
      }

      return {
        status: 200,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ success: true, driveDeleted: !!fileId, row, outcome, auditWarning }),
      };
''',
'mark scan return')

# PH IMPORT: one audit event per successfully imported sample.
replace('api/src/import-ph.js',
"const { getToken } = require('../shared/graph');\n",
"const { getToken } = require('../shared/graph');\nconst { writeActivityLog } = require('../shared/audit');\n",
'ph audit import')
replace('api/src/import-ph.js',
"      let totalUpdated = 0;\n      const log = [];\n",
"      let totalUpdated = 0;\n      const log = [];\n      const importedSamples = [];\n",
'ph imported list')
replace('api/src/import-ph.js',
'''          if (patch.ok) {
            totalUpdated++;
            log.push(`✅ ${labId}: pH=${result.ph}`);
          } else {
''',
'''          if (patch.ok) {
            totalUpdated++;
            log.push(`✅ ${labId}: pH=${result.ph}`);
            importedSamples.push({ labId, notes: `pH=${result.ph} | Analysis date/time: ${toMilitaryDT(result.dt)}` });
          } else {
''',
'ph success')
replace('api/src/import-ph.js',
"      return { status: 200, jsonBody: { success: true, updated: totalUpdated, log } };\n",
'''      const actor = body.importedBy || body.updatedBy || 'Lab Staff';
      const auditWarnings = [];
      for (const sample of importedSamples) {
        const audit = await writeActivityLog({ labId: sample.labId, type: 'Results Imported - pH', notes: sample.notes, by: actor, context });
        if (!audit.success) auditWarnings.push(`${sample.labId}: ${audit.error}`);
      }
      return { status: 200, jsonBody: { success: true, updated: totalUpdated, log, auditWarnings } };
''',
'ph return')

# BACTERIA IMPORT.
replace('api/src/import-bacteria.js',
"const { getToken } = require('../shared/graph');\n",
"const { getToken } = require('../shared/graph');\nconst { writeActivityLog } = require('../shared/audit');\n",
'bac audit import')
replace('api/src/import-bacteria.js',
"      let totalUpdated = 0;\n      const log = [];\n",
"      let totalUpdated = 0;\n      const log = [];\n      const importedSamples = [];\n",
'bac imported list')
replace('api/src/import-bacteria.js',
'''          if (patch.ok) {
            totalUpdated++;
            log.push(`✅ ${labId}: coliform=${result.coliform} ecoli=${result.ecoli}`);
          } else {
''',
'''          if (patch.ok) {
            totalUpdated++;
            log.push(`✅ ${labId}: coliform=${result.coliform} ecoli=${result.ecoli}`);
            importedSamples.push({ labId, notes: `Total Coliform=${result.coliform} | E. Coli=${result.ecoli} | Start=${toMilitaryDT(result.timeIn)} | End=${toMilitaryDT(result.timeOut)} | Source: ${bacName}` });
          } else {
''',
'bac success')
replace('api/src/import-bacteria.js',
"      return { status: 200, jsonBody: { success: true, updated: totalUpdated, log } };\n",
'''      const actor = body.importedBy || body.updatedBy || 'Lab Staff';
      const auditWarnings = [];
      for (const sample of importedSamples) {
        const audit = await writeActivityLog({ labId: sample.labId, type: 'Results Imported - Bacteria', notes: sample.notes, by: actor, context });
        if (!audit.success) auditWarnings.push(`${sample.labId}: ${audit.error}`);
      }
      return { status: 200, jsonBody: { success: true, updated: totalUpdated, log, auditWarnings } };
''',
'bac return')

# RADON IMPORT.
replace('api/src/import-radon.js',
"const { getToken } = require('../shared/graph');\n",
"const { getToken } = require('../shared/graph');\nconst { writeActivityLog } = require('../shared/audit');\n",
'radon audit import')
replace('api/src/import-radon.js',
"      let totalUpdated = 0;\n      const log = [];\n",
"      let totalUpdated = 0;\n      const log = [];\n      const importedSamples = [];\n",
'radon imported list')
replace('api/src/import-radon.js',
'''          totalUpdated++;
          log.push(`✅ ${labId}: ${rcsRow.result} pCi/L`);
''',
'''          totalUpdated++;
          log.push(`✅ ${labId}: ${rcsRow.result} pCi/L`);
          importedSamples.push({ labId, notes: `Radon Water=${rcsRow.result} pCi/L | Date tested: ${rcsRow.dateTested || '—'} | Time tested: ${rcsRow.timeTested || '—'} | Source: ${rcsName}` });
''',
'radon success')
replace('api/src/import-radon.js',
'''      return {
        status: 200,
        jsonBody: { success: true, updated: totalUpdated, log },
      };
''',
'''      const actor = body.importedBy || body.updatedBy || 'Lab Staff';
      const auditWarnings = [];
      for (const sample of importedSamples) {
        const audit = await writeActivityLog({ labId: sample.labId, type: 'Results Imported - Radon', notes: sample.notes, by: actor, context });
        if (!audit.success) auditWarnings.push(`${sample.labId}: ${audit.error}`);
      }
      return {
        status: 200,
        jsonBody: { success: true, updated: totalUpdated, log, auditWarnings },
      };
''',
'radon return')
