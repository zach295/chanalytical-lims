from pathlib import Path

p = Path('api/src/patch-report-cell.js')
s = p.read_text()

old_decl = """      var allPatchReqs = [];
      var newHex = null;
      var hardnessUpdate = null;
"""
new_decl = """      var allPatchReqs = [];
      var newHex = null;
      var hardnessUpdate = null;
      var oldValue = null; // original result value captured before any workbook write
"""
if old_decl not in s:
    raise SystemExit('declaration block not found')
s = s.replace(old_decl, new_decl, 1)

old_target = """        if (targetRow < 0) continue; // param not in this sheet, skip

        // Build writes for this sheet
"""
new_target = """        if (targetRow < 0) continue; // param not in this sheet, skip

        // Capture the original value before patching. Use the first matching sheet as
        // the audit source so Lab Report/FHA duplicate rows do not overwrite it.
        if (field === 'value' && oldValue === null && colResult >= 0) {
          const originalCell = (rows[targetRow] || [])[colResult];
          oldValue = originalCell === undefined || originalCell === null ? '' : String(originalCell);
        }

        // Build writes for this sheet
"""
if old_target not in s:
    raise SystemExit('target block not found')
s = s.replace(old_target, new_target, 1)

old_notes = """            Notes:        `${paramName} changed to \"${value}\"`,
"""
new_notes = """            Notes:        `${paramName} changed from \"${oldValue ?? ''}\" to \"${value}\"`,
"""
if old_notes not in s:
    raise SystemExit('notes line not found')
s = s.replace(old_notes, new_notes, 1)

p.write_text(s)
