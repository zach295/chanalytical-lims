from pathlib import Path

p = Path('api/src/update-sample.js')
s = p.read_text()

old_sig = "async function updateRadonSheet(siteIdArg, datePrefix, baseId, newLabId, tokenArg, context) {"
new_sig = "async function updateRadonSheet(siteIdArg, datePrefix, baseId, newLabId, tokenArg, context, updates = {}) {"
if old_sig not in s:
    raise SystemExit('updateRadonSheet signature anchor not found')
s = s.replace(old_sig, new_sig, 1)

old_block = """    await fetch(`${wbBase}/worksheets/${wsId}/range(address='A${targetRow}')`,
      { method: 'PATCH', headers: wbHdr, body: JSON.stringify({ values: [[newLabId]] }) });
    if (context) context.log(`[radonSheet] Updated A${targetRow}: ${newLabId}`);
    return { updated: true, row: targetRow };
"""
new_block = """    await fetch(`${wbBase}/worksheets/${wsId}/range(address='A${targetRow}')`,
      { method: 'PATCH', headers: wbHdr, body: JSON.stringify({ values: [[newLabId]] }) });
    if (context) context.log(`[radonSheet] Updated A${targetRow}: ${newLabId}`);

    // Sample Correction can change the reviewed Date/Time Drawn after accession.
    // Keep the Radon Control Sheet collection fields synchronized with those edits.
    if (updates.dateDrawn !== undefined) {
      const drawDateRaw = String(updates.dateDrawn || '').trim();
      let drawDate = drawDateRaw;
      let m = drawDateRaw.match(/^(\\d{4})-(\\d{2})-(\\d{2})$/);
      if (m) drawDate = `${m[2]}/${m[3]}/${m[1]}`;
      else {
        m = drawDateRaw.match(/^(\\d{1,2})[-\\/](\\d{1,2})[-\\/](\\d{2})$/);
        if (m) drawDate = `${String(m[1]).padStart(2,'0')}/${String(m[2]).padStart(2,'0')}/20${m[3]}`;
      }
      const dRes = await fetch(`${wbBase}/worksheets/${wsId}/range(address='E${targetRow}')`,
        { method: 'PATCH', headers: wbHdr, body: JSON.stringify({ values: [[drawDate]] }) });
      if (!dRes.ok) throw new Error(`Radon draw-date update failed (${dRes.status})`);
      if (context) context.log(`[radonSheet] Updated E${targetRow} draw date: ${drawDate}`);
    }

    if (updates.timeDrawn !== undefined) {
      const drawTime = to24h(updates.timeDrawn);
      const tRes = await fetch(`${wbBase}/worksheets/${wsId}/range(address='F${targetRow}')`,
        { method: 'PATCH', headers: wbHdr, body: JSON.stringify({ values: [[drawTime]] }) });
      if (!tRes.ok) throw new Error(`Radon draw-time update failed (${tRes.status})`);
      if (context) context.log(`[radonSheet] Updated F${targetRow} draw time: ${drawTime}`);
    }

    return { updated: true, row: targetRow };
"""
if old_block not in s:
    raise SystemExit('radon update block anchor not found')
s = s.replace(old_block, new_block, 1)

old_call = "const rwResult = await updateRadonSheet(siteId, datePrefix, baseId, finalLabId, token, context);"
new_call = "const rwResult = await updateRadonSheet(siteId, datePrefix, baseId, finalLabId, token, context, updates);"
if old_call not in s:
    raise SystemExit('updateRadonSheet call anchor not found')
s = s.replace(old_call, new_call, 1)

p.write_text(s)
