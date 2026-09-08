from pathlib import Path

p = Path('api/src/approve-scan.js')
s = p.read_text()

s = s.replace(
"""      let csWarning = '';
      try {
""",
"""      let csWarning = '';
      const csOverallStart = Date.now();
      context.log(`[CS][timing] Start control sheet step for ${allFullIds.join(', ')}`);
      try {
""",
1)

start_marker = """        if (csFileId) {
          const wbBase2  = `${GRAPH}/sites/${_siteId}/drive/items/${csFileId}/workbook`;
"""
end_marker = """        } else if (!csWarning) {
          csWarning = `Control sheet ${csFileName} not found — lab IDs not written. Create it and add manually.`;
        }
"""

start = s.find(start_marker)
end = s.find(end_marker, start)
if start < 0 or end < 0:
    raise SystemExit('Control Sheet block markers not found')
end += len(end_marker)

new_block = r'''        if (csFileId) {
          const wbBase2  = `${GRAPH}/sites/${_siteId}/drive/items/${csFileId}/workbook`;

          // Excel/Graph can intermittently report workbook-busy or take longer to
          // commit a write. Keep this step in its existing place in the approval
          // flow, but retry the workbook session and verify the Lab IDs after writing.
          let sid2 = '';
          let sessionLastError = '';
          for (let attempt = 1; attempt <= 3 && !sid2; attempt++) {
            const sessionStart = Date.now();
            try {
              const ses2Res = await fetch(`${wbBase2}/createSession`, {
                method:'POST', headers:{ Authorization:`Bearer ${_token}`, 'Content-Type':'application/json' },
                body: JSON.stringify({ persistChanges:true }),
              });
              const ses2Data = await ses2Res.json().catch(()=>({}));
              context.log(`[CS][timing] createSession attempt ${attempt}: ${Date.now()-sessionStart}ms status=${ses2Res.status}`);
              if (ses2Res.ok && ses2Data.id) {
                sid2 = ses2Data.id;
                break;
              }
              sessionLastError = `createSession ${ses2Res.status}`;
            } catch (sessionErr) {
              sessionLastError = sessionErr.message;
              context.log(`[CS] createSession attempt ${attempt} error: ${sessionErr.message}`);
            }
            if (attempt < 3) await new Promise(r => setTimeout(r, 750 * attempt));
          }
          if (!sid2) throw new Error(`Workbook session failed after 3 attempts: ${sessionLastError || 'unknown error'}`);

          const wbHdr2 = { Authorization:`Bearer ${_token}`, 'workbook-session-id':sid2, 'Content-Type':'application/json' };
          try {
            const worksheetStart = Date.now();
            const sheetsRes2 = await fetch(`${wbBase2}/worksheets`, { headers:wbHdr2 });
            if (!sheetsRes2.ok) throw new Error(`Worksheet list failed (${sheetsRes2.status})`);
            const wsId2 = ((await sheetsRes2.json()).value||[])[0]?.id;
            context.log(`[CS][timing] worksheet lookup: ${Date.now()-worksheetStart}ms status=${sheetsRes2.status}`);
            if (!wsId2) throw new Error('No worksheet found in Control Sheet');

            // Scan column A for existing IDs and first empty row.
            const initialReadStart = Date.now();
            const colARes2 = await fetch(`${wbBase2}/worksheets/${wsId2}/range(address='A1:A500')?$select=values`, { headers:wbHdr2 });
            if (!colARes2.ok) throw new Error(`Control Sheet column read failed (${colARes2.status})`);
            const colAVals2 = (await colARes2.json()).values || [];
            context.log(`[CS][timing] initial A-column read: ${Date.now()-initialReadStart}ms status=${colARes2.status}`);
            const existing = new Set(colAVals2.map(r => String(r[0]||'').trim()).filter(Boolean));
            let nextRow2 = 2;
            for (let i = 1; i < colAVals2.length; i++) {
              if (!String(colAVals2[i][0]||'').trim()) { nextRow2 = i+1; break; }
              nextRow2 = i+2;
            }
            const newIds = allFullIds.filter(id => !existing.has(id.trim()));
            context.log(`[CS] existing=${existing.size} toWrite=${newIds.length} nextRow=${nextRow2}`);

            if (newIds.length) {
              const endRow2 = nextRow2 + newIds.length - 1;
              let verified = false;
              let lastWriteError = '';

              for (let attempt = 1; attempt <= 3 && !verified; attempt++) {
                const patchStart = Date.now();
                try {
                  const patchRes2 = await fetch(`${wbBase2}/worksheets/${wsId2}/range(address='A${nextRow2}:A${endRow2}')`, {
                    method:'PATCH', headers:wbHdr2,
                    body: JSON.stringify({ values: newIds.map(id=>[id]) }),
                  });
                  context.log(`[CS][timing] PATCH attempt ${attempt}: ${Date.now()-patchStart}ms status=${patchRes2.status}`);
                  if (!patchRes2.ok) {
                    const patchText = await patchRes2.text().catch(()=>'');
                    lastWriteError = `PATCH ${patchRes2.status}: ${patchText.slice(0,200)}`;
                    context.log(`[CS] attempt ${attempt} failed: ${lastWriteError}`);
                  } else {
                    // Verify from the workbook instead of assuming the successful PATCH
                    // has already committed. This catches slow/stale Excel writes.
                    const verifyStart = Date.now();
                    const verifyRes = await fetch(`${wbBase2}/worksheets/${wsId2}/range(address='A1:A500')?$select=values`, { headers:wbHdr2 });
                    if (verifyRes.ok) {
                      const verifyVals = (await verifyRes.json()).values || [];
                      const verifySet = new Set(verifyVals.map(r => String(r[0]||'').trim()).filter(Boolean));
                      const missing = newIds.filter(id => !verifySet.has(id.trim()));
                      context.log(`[CS][timing] verify attempt ${attempt}: ${Date.now()-verifyStart}ms missing=${missing.length}`);
                      if (!missing.length) {
                        verified = true;
                        context.log(`[CS] ✓ Wrote and verified ${newIds.join(',')} in ${csFileName} rows ${nextRow2}-${endRow2}`);
                      } else {
                        lastWriteError = `verification missing ${missing.join(', ')}`;
                        context.log(`[CS] verify attempt ${attempt} incomplete: ${lastWriteError}`);
                      }
                    } else {
                      lastWriteError = `verification read failed (${verifyRes.status})`;
                      context.log(`[CS] verify attempt ${attempt} failed: ${lastWriteError}`);
                    }
                  }
                } catch (writeErr) {
                  lastWriteError = writeErr.message;
                  context.log(`[CS] write attempt ${attempt} error: ${writeErr.message}`);
                }

                if (!verified && attempt < 3) {
                  await new Promise(r => setTimeout(r, 750 * attempt));
                }
              }

              if (!verified) {
                csWarning = `⚠️ Sample approved, but Control Sheet write could not be verified for ${newIds.join(', ')}. Check ${csFileName}.`;
                context.log(`[CS] ${csWarning} Last error: ${lastWriteError || 'unknown'}`);
              }
            } else {
              context.log(`[CS] All IDs already present`);
            }
          } finally {
            const closeStart = Date.now();
            await fetch(`${wbBase2}/closeSession`, { method:'POST', headers:wbHdr2 }).catch(()=>{});
            context.log(`[CS][timing] closeSession: ${Date.now()-closeStart}ms`);
          }
        } else if (!csWarning) {
          csWarning = `Control sheet ${csFileName} not found — lab IDs not written. Create it and add manually.`;
        }
'''

s = s[:start] + new_block + s[end:]

# Add total timing on success/failure without changing approval order.
old_catch = """      } catch(e) { context.log('[CS] Control sheet write failed:', e.message); csWarning = `Control sheet write failed: ${e.message}`; }

      // ── Write to Radon Control Sheet if Radon Water approved ─────────────────
"""
new_catch = """      } catch(e) { context.log('[CS] Control sheet write failed:', e.message); csWarning = `⚠️ Sample approved, but Control Sheet write failed for ${allFullIds.join(', ')}: ${e.message}`; }
      context.log(`[CS][timing] Total control sheet step: ${Date.now()-csOverallStart}ms warning=${csWarning ? 'yes' : 'no'}`);

      // ── Write to Radon Control Sheet if Radon Water approved ─────────────────
"""
if old_catch not in s:
    raise SystemExit('Control Sheet catch marker not found')
s = s.replace(old_catch, new_catch, 1)

p.write_text(s)
