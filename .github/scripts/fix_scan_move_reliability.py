from pathlib import Path
p=Path('api/src/scan-folder.js')
s=p.read_text()
old='''async function moveSpFile(itemId, destFolderPath, token) {
  const siteId    = process.env.SP_SITE_ID;
  const drivePath = toDrivePath(destFolderPath);

  // Resolve destination folder ID
  const folderRes = await fetch(
    `${GRAPH}/sites/${siteId}/drive/root:/${drivePath}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!folderRes.ok) {
    console.warn(`[moveSpFile] Cannot resolve destination ${destFolderPath}: ${folderRes.status}`);
    return;
  }
  const folderData = await folderRes.json();
  const destId     = folderData.id;

  const patchRes = await fetch(
    `${GRAPH}/sites/${siteId}/drive/items/${itemId}`,
    {
      method:  'PATCH',
      headers: { Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
      body:    JSON.stringify({ parentReference: { id: destId } }),
    }
  );
  if (!patchRes.ok) {
    const err = await patchRes.text();
    console.warn(`[moveSpFile] Move failed for ${itemId}: ${patchRes.status} ${err}`);
  }
}
'''
new='''async function moveSpFile(itemId, destFolderPath, token) {
  const siteId    = process.env.SP_SITE_ID;
  const drivePath = toDrivePath(destFolderPath);
  let lastError = null;

  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      // Resolve destination folder ID on each attempt in case Graph returned stale data.
      const folderRes = await fetch(
        `${GRAPH}/sites/${siteId}/drive/root:/${drivePath}?$select=id,name,parentReference`,
        { headers: { Authorization: `Bearer ${token}` }, cache: 'no-store' }
      );
      if (!folderRes.ok) {
        throw new Error(`cannot resolve destination ${destFolderPath}: ${folderRes.status} ${await folderRes.text()}`);
      }
      const folderData = await folderRes.json();
      const destId = folderData.id;
      if (!destId) throw new Error(`destination ${destFolderPath} returned no folder id`);

      const patchRes = await fetch(
        `${GRAPH}/sites/${siteId}/drive/items/${itemId}`,
        {
          method:  'PATCH',
          headers: { Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
          body:    JSON.stringify({ parentReference: { id: destId } }),
        }
      );
      if (!patchRes.ok) {
        throw new Error(`move failed: ${patchRes.status} ${await patchRes.text()}`);
      }

      // Verify the item really landed in Review before OCR starts. Previously move
      // failures were only logged and processing continued while the file stayed in Incoming.
      for (let verify = 0; verify < 5; verify++) {
        const metaRes = await fetch(
          `${GRAPH}/sites/${siteId}/drive/items/${itemId}?$select=id,name,parentReference`,
          { headers: { Authorization: `Bearer ${token}` }, cache: 'no-store' }
        );
        if (metaRes.ok) {
          const meta = await metaRes.json();
          if (meta.parentReference?.id === destId) return meta;
        }
        await new Promise(r => setTimeout(r, 400));
      }
      throw new Error('move verification failed; item is not in destination folder yet');
    } catch (e) {
      lastError = e;
      console.warn(`[moveSpFile] Attempt ${attempt}/3 failed for ${itemId}: ${e.message}`);
      if (attempt < 3) await new Promise(r => setTimeout(r, 750 * attempt));
    }
  }

  throw new Error(`Could not move scan into Review after 3 attempts: ${lastError?.message || 'unknown error'}`);
}
'''
if old not in s:
    raise SystemExit('moveSpFile anchor not found')
s=s.replace(old,new,1)
p.write_text(s)
print('Hardened scan move to Review with retry + verification')
