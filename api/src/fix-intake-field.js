const { app } = require('@azure/functions');

async function getToken() {
  const res = await fetch(
    `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
    { method:'POST', headers:{'Content-Type':'application/x-www-form-urlencoded'},
      body: new URLSearchParams({
        grant_type:'client_credentials',
        client_id: process.env.AZURE_CLIENT_ID,
        client_secret: process.env.AZURE_CLIENT_SECRET,
        scope:'https://graph.microsoft.com/.default'
      })
    }
  );
  return (await res.json()).access_token;
}

app.http('fix-intake-field', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { itemId, field, value } = await request.json();
      if (!itemId || !field || value === undefined)
        return { status:400, jsonBody:{ error:'itemId, field, value required' }};
      const siteId = process.env.SP_SITE_ID;
      const token  = await getToken();
      // Find the Archived Intake list
      const listsRes = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$select=id,displayName`,
        { headers:{ Authorization:`Bearer ${token}` }}
      );
      const lists = await listsRes.json();
      const list  = (lists.value||[]).find(l => l.displayName === 'Archived Intake');
      if (!list) return { status:404, jsonBody:{ error:'Archived Intake list not found' }};
      const patchRes = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`,
        { method:'PATCH', headers:{ Authorization:`Bearer ${token}`, 'Content-Type':'application/json' },
          body: JSON.stringify({ [field]: value }) }
      );
      if (!patchRes.ok) {
        const err = await patchRes.text();
        return { status:500, jsonBody:{ error:err }};
      }
      context.log(`[fix] Patched item ${itemId}: ${field} = "${value}"`);
      return { status:200, jsonBody:{ success:true, itemId, field, value }};
    } catch(e) {
      return { status:500, jsonBody:{ error: e.message }};
    }
  }
});
