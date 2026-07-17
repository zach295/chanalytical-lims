const { app } = require('@azure/functions');
const { listItems, createItem, updateItem, findItem, LISTS } = require('../shared/graph');

app.http('users-manage', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (request.method === 'GET') {
        const items = await listItems(LISTS.USERS);
        const users = items.map(r => ({
          email:     r.Email     || r.Title || '',
          name:      r.Name      || '',
          role:      r.Role      || 'lab',
          clientKey: r.ClientKey || '',
          regCode:   r.RegCode   || '',
          createdBy: r.CreatedBy || '',
          createdAt: r.CreatedAt || '',
          mustReset: r.MustReset === true || r.MustReset === 'true',
          active:    r.Active !== false && r.Active !== 'FALSE',
          _id:       r._id,
        }));
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ users }) };
      }

      const body = await request.json();
      const { action } = body;

      if (action === 'create') {
        const { email, name, role, clientKey, createdBy } = body;
        const existing = await findItem(LISTS.USERS, 'Email', email);
        if (existing) return { status: 409, body: JSON.stringify({ error: 'User already exists' }) };
        await createItem(LISTS.USERS, {
          Title: email, Email: email, Name: name, Role: role || 'lab',
          ClientKey: clientKey || '', CreatedBy: createdBy || '',
          CreatedAt: new Date().toISOString(), MustReset: true, Active: 'TRUE',
        });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'setrole') {
        const { email, role } = body;
        const user = await findItem(LISTS.USERS, 'Email', email);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        await updateItem(LISTS.USERS, user._id, { Role: role });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'deactivate') {
        const { email } = body;
        const user = await findItem(LISTS.USERS, 'Email', email);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        await updateItem(LISTS.USERS, user._id, { Role: 'deactivated', Active: 'FALSE' });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'checklogin') {
        const { email } = body;
        const user = await findItem(LISTS.USERS, 'Email', email);
        if (!user) return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ found: false }) };
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ found: true, role: user.Role, mustReset: user.MustReset, active: user.Active !== 'FALSE' }),
        };
      }

      return { status: 400, body: JSON.stringify({ error: 'Unknown action: ' + action }) };
    } catch(e) {
      context.log('[users-manage] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
