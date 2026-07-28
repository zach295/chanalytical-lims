const { app } = require('@azure/functions');
const { listItems, createItem, updateItem, LISTS } = require('../shared/graph');

async function findUserByEmail(email) {
  const emailLower = (email || '').toLowerCase().trim();
  const items = await listItems(LISTS.USERS, { top: 200 });
  return items.find(r => {
    const e = (r.email || r.Email || r.Title || '').toLowerCase().trim();
    return e === emailLower;
  }) || null;
}

const mapUser = r => ({
  _id:       r._id,
  email:     r.email     || r.Email     || r.Title || '',
  name:      r.name      || r.Name      || '',
  role:      r.role      || r.Role      || 'lab',
  clientKey: r.clientKey || r.ClientKey || '',
  regCode:   r.regCode   || r.RegCode   || '',
  createdBy: r.createdBy || r.CreatedBy || '',
  createdAt: r.createdAt || r.CreatedAt || '',
  mustReset: r.mustReset === true || r.mustReset === 'true' || r.MustReset === true || r.MustReset === 'true',
  active:    r.active !== false && r.active !== 'FALSE' && r.Active !== false && r.Active !== 'FALSE',
});

app.http('users-manage', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (request.method === 'GET') {
        const items = await listItems(LISTS.USERS);
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ users: items.map(mapUser) }),
        };
      }

      const body = await request.json();
      const { action } = body;

      if (action === 'list') {
        const items = await listItems(LISTS.USERS, { top: 200 });
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ users: items.map(mapUser) }),
        };
      }

      if (action === 'create') {
        const { email, name, role, clientKey, createdBy, regCode } = body;
        const existing = await findUserByEmail(email);
        if (existing) return { status: 409, body: JSON.stringify({ error: 'User already exists' }) };
        await createItem(LISTS.USERS, {
          Title:     email,
          email:     email,
          name:      name      || '',
          role:      role      || 'lab',
          clientKey: clientKey || '',
          regCode:   regCode   || '',
          createdBy: createdBy || '',
          createdAt: new Date().toISOString(),
          mustReset: true,
          active:    true,
        });
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: true }),
        };
      }

      if (action === 'edit') {
        const { email, name, role, clientKey, regCode } = body;
        const user = await findUserByEmail(email);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        await updateItem(LISTS.USERS, user._id, {
          name:      name      || '',
          role:      role      || 'lab',
          clientKey: clientKey || '',
          regCode:   regCode   || '',
        });
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: true }),
        };
      }

      if (action === 'setrole') {
        const { email, role } = body;
        const user = await findUserByEmail(email);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        await updateItem(LISTS.USERS, user._id, { role });
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: true }),
        };
      }

      if (action === 'deactivate') {
        const { email } = body;
        const user = await findUserByEmail(email);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        await updateItem(LISTS.USERS, user._id, { role: 'deactivated', active: false });
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: true }),
        };
      }

      if (action === 'checklogin') {
        const { email } = body;
        const user = await findUserByEmail(email);
        if (!user) return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ found: false }),
        };
        const role     = user.role || user.Role || 'lab';
        const active   = user.active !== false && user.active !== 'FALSE' && user.Active !== false && user.Active !== 'FALSE';
        const mustReset = user.mustReset === true || user.mustReset === 'true' || user.MustReset === true;
        const name     = user.name || user.Name || email.split('@')[0];

        if (!active || role === 'deactivated') {
          return {
            status: 200,
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ found: true, valid: false, error: 'Account is deactivated' }),
          };
        }

        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ found: true, valid: true, role, name, mustReset, active }),
        };
      }

      return { status: 400, body: JSON.stringify({ error: 'Unknown action: ' + action }) };

    } catch(e) {
      context.log('[users-manage] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
