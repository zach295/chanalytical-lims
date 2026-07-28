const { app } = require('@azure/functions');
const { listItems, createItem, updateItem, LISTS } = require('../shared/graph');

// Helper: find user by email without relying on indexed filter
async function findUserByEmail(email) {
  const emailLower = (email || '').toLowerCase().trim();
  const items = await listItems(LISTS.USERS, { top: 200 });
  return items.find(r => {
    const e = (r.Email || r.Title || '').toLowerCase().trim();
    return e === emailLower;
  }) || null;
}

const mapUser = r => ({
  _id:       r._id,
  email:     r.Email     || r.Title || '',
  name:      r.Name      || r.name  || '',
  role:      r.Role      || r.role  || 'lab',
  clientKey: r.ClientKey || '',
  regCode:   r.RegCode   || '',
  createdBy: r.CreatedBy || '',
  createdAt: r.CreatedAt || '',
  mustReset: r.MustReset === true || r.MustReset === 'true' || r.MustReset === 'TRUE',
  active:    r.Active !== false && r.Active !== 'FALSE',
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
          Email:     email,
          Name:      name      || '',
          Role:      role      || 'lab',
          ClientKey: clientKey || '',
          RegCode:   regCode   || '',
          CreatedBy: createdBy || '',
          CreatedAt: new Date().toISOString(),
          MustReset: true,
          Active:    'TRUE',
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
          Name:      name      || '',
          Role:      role      || 'lab',
          ClientKey: clientKey || '',
          RegCode:   regCode   || '',
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
        await updateItem(LISTS.USERS, user._id, { Role: role });
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
        await updateItem(LISTS.USERS, user._id, { Role: 'deactivated', Active: 'FALSE' });
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: true }),
        };
      }

      if (action === 'checklogin') {
        const { email, regCode } = body;
        const user = await findUserByEmail(email);
        if (!user) return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ found: false }),
        };
        const role = user.Role || user.role || 'lab';
        const active = user.Active !== 'FALSE' && user.Active !== false;
        const mustReset = user.MustReset === true || user.MustReset === 'true' || user.MustReset === 'TRUE';

        // If regCode provided, validate it
        if (regCode) {
          const storedCode = (user.RegCode || '').trim();
          if (!storedCode || storedCode !== regCode.trim()) {
            return {
              status: 200,
              headers: { 'content-type': 'application/json' },
              body: JSON.stringify({ found: true, valid: false, error: 'Invalid registration code' }),
            };
          }
        }

        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({
            found:     true,
            valid:     true,
            role,
            name:      user.Name || user.name || email.split('@')[0],
            mustReset,
            active,
          }),
        };
      }

      return { status: 400, body: JSON.stringify({ error: 'Unknown action: ' + action }) };

    } catch(e) {
      context.log('[users-manage] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
