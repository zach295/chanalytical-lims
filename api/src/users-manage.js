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

// Try every possible field name variation SharePoint might use
const mapUser = r => {
  // Dump all keys to find what Graph actually returns
  const role = r.role || r.Role || r.role0 || r.Role0 || 'lab';
  const name = r.name || r.Name || r.name0 || r.LinkTitle || '';
  const email = r.email || r.Email || r.Title || '';
  const clientKey = r.clientKey || r.ClientKey || r.clientKey0 || '';
  const regCode = r.regCode || r.RegCode || r.regCode0 || '';
  const createdBy = r.createdBy || r.CreatedBy || r.createdBy0 || '';
  const createdAt = r.createdAt || r.CreatedAt || r.createdAt0 || '';
  const mustReset = r.mustReset === true || r.mustReset === 'true' ||
                    r.MustReset === true || r.MustReset === 'true';
  const active = r.active !== false && r.active !== 'FALSE' &&
                 r.Active !== false && r.Active !== 'FALSE';
  return { _id: r._id, email, name, role, clientKey, regCode, createdBy, createdAt, mustReset, active, _raw: r };
};

app.http('users-manage', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      if (request.method === 'GET') {
        const items = await listItems(LISTS.USERS);
        const mapped = items.map(mapUser);
        // Include raw data so we can see what Graph returns
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ users: mapped }),
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

      if (action === 'debug') {
        // Returns raw Graph fields for one user so we can fix the mapping
        const items = await listItems(LISTS.USERS, { top: 5 });
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ raw: items }),
        };
      }

      if (action === 'create') {
        const { email, name, role, clientKey, createdBy, regCode } = body;
        const existing = await findUserByEmail(email);
        if (existing) return { status: 409, body: JSON.stringify({ error: 'User already exists' }) };
        await createItem(LISTS.USERS, {
          Title: email, email, name: name || '', role: role || 'lab',
          clientKey: clientKey || '', regCode: regCode || '',
          createdBy: createdBy || '', createdAt: new Date().toISOString(),
          mustReset: true, active: true,
        });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'edit') {
        const { email, name, role, clientKey, regCode } = body;
        const user = await findUserByEmail(email);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        // Write with lowercase field names matching SP column names
        await updateItem(LISTS.USERS, user._id, {
          name: name || '', role: role || 'lab',
          clientKey: clientKey || '', regCode: regCode || '',
        });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'setrole') {
        const { email, role } = body;
        const user = await findUserByEmail(email);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        await updateItem(LISTS.USERS, user._id, { role });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'deactivate') {
        const { email } = body;
        const user = await findUserByEmail(email);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        await updateItem(LISTS.USERS, user._id, { role: 'deactivated', active: false });
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'checklogin') {
        const { email } = body;
        const user = await findUserByEmail(email);
        if (!user) return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ found: false }),
        };
        const mapped = mapUser(user);
        if (!mapped.active || mapped.role === 'deactivated') {
          return { status: 200, headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ found: true, valid: false, error: 'Account is deactivated' }) };
        }
        return { status: 200, headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ found: true, valid: true, role: mapped.role,
            name: mapped.name, mustReset: mapped.mustReset, active: mapped.active }) };
      }

      return { status: 400, body: JSON.stringify({ error: 'Unknown action: ' + action }) };

    } catch(e) {
      context.log('[users-manage] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
