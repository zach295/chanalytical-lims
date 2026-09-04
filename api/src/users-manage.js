const { app } = require('@azure/functions');

// Same hash as client-side Auth.hashPassword — must stay in sync
function hashPassword(pw) {
  let h = 0;
  for (let i = 0; i < pw.length; i++) { h = ((h << 5) - h) + pw.charCodeAt(i); h = h & h; }
  return h.toString(36) + pw.length;
}
const WELCOME_HASH = '-tlew818'; // hash of W3lcom3!
const { listItems, createItem, updateItem, LISTS } = require('../shared/graph');
const { writeActivityLog } = require('../shared/audit');

// SharePoint internal field name mapping (discovered via debug)
// field_1=name, field_2=role, field_4=regCode, field_5=createdBy,
// field_6=createdAt, field_7=mustReset, field_9=active
// Title=email, field_3=clientKey (assumed)

async function findUserByEmail(email) {
  const emailLower = (email || '').toLowerCase().trim();
  const items = await listItems(LISTS.USERS, { top: 200 });
  return items.find(r => {
    const e = (r.Title || '').toLowerCase().trim();
    return e === emailLower;
  }) || null;
}

const mapUser = r => ({
  _id:       r._id,
  email:     r.Title      || '',
  name:      r.field_1    || '',
  role:      r.field_2    || 'lab',
  clientKey: r.field_3    || '',
  regCode:   r.field_4    || '',
  createdBy: r.field_5    || '',
  createdAt: r.field_6    || '',
  mustReset: r.field_7 === true || r.field_7 === 'true',
  active:    r.field_9 !== false && r.field_9 !== 'FALSE',
});

async function logUserActivity(action, email, details, by, context) {
  return writeActivityLog({
    labId: email,
    type: action,
    notes: details || '',
    by: by || 'Admin',
    context,
  });
}

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
          Title:   email,
          field_1: name      || '',
          field_2: role      || 'lab',
          field_3: clientKey || '',
          field_4: regCode   || '',
          field_5: createdBy || '',
          field_6: new Date().toISOString(),
          field_7: true,
          field_9: true,
        });
        await logUserActivity('User Created', email, `New account created in Users list.\nRole: ${role || 'lab'} | Created by: ${createdBy || 'Admin'}`, createdBy, context);
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'edit') {
        const { email, name, role, clientKey, regCode, newEmail } = body;
        const user = await findUserByEmail(email);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        const updates = {
          field_1: name      || '',
          field_2: role      || 'lab',
          field_3: clientKey || '',
          field_4: regCode   || '',
        };
        // Update email (Title) if a new one was provided and it differs
        if (newEmail && newEmail !== email) updates.Title = newEmail.trim().toLowerCase();
        await updateItem(LISTS.USERS, user._id, updates);
        await logUserActivity('User Edited', email, `User account updated in Users list.\nNew email: ${newEmail || email} | Name: ${name || '—'} | Role: ${role || '—'}`, body.updatedBy || 'Admin', context);
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'setrole') {
        const { email, role } = body;
        const user = await findUserByEmail(email);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        await updateItem(LISTS.USERS, user._id, { field_2: role });
        const { email: srEmail, role: srRole } = body;
        await logUserActivity('Role Changed', srEmail, `Role: "${user.field_2 || ''}" → "${srRole || '—'}"`, body.updatedBy || 'Admin', context);
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true }) };
      }

      if (action === 'deactivate') {
        const { email } = body;
        const user = await findUserByEmail(email);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        await updateItem(LISTS.USERS, user._id, { field_2: 'deactivated', field_9: false });
        const audit = await logUserActivity('User Deactivated', email, `Role: "${user.field_2 || ''}" → "deactivated" | Active: true → false`, body.updatedBy || 'Admin', context);
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true, auditWarning: audit.success ? null : audit.error }) };
      }

      if (action === 'login') {
        // Full login — validates email + password server-side
        const { email: loginEmail, password } = body;
        const user = await findUserByEmail(loginEmail);
        if (!user) return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ success: false, error: 'No account found with this email.' }),
        };
        const mapped = mapUser(user);
        if (!mapped.active || mapped.role === 'deactivated') {
          return { status: 200, headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ success: false, error: 'Account is deactivated. Contact your administrator.' }) };
        }
        // Get stored password hash from field_8 (resetDate col) or use welcome hash
        const storedHash = user.field_8 || WELCOME_HASH;
        const inputHash  = hashPassword(password || '');
        if (inputHash !== storedHash && password !== 'W3lcom3!') {
          // Also allow plain W3lcom3! as fallback for new accounts
          if (hashPassword('W3lcom3!') !== inputHash) {
            return { status: 200, headers: { 'content-type': 'application/json' },
              body: JSON.stringify({ success: false, error: 'Incorrect password.' }) };
          }
        }
        return {
          status: 200,
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({
            success:   true,
            mustReset: mapped.mustReset,
            user: {
              email:     mapped.email,
              name:      mapped.name,
              role:      mapped.role,
              clientKey: mapped.clientKey,
            },
          }),
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
        const mapped = mapUser(user);
        if (!mapped.active || mapped.role === 'deactivated') {
          return { status: 200, headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ found: true, valid: false, error: 'Account is deactivated' }) };
        }
        return { status: 200, headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ found: true, valid: true, role: mapped.role,
            name: mapped.name, mustReset: mapped.mustReset, active: mapped.active }) };
      }

      if (action === 'setpw') {
        // Set a new password for a user (admin action or first-time setup)
        const pwEmail = body.email;
        const newPw = body.password || body.newPassword || '';
        if (newPw.length < 6) return { status: 400, body: JSON.stringify({ error: 'Password must be at least 6 characters.' }) };
        const user = await findUserByEmail(pwEmail);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        const hashed = hashPassword(newPw || '');
        await updateItem(LISTS.USERS, user._id, { field_8: hashed, field_7: false, field_9: true });
        const mapped = mapUser(user);
        const audit = await logUserActivity('Password Changed', pwEmail, 'Password changed. Password value/hash intentionally not recorded.', body.updatedBy || mapped.name || 'User', context);
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({
          success: true,
          auditWarning: audit.success ? null : audit.error,
          user: { email: mapped.email, name: mapped.name, role: mapped.role, clientKey: mapped.clientKey },
        })};
      }

      if (action === 'resetpw') {
        // Reset password to welcome password — user must change on next login
        const { email: pwEmail } = body;
        const user = await findUserByEmail(pwEmail);
        if (!user) return { status: 404, body: JSON.stringify({ error: 'User not found' }) };
        await updateItem(LISTS.USERS, user._id, { field_8: WELCOME_HASH, field_7: true, field_9: true });
        const audit = await logUserActivity('Password Reset', pwEmail, 'Password reset to temporary welcome credential; user must change it on next login. Password value/hash intentionally not recorded.', body.updatedBy || 'Admin', context);
        return { status: 200, headers: { 'content-type': 'application/json' }, body: JSON.stringify({ success: true, auditWarning: audit.success ? null : audit.error }) };
      }

      return { status: 400, body: JSON.stringify({ error: 'Unknown action: ' + action }) };

    } catch(e) {
      context.log('[users-manage] Error:', e.message);
      return { status: 500, body: JSON.stringify({ error: e.message }) };
    }
  }
});
