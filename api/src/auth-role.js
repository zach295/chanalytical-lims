/**
 * auth-role.js — Azure version
 * Returns a user's role based on their email address.
 * Fetches all users and filters in JS to avoid SharePoint index limitations.
 *
 * POST { email }
 * Returns { success, role, name, email } or { success:false, error }
 */
const { app }       = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

app.http('auth-role', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { email } = await request.json().catch(() => ({}));
      if (!email) return {
        status: 400,
        jsonBody: { success: false, error: 'Email required' },
      };

      const emailLower = email.toLowerCase().trim();

      // Fetch all users and match by email in JS
      const items = await listItems(LISTS.USERS, { top: 200 });
      const user  = items.find(r => {
        const e = (r.Email || r.Title || '').toLowerCase().trim();
        return e === emailLower;
      });

      if (user) {
        const active = user.Active !== 'FALSE' && user.Active !== false;
        const role   = (user.Role || user.role || 'lab').toLowerCase();
        if (!active || role === 'deactivated') {
          return {
            status: 200,
            jsonBody: { success: false, error: 'Account is deactivated. Contact your administrator.' },
          };
        }
        return {
          status: 200,
          jsonBody: {
            success:   true,
            email:     emailLower,
            name:      user.Name || user.name || email.split('@')[0],
            role,
            clientKey: user.ClientKey || null,
          },
        };
      }

      return {
        status: 200,
        jsonBody: {
          success: false,
          error: 'Your account is not authorized. Contact your administrator to be added.',
        },
      };

    } catch(e) {
      context.log('[auth-role] Error:', e.message);
      return { status: 500, jsonBody: { success: false, error: e.message } };
    }
  },
});
