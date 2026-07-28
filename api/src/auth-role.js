/**
 * auth-role.js — Azure version
 * Returns a user's role based on their email address.
 * Called after Microsoft OAuth2 completes — email is verified by Microsoft,
 * we just need to return the role from the SharePoint Users list.
 *
 * POST { email }
 * Returns { success, role, name, email } or { success:false, error }
 */
const { app }    = require('@azure/functions');
const { findItem, LISTS } = require('../shared/graph');

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

      // Look up user in SharePoint Users list by Title (which stores email)
      const user = await findItem(LISTS.USERS, 'Title', emailLower).catch(() => null);

      if (user) {
        const active = user.Active !== 'FALSE' && user.Active !== false;
        const role   = (user.Role || 'lab').toLowerCase();
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
            name:      user.Name || email.split('@')[0],
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
