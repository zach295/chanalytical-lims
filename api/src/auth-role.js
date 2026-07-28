/**
 * auth-role.js — Azure version
 * Returns a user's role based on their email address.
 * Uses SharePoint internal field names (field_1=name, field_2=role, field_9=active).
 *
 * POST { email }
 * Returns { success, role, name, email } or { success:false, error }
 */
const { app } = require('@azure/functions');
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

      const items = await listItems(LISTS.USERS, { top: 200 });
      const user  = items.find(r => {
        const e = (r.Title || '').toLowerCase().trim();
        return e === emailLower;
      });

      if (user) {
        const active = user.field_9 !== false && user.field_9 !== 'FALSE';
        const role   = (user.field_2 || 'lab').toLowerCase();
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
            name:      user.field_1 || email.split('@')[0],
            role,
            clientKey: user.field_3 || null,
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
