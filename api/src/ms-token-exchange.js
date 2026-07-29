/**
 * ms-token-exchange.js — Azure version
 * Server-side Microsoft OAuth2 PKCE token exchange.
 * Receives the auth code + verifier from the browser, exchanges for tokens,
 * extracts the email, and returns the user's role from the Users list.
 *
 * POST { code, verifier, redirectUri }
 * Returns { success, email, name, role, clientKey } or { success:false, error }
 */
const { app }       = require('@azure/functions');
const { listItems, LISTS } = require('../shared/graph');

const MS_TENANT_ID = process.env.MS_TENANT_ID || 'organizations';
const MS_CLIENT_ID = process.env.MS_CLIENT_ID  || 'c31d824c-e8d2-4557-8198-bfcaba46b338';
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET || '';

app.http('ms-token-exchange', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { code, verifier, redirectUri } = await request.json().catch(() => ({}));

      if (!code)      return { status: 400, jsonBody: { success: false, error: 'code required' } };
      if (!verifier)  return { status: 400, jsonBody: { success: false, error: 'verifier required' } };
      if (!redirectUri) return { status: 400, jsonBody: { success: false, error: 'redirectUri required' } };

      // Exchange code for tokens server-side
      const tokenRes = await fetch(
        `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: new URLSearchParams({
            client_id:     MS_CLIENT_ID,
            client_secret: MS_CLIENT_SECRET,
            code,
            redirect_uri:  redirectUri,
            grant_type:    'authorization_code',
            code_verifier: verifier,
            scope:         'openid profile email User.Read',
          }),
        }
      );

      if (!tokenRes.ok) {
        const errBody = await tokenRes.text();
        context.log('[ms-token-exchange] Token exchange failed:', tokenRes.status, errBody);
        return { status: 400, jsonBody: {
          success: false,
          error: `Microsoft token exchange failed (${tokenRes.status}). Check Azure redirect URI and client secret.`,
          detail: errBody.slice(0, 300),
        }};
      }

      const tokens = await tokenRes.json();

      if (!tokens.id_token) {
        return { status: 400, jsonBody: {
          success: false,
          error: 'Microsoft did not return an ID token. Enable ID tokens in Azure App Registration → Authentication → Implicit grant.',
        }};
      }

      // Decode id_token
      let payload;
      try {
        const b64 = tokens.id_token.split('.')[1].replace(/-/g,'+').replace(/_/g,'/');
        payload = JSON.parse(atob(b64));
      } catch(e) {
        return { status: 400, jsonBody: { success: false, error: 'Failed to decode Microsoft ID token: ' + e.message } };
      }

      const email = (payload.email || payload.preferred_username || payload.upn || '').toLowerCase().trim();
      const name  = payload.name || email.split('@')[0];

      if (!email) {
        return { status: 400, jsonBody: {
          success: false,
          error: 'Microsoft did not provide an email address. Ensure the "email" scope is granted.',
        }};
      }

      context.log('[ms-token-exchange] Authenticated:', email);

      // Look up user role from SharePoint Users list
      const items = await listItems(LISTS.USERS, { top: 200 });
      const user  = items.find(r => (r.Title || '').toLowerCase().trim() === email);

      if (!user) {
        return { status: 200, jsonBody: {
          success: false,
          error: `Your account (${email}) is not authorized. Contact your administrator to be added.`,
        }};
      }

      const active = user.field_9 !== false;
      const role   = (user.field_2 || 'lab').toLowerCase();

      if (!active || role === 'deactivated') {
        return { status: 200, jsonBody: {
          success: false,
          error: 'Your account is deactivated. Contact your administrator.',
        }};
      }

      return {
        status: 200,
        jsonBody: {
          success:   true,
          email,
          name:      user.field_1 || name,
          role,
          clientKey: user.field_3 || null,
        },
      };

    } catch(e) {
      context.log('[ms-token-exchange] Error:', e.message);
      return { status: 500, jsonBody: { success: false, error: e.message } };
    }
  },
});
