// ── CHANALYTICAL AUTH ─────────────────────────────────────────────────────────
// Uses both localStorage AND cookies so accounts created in the browser
// are automatically available when the app is opened as an installed PWA,
// and vice versa. iOS Safari gives installed PWAs separate localStorage
// but shares cookies with the browser.

const AUTH_KEY = 'cha_users_v2';
const SESSION_KEY = 'cha_session';

const REGISTRATION_CODES = {
  'CHAN2022':   { role: 'admin' },
  'CHANLAB2024':{ role: 'lab' },
  'WARD0823':  { role: 'wq', clientKey: 'ward-water' },
  'MERW0823':  { role: 'wq', clientKey: 'maine-radon' },
  'P2P0823':   { role: 'wq', clientKey: 'downeast' },
  'MHI0226':   { role: 'wq', clientKey: 'madden' },
  'AZW0224':   { role: 'wq', clientKey: 'az-water' },
  'FONT0823':  { role: 'wq', clientKey: 'fontus' },
  'MCI0823':   { role: 'wq', clientKey: 'main-choice' },
  'DHI0226':   { role: 'wq', clientKey: 'defender' },
  'EIS0226':   { role: 'wq', clientKey: 'elliotts' },
  'CPI0224':   { role: 'wq', clientKey: 'campbell' },
  'OPH0224':   { role: 'wq', clientKey: 'onpoint' },
  'PMASON23':  { role: 'wq', clientKey: 'peter-mason' },
  'ADV1123':   { role: 'wq', clientKey: 'advanced' },
  'NOV1224':   { role: 'wq', clientKey: 'nova-enviro' },
  'EVGRN0325': { role: 'wq', clientKey: 'evergreen' },
  'FPI0823':   { role: 'wq', clientKey: 'fpi-chancorp' },
  'SUPIN0526': { role: 'wq', clientKey: 'super-inspector' },
  'AIOHI0526': { role: 'wq', clientKey: 'all-in-one' },
  'YANKEE0823':{ role: 'wq', clientKey: 'yankee' },
};

const Auth = {

  // ── COOKIE BRIDGE (iOS PWA ↔ Safari browser) ──────────────────────────────
  // iOS gives installed PWAs completely separate localStorage from Safari.
  // Cookies ARE shared between Safari and the installed PWA on the same domain,
  // so we write to both. When localStorage is empty (fresh PWA context),
  // we read from the cookie and restore into localStorage.

  _c(n, v, days) { // set cookie
    const e = new Date(Date.now() + (days||365)*864e5).toUTCString();
    // Cap cookie size — only store what's needed (cookies have a 4KB limit)
    const val = typeof v === 'string' ? v : JSON.stringify(v);
    if (val.length < 3800) { // leave headroom
      document.cookie = `${n}=${encodeURIComponent(val)};expires=${e};path=/;SameSite=Lax`;
    }
  },

  _g(n) { // get cookie
    const m = document.cookie.match(new RegExp('(?:^|; )' + n + '=([^;]*)'));
    return m ? decodeURIComponent(m[1]) : null;
  },

  _d(n) { // delete cookie
    document.cookie = `${n}=;expires=Thu, 01 Jan 1970 00:00:00 GMT;path=/`;
  },

  // ── USERS ──────────────────────────────────────────────────────────────────
  getUsers() {
    let raw = localStorage.getItem(AUTH_KEY);
    if (!raw) raw = this._g(AUTH_KEY); // fall back to cookie (PWA context)
    if (!raw) return {};
    try {
      const parsed = JSON.parse(raw);
      // Sync into localStorage if it was missing
      if (!localStorage.getItem(AUTH_KEY)) localStorage.setItem(AUTH_KEY, raw);
      return parsed;
    } catch { return {}; }
  },

  saveUsers(u) {
    const s = JSON.stringify(u);
    localStorage.setItem(AUTH_KEY, s);
    this._c(AUTH_KEY, s); // keep cookie in sync
  },

  hashPassword(pw) {
    let h = 0;
    for (let i = 0; i < pw.length; i++) { h = ((h << 5) - h) + pw.charCodeAt(i); h = h & h; }
    return h.toString(36) + pw.length;
  },

  // ── REGISTER ────────────────────────────────────────────────────────────────
  register(email, password, code, fullName) {
    const users = this.getUsers();
    const key = email.toLowerCase().trim();
    if (users[key]) return { success: false, error: 'An account with this email already exists.' };
    if (password.length < 6) return { success: false, error: 'Password must be at least 6 characters.' };

    let role = 'public', clientKey = null;
    if (code) {
      const d = REGISTRATION_CODES[code.toUpperCase().trim()];
      if (!d) return { success: false, error: 'Invalid registration code.' };
      role = d.role; clientKey = d.clientKey || null;
    }

    const parts = (fullName || '').trim().split(' ');
    const user = {
      email: key, password: this.hashPassword(password), role, clientKey,
      name: fullName || key.split('@')[0],
      firstName: parts[0] || '', lastName: parts.slice(1).join(' ') || '',
      createdAt: new Date().toISOString(),
    };
    users[key] = user;
    this.saveUsers(users);
    this.createSession(user);

    // Sync to Google Sheets so admin can see all accounts from any device
    // NO password sent to server — Sheets never stores credentials
    fetch('/api/users-manage', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'create',
        email: key, name: user.name,
        role, clientKey: clientKey || '',
        regCode: code || '',
        createdBy: 'Self-registered (' + (code || 'no code') + ')',
      }),
    }).catch(() => {}); // silently ignore if server unavailable

    return { success: true, user };
  },

  // ── LOGIN ───────────────────────────────────────────────────────────────────
  async loginAsync(email, password) {
    // Always check server first for deactivation status
    // (deactivated accounts must be blocked even if local copy exists)
    try {
      const res = await fetch('/api/users-manage', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'login', email, password }),
      });
      const data = await res.json();

      // Server explicitly blocked this account (deactivated) — override everything
      if (!data.success && data.error?.includes('deactivated')) {
        return { success: false, error: data.error };
      }

      // Server says mustReset — show password reset flow
      if (data.success && data.mustReset) {
        const user = { ...data.user, fromServer: true };
        this.createSession(user);
        return { success: true, user, mustReset: true };
      }

      // Server login succeeded for a server-only account
      if (data.success && !data.mustReset) {
        const user = { ...data.user, fromServer: true };
        this.createSession(user);
        return { success: true, user };
      }
    } catch {
      // Server unavailable — fall through to local check
    }

    // Fall back to local account (self-registered with code)
    const localResult = this.login(email, password);
    return localResult;
  },

  login(email, password) {
    const users = this.getUsers();
    const key = email.toLowerCase().trim();
    const user = users[key];
    if (!user) return { success: false, error: 'No account found with this email.' };
    if (user.password !== this.hashPassword(password)) return { success: false, error: 'Incorrect password.' };
    this.createSession(user);
    return { success: true, user };
  },

  // ── SESSION ─────────────────────────────────────────────────────────────────
  createSession(user) {
    const isAdmin = user.role === 'admin';
    const session = JSON.stringify({
      email: user.email, role: user.role, clientKey: user.clientKey,
      name: user.name, loginAt: new Date().toISOString(), isAdmin,
    });
    localStorage.setItem(SESSION_KEY, session);
    // 30-day session cookie — shared between browser and PWA
    this._c(SESSION_KEY, session, 30);
    if (isAdmin) sessionStorage.setItem('cha_admin_alive', '1');
  },

  getSession() {
    let raw = localStorage.getItem(SESSION_KEY);
    if (!raw) raw = this._g(SESSION_KEY); // fall back to cookie
    if (!raw) return null;

    try {
      const session = JSON.parse(raw);
      // Restore into localStorage if it came from cookie
      if (!localStorage.getItem(SESSION_KEY)) localStorage.setItem(SESSION_KEY, raw);

      if (session.isAdmin && !sessionStorage.getItem('cha_admin_alive')) {
        localStorage.removeItem(SESSION_KEY);
        this._d(SESSION_KEY);
        return null;
      }
      return session;
    } catch { return null; }
  },

  logout() {
    localStorage.removeItem(SESSION_KEY);
    localStorage.removeItem(AUTH_KEY); // force re-read from cookie on next load
    sessionStorage.removeItem('cha_admin_alive');
    this._d(SESSION_KEY);
    window.location.href = '/login.html';
  },

  requireAuth(roles) {
    const s = this.getSession();
    if (!s || (roles && !roles.includes(s.role))) {
      window.location.href = '/login.html';
      return null;
    }
    return s;
  },
};
