from pathlib import Path
p = Path('admin-dashboard.html')
s = p.read_text()

# 1) Add secure Admin nav entry.
nav_anchor = '''    <button class="sb-btn" id="sb-statistics-btn" onclick="showSection('statistics',this)" style="display:none;">📈 Statistics</button>\n\n\n  </div>'''
nav_new = '''    <button class="sb-btn" id="sb-statistics-btn" onclick="showSection('statistics',this)" style="display:none;">📈 Statistics</button>\n\n    <div class="sb-label" id="sb-secure-admin-label" style="display:none;">Administration</div>\n    <button class="sb-btn" id="sb-secure-admin-btn" onclick="openSecureAdmin(this)" style="display:none;">🔒 Admin</button>\n\n\n  </div>'''
if nav_anchor not in s: raise SystemExit('nav anchor not found')
s = s.replace(nav_anchor, nav_new, 1)

# 2) Add Admin section before Results Entry.
section_anchor = '''  <!-- ══ RESULTS ENTRY ════════════════════════════════════════════════════════ -->'''
admin_section = r'''  <!-- ══ SECURE ADMIN ═══════════════════════════════════════════════════════ -->
  <div class="section" id="section-admin">
    <div class="section-hdr">
      <div>
        <div class="section-hdr-title">🔒 Administration</div>
        <div style="font-size:12px;color:var(--slate);margin-top:3px;">Password-protected tools. Admin verification expires after 15 minutes.</div>
      </div>
      <button class="action-btn" onclick="lockSecureAdmin()">🔐 Lock Admin</button>
    </div>

    <div class="card" style="margin-bottom:18px;">
      <div class="card-header"><div class="card-title">🗑️ Sample Cleanup</div></div>
      <div style="padding:20px;">
        <div style="font-size:13px;color:var(--slate);margin-bottom:14px;line-height:1.5;">
          Finds the complete accession group using the exact <strong>CreatedAt</strong> value in the Accession Log. Multi-element billing rows are included. Radon is treated as its separate Lab ID and is cleared from both the regular Control Sheet and Radon Control Sheet when linked to the same accession.
        </div>
        <div style="display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap;">
          <div style="flex:1;min-width:240px;">
            <label style="display:block;font-size:12px;font-weight:600;color:var(--slate);margin-bottom:5px;">Lab ID</label>
            <input id="admin-cleanup-labid" type="text" placeholder="090426-014" style="width:100%;padding:10px 12px;border:1.5px solid var(--border);border-radius:8px;font-size:15px;font-family:monospace;" />
          </div>
          <button class="btn-ok" style="flex:none;padding:10px 18px;" onclick="adminFindSampleGroup()">Find Sample Group</button>
        </div>
        <div id="admin-cleanup-status" style="font-size:12px;color:var(--slate);margin-top:10px;min-height:18px;"></div>
        <div id="admin-cleanup-preview" style="display:none;margin-top:16px;"></div>
      </div>
    </div>

    <div class="card">
      <div class="card-header">
        <div>
          <div class="card-title">↩️ Deleted Samples / Recovery</div>
          <div style="font-size:11px;color:var(--slate);margin-top:3px;">Recovery snapshots are permanently purged after 30 days. Activity Log entries remain.</div>
        </div>
        <button class="action-btn" onclick="adminLoadRecoveries()">↻ Refresh</button>
      </div>
      <div id="admin-recovery-wrap" style="overflow-x:auto;">
        <table>
          <thead><tr><th>Requested Lab ID</th><th>Related Lab IDs</th><th>Deleted</th><th>Expires</th><th>Deleted By</th><th>Reason</th><th>Action</th></tr></thead>
          <tbody id="admin-recovery-body"><tr><td colspan="7"><div class="empty-state">Unlock Admin to view recoveries.</div></td></tr></tbody>
        </table>
      </div>
    </div>
  </div>

'''
if section_anchor not in s: raise SystemExit('section anchor not found')
s = s.replace(section_anchor, admin_section + section_anchor, 1)

# 3) Add password verification modal before toast.
modal_anchor = '''<div class="toast" id="toast"></div>'''
modal = r'''<!-- Secure Admin password verification -->
<div class="modal-overlay" id="modal-admin-auth">
  <div class="modal" style="max-width:430px;">
    <div class="modal-title">🔒 Admin Verification</div>
    <div style="font-size:13px;color:var(--slate);line-height:1.5;margin-bottom:14px;">
      Re-enter your <strong>LIMS password</strong> to open the Administration tools. This does not store your password.
    </div>
    <div class="form-field">
      <label>Signed in as</label>
      <input id="admin-auth-email" type="text" readonly />
    </div>
    <div class="form-field">
      <label>LIMS Password</label>
      <input id="admin-auth-password" type="password" data-allow-autofill="true" autocomplete="current-password" onkeydown="if(event.key==='Enter')submitAdminVerification()" />
    </div>
    <div id="admin-auth-error" class="msg-inline"></div>
    <div class="modal-actions">
      <button class="btn-cancel" onclick="closeModal('admin-auth')">Cancel</button>
      <button class="btn-ok" id="admin-auth-submit" onclick="submitAdminVerification()">Unlock Admin</button>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>'''
if modal_anchor not in s: raise SystemExit('modal anchor not found')
s = s.replace(modal_anchor, modal, 1)

# 4) Add sidebar permission IDs for admin only.
perm_anchor = '''             'sb-admin-label','sb-clients-btn','sb-users-btn',\n             'sb-testtypes-btn','sb-statistics-btn','sb-intake-btn'],'''
perm_new = '''             'sb-admin-label','sb-clients-btn','sb-users-btn',\n             'sb-testtypes-btn','sb-statistics-btn','sb-intake-btn',\n             'sb-secure-admin-label','sb-secure-admin-btn'],'''
if perm_anchor not in s: raise SystemExit('permission anchor not found')
s = s.replace(perm_anchor, perm_new, 1)

# 5) Add admin section role guard.
role_anchor = '''      testtypes:         ['lab','admin'],\n    };'''
role_new = '''      testtypes:         ['lab','admin'],\n      admin:             ['admin'],\n    };'''
if role_anchor not in s: raise SystemExit('section role anchor not found')
s = s.replace(role_anchor, role_new, 1)

# 6) Add admin load hook.
hook_anchor = '''    if (id === 'testtypes') loadTestTypesTab();\n    if (id === 'coc') { loadCOCSettings().then(() => renderPreview()); }'''
hook_new = '''    if (id === 'testtypes') loadTestTypesTab();\n    if (id === 'admin') adminLoadRecoveries();\n    if (id === 'coc') { loadCOCSettings().then(() => renderPreview()); }'''
if hook_anchor not in s: raise SystemExit('hook anchor not found')
s = s.replace(hook_anchor, hook_new, 1)

# 7) Insert secure admin functions before showSection.
js_anchor = '''  function showSection(id,btn) {'''
js = r'''  // ── SECURE ADMIN / SAMPLE CLEANUP ─────────────────────────────────────
  let _adminPendingButton = null;
  let _adminPreviewGroup = null;
  const ADMIN_TOKEN_KEY = 'cha_admin_reauth_token';
  const ADMIN_EXP_KEY   = 'cha_admin_reauth_exp';

  function adminRoles() {
    return String(_session?.role || '').split(',').map(v=>v.trim());
  }
  function adminIsRole() { return adminRoles().includes('admin'); }
  function adminToken() {
    const t = sessionStorage.getItem(ADMIN_TOKEN_KEY) || '';
    const exp = parseInt(sessionStorage.getItem(ADMIN_EXP_KEY) || '0', 10);
    if (!t || !exp || Date.now() >= exp) {
      sessionStorage.removeItem(ADMIN_TOKEN_KEY);
      sessionStorage.removeItem(ADMIN_EXP_KEY);
      return '';
    }
    return t;
  }
  function lockSecureAdmin() {
    sessionStorage.removeItem(ADMIN_TOKEN_KEY);
    sessionStorage.removeItem(ADMIN_EXP_KEY);
    _adminPreviewGroup = null;
    showSection('overview', document.querySelector('.sb-btn[onclick*="overview"]'));
    showToast('🔐 Admin tools locked');
  }
  function openSecureAdmin(btn) {
    if (!adminIsRole()) { showToast('⚠️ Admin access required.'); return; }
    _adminPendingButton = btn || document.getElementById('sb-secure-admin-btn');
    if (adminToken()) {
      showSection('admin', _adminPendingButton);
      return;
    }
    const s = Auth.getSession();
    const email = document.getElementById('admin-auth-email');
    const pw = document.getElementById('admin-auth-password');
    const err = document.getElementById('admin-auth-error');
    if (email) email.value = s?.email || '';
    if (pw) pw.value = '';
    if (err) { err.className='msg-inline'; err.textContent=''; }
    openModal('admin-auth');
    setTimeout(()=>pw?.focus(), 50);
  }
  window.openSecureAdmin = openSecureAdmin;

  async function submitAdminVerification() {
    const s = Auth.getSession();
    const pw = document.getElementById('admin-auth-password')?.value || '';
    const err = document.getElementById('admin-auth-error');
    const btn = document.getElementById('admin-auth-submit');
    if (!pw) { if(err){err.className='msg-inline err';err.textContent='Enter your LIMS password.';} return; }
    if (btn) { btn.disabled=true; btn.textContent='Verifying...'; }
    try {
      const res = await fetch('/api/sample-admin', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({ action:'verify', email:s?.email || '', password:pw })
      });
      const data = await res.json();
      if (!res.ok || !data.success) throw new Error(data.error || 'Admin verification failed.');
      sessionStorage.setItem(ADMIN_TOKEN_KEY, data.token);
      sessionStorage.setItem(ADMIN_EXP_KEY, String(Date.now() + (data.expiresInMinutes || 15)*60000));
      document.getElementById('admin-auth-password').value='';
      closeModal('admin-auth');
      showSection('admin', _adminPendingButton || document.getElementById('sb-secure-admin-btn'));
      showToast('🔓 Admin unlocked for 15 minutes');
    } catch(e) {
      if (err) { err.className='msg-inline err'; err.textContent=e.message; }
    } finally {
      if (btn) { btn.disabled=false; btn.textContent='Unlock Admin'; }
    }
  }
  window.submitAdminVerification = submitAdminVerification;

  async function adminApi(action, extra={}) {
    const token = adminToken();
    if (!token) {
      openSecureAdmin(document.getElementById('sb-secure-admin-btn'));
      throw new Error('Admin verification required.');
    }
    const res = await fetch('/api/sample-admin', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({ action, adminToken:token, ...extra })
    });
    const data = await res.json().catch(()=>({error:'Server returned an invalid response.'}));
    if (res.status === 401) {
      sessionStorage.removeItem(ADMIN_TOKEN_KEY); sessionStorage.removeItem(ADMIN_EXP_KEY);
      openSecureAdmin(document.getElementById('sb-secure-admin-btn'));
    }
    if (!res.ok || data.error) throw new Error(data.error || `Admin request failed (${res.status})`);
    return data;
  }

  function adminEsc(v) {
    return String(v ?? '').replace(/[&<>"']/g, ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
  }
  function adminFmtDT(v) {
    if (!v) return '—';
    try { return new Date(v).toLocaleString('en-US',{timeZone:'America/New_York'}); } catch { return v; }
  }

  async function adminFindSampleGroup() {
    const input = document.getElementById('admin-cleanup-labid');
    const status = document.getElementById('admin-cleanup-status');
    const preview = document.getElementById('admin-cleanup-preview');
    const baseId = String(input?.value || '').trim().split(' ')[0];
    if (!baseId) { showToast('Enter a Lab ID.'); return; }
    status.textContent = 'Searching Accession Log and linked records...';
    preview.style.display='none';
    _adminPreviewGroup = null;
    try {
      const data = await adminApi('preview',{baseId});
      const g = data.group;
      _adminPreviewGroup = g;
      const listRows = Object.entries(g.counts || {}).map(([k,v])=>`<tr><td>${adminEsc(k)}</td><td style="font-weight:700;">${v}</td></tr>`).join('');
      const sheetRows = (g.controlSheets || []).map(x=>`<tr><td>${x.kind==='radon'?'Radon Control Sheet':'Control Sheet'}</td><td><code>${adminEsc(x.baseId)}</code></td><td>${adminEsc(x.file)}</td><td>${x.row}</td></tr>`).join('');
      const tests = (g.tests || []).map(t=>`<div style="padding:5px 0;border-bottom:1px solid var(--border);"><code>${adminEsc(t.baseId)}</code> — ${adminEsc(t.test || t.suffix || '—')}</div>`).join('');
      preview.innerHTML = `
        <div style="border:1.5px solid var(--border);border-radius:10px;padding:16px;background:var(--off-white);">
          <div style="display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-bottom:14px;">
            <div><div style="font-size:11px;color:var(--slate);font-weight:700;text-transform:uppercase;">CreatedAt Group</div><div style="font-family:monospace;font-weight:700;">${adminEsc(g.createdAt)}</div></div>
            <div><div style="font-size:11px;color:var(--slate);font-weight:700;text-transform:uppercase;">Related Lab IDs</div><div>${(g.baseIds||[]).map(x=>`<code>${adminEsc(x)}</code>`).join(' ')}</div></div>
            <div><div style="font-size:11px;color:var(--slate);font-weight:700;text-transform:uppercase;">Radon Lab IDs</div><div>${(g.radonBaseIds||[]).length?(g.radonBaseIds||[]).map(x=>`<code>${adminEsc(x)}</code>`).join(' '):'None'}</div></div>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;align-items:start;">
            <div><div style="font-weight:700;color:var(--navy);margin-bottom:7px;">Accession Tests</div>${tests || '<div>—</div>'}</div>
            <div><div style="font-weight:700;color:var(--navy);margin-bottom:7px;">Records Found</div><table><tbody>${listRows}</tbody></table></div>
          </div>
          ${sheetRows?`<div style="margin-top:14px;"><div style="font-weight:700;color:var(--navy);margin-bottom:7px;">Control Sheet Entries</div><div style="overflow-x:auto;"><table><thead><tr><th>Sheet</th><th>Lab ID</th><th>File</th><th>Row</th></tr></thead><tbody>${sheetRows}</tbody></table></div></div>`:''}
          <div style="margin-top:16px;">
            <label style="display:block;font-size:12px;font-weight:700;color:var(--slate);margin-bottom:5px;">Reason for deletion <span style="color:var(--error);">*</span></label>
            <textarea id="admin-delete-reason" rows="3" placeholder="Example: Test sample used for approval workflow validation" style="width:100%;padding:10px 12px;border:1.5px solid var(--border);border-radius:8px;font-family:var(--font-body);font-size:13px;resize:vertical;"></textarea>
          </div>
          <div style="margin-top:12px;padding:10px 12px;background:#fff3e0;border:1px solid #f5c98a;border-radius:8px;font-size:12px;color:#8a4b08;line-height:1.45;">
            A 30-day recovery snapshot is created and verified <strong>before</strong> operational records are removed. Existing Activity Log history is preserved.
          </div>
          <button class="btn-ok" style="margin-top:14px;background:var(--error);width:100%;" onclick="adminDeleteSampleGroup()">Delete This Accession Group</button>
        </div>`;
      preview.style.display='block';
      status.textContent = `Found ${g.baseIds.length} related Lab ID${g.baseIds.length===1?'':'s'}. Review everything below before deleting.`;
    } catch(e) {
      status.textContent = '⚠️ ' + e.message;
      showToast('⚠️ ' + e.message);
    }
  }
  window.adminFindSampleGroup = adminFindSampleGroup;

  async function adminDeleteSampleGroup() {
    const g = _adminPreviewGroup;
    if (!g) { showToast('Find a sample group first.'); return; }
    const reason = String(document.getElementById('admin-delete-reason')?.value || '').trim();
    if (reason.length < 3) { showToast('A deletion reason is required.'); return; }
    const ids = (g.baseIds || []).join(', ');
    if (!confirm(`DELETE ACCESSION GROUP?\n\nCreatedAt: ${g.createdAt}\nLab IDs: ${ids}\n\nA 30-day recovery snapshot will be created first.\n\nReason: ${reason}`)) return;
    const status = document.getElementById('admin-cleanup-status');
    status.textContent='Creating recovery snapshot, then deleting linked records...';
    try {
      const data = await adminApi('delete',{ baseId:g.requestedBaseId, createdAt:g.createdAt, reason });
      document.getElementById('admin-cleanup-preview').style.display='none';
      _adminPreviewGroup=null;
      status.textContent=`✅ Deleted. Recovery ID: ${data.deletionId}. Recoverable until ${adminFmtDT(data.expiresAt)}.`;
      showToast('✅ Sample group deleted with 30-day recovery');
      await adminLoadRecoveries();
    } catch(e) {
      status.textContent='⚠️ ' + e.message;
      showToast('⚠️ ' + e.message);
    }
  }
  window.adminDeleteSampleGroup = adminDeleteSampleGroup;

  async function adminLoadRecoveries() {
    const body = document.getElementById('admin-recovery-body');
    if (!body) return;
    if (!adminToken()) { body.innerHTML='<tr><td colspan="7"><div class="empty-state">Admin verification required.</div></td></tr>'; return; }
    body.innerHTML='<tr><td colspan="7"><div class="empty-state">Loading recoveries...</div></td></tr>';
    try {
      const data = await adminApi('recoveries');
      const rows = data.recoveries || [];
      if (!rows.length) { body.innerHTML='<tr><td colspan="7"><div class="empty-state">No deleted samples are currently recoverable.</div></td></tr>'; return; }
      body.innerHTML = rows.map(r=>`<tr>
        <td><code>${adminEsc(r.requestedBaseId)}</code></td>
        <td>${(r.baseIds||[]).map(x=>`<code>${adminEsc(x)}</code>`).join(' ')}</td>
        <td style="font-size:12px;">${adminEsc(adminFmtDT(r.deletedAt))}</td>
        <td style="font-size:12px;">${adminEsc(adminFmtDT(r.expiresAt))}</td>
        <td>${adminEsc(r.deletedBy)}</td>
        <td style="max-width:260px;white-space:normal;">${adminEsc(r.reason)}</td>
        <td><button class="action-btn" onclick='adminRestoreSample(${JSON.stringify(JSON.stringify(r.path))})'>↩ Restore</button></td>
      </tr>`).join('');
    } catch(e) {
      body.innerHTML=`<tr><td colspan="7"><div class="empty-state">⚠️ ${adminEsc(e.message)}</div></td></tr>`;
    }
  }
  window.adminLoadRecoveries = adminLoadRecoveries;

  async function adminRestoreSample(pathJson) {
    const path = JSON.parse(pathJson);
    try {
      const p = await adminApi('restore-preview',{path});
      const s = p.snapshot;
      const counts = Object.entries(s.counts||{}).map(([k,v])=>`${k}: ${v}`).join('\n');
      if (!confirm(`RESTORE DELETED SAMPLE?\n\nRequested Lab ID: ${s.requestedBaseId}\nRelated Lab IDs: ${(s.baseIds||[]).join(', ')}\nOriginal reason: ${s.reason}\nDeleted by: ${s.deletedBy}\n\n${counts}\nControl sheet entries: ${(s.controlSheets||[]).length}\n\nThis will recreate the operational records and write a restoration audit entry.`)) return;
      const data = await adminApi('restore',{path});
      showToast('✅ Deleted sample restored');
      await adminLoadRecoveries();
    } catch(e) { showToast('⚠️ ' + e.message); }
  }
  window.adminRestoreSample = adminRestoreSample;

'''
if js_anchor not in s: raise SystemExit('JS anchor not found')
s = s.replace(js_anchor, js + js_anchor, 1)

p.write_text(s)
