from pathlib import Path

# 1) Fix forced password setup payload in login.html
p = Path('login.html')
s = p.read_text()
old = "body: JSON.stringify({ action: 'setpw', email, newPassword: pw }),"
new = "body: JSON.stringify({ action: 'setpw', email, password: pw }),"
if old not in s:
    raise SystemExit('login.html setpw payload anchor not found')
s = s.replace(old, new, 1)
p.write_text(s)

# 2) Make users-manage accept the old/new field name and reject blank passwords
p = Path('api/src/users-manage.js')
s = p.read_text()
old = "const { email: pwEmail, password: newPw } = body;\n        const user = await findUserByEmail(pwEmail);"
new = "const pwEmail = body.email;\n        const newPw = body.password || body.newPassword || '';\n        if (newPw.length < 6) return { status: 400, body: JSON.stringify({ error: 'Password must be at least 6 characters.' }) };\n        const user = await findUserByEmail(pwEmail);"
if old not in s:
    raise SystemExit('users-manage setpw anchor not found')
s = s.replace(old, new, 1)
p.write_text(s)

# 3) Make Admin unlock use the same login path, then sync the password server-side
p = Path('admin-dashboard.html')
s = p.read_text()
old = """    try {\n      const res = await fetch('/api/sample-admin', {\n        method:'POST', headers:{'Content-Type':'application/json'},\n        body:JSON.stringify({ action:'verify', email:s?.email || '', password:pw })\n      });\n      const data = await res.json();\n      if (!res.ok || !data.success) throw new Error(data.error || 'Admin verification failed.');\n"""
new = """    try {\n      // Verify through the same authentication path used by the normal LIMS login.\n      // This also supports accounts whose last forced password change was saved locally\n      // by the older reset flow before the server-side password field was corrected.\n      const loginResult = await Auth.loginAsync(s?.email || '', pw);\n      if (!loginResult?.success) throw new Error(loginResult?.error || 'Incorrect LIMS password.');\n      const verifiedRoles = String(loginResult.user?.role || '').split(',').map(v=>v.trim());\n      if (!verifiedRoles.includes('admin')) throw new Error('This account does not have Admin access.');\n\n      // Keep the Users-list password hash synchronized with the password that just\n      // successfully authenticated in the LIMS. This repairs older forced-reset accounts.\n      const syncRes = await fetch('/api/users-manage', {\n        method:'POST', headers:{'Content-Type':'application/json'},\n        body:JSON.stringify({ action:'setpw', email:s?.email || '', password:pw, updatedBy:s?.name || s?.email || 'Admin' })\n      });\n      const syncData = await syncRes.json().catch(()=>({}));\n      if (!syncRes.ok || !syncData.success) throw new Error(syncData.error || 'Could not synchronize the LIMS password.');\n\n      // Now issue the short-lived server-side Admin token.\n      const res = await fetch('/api/sample-admin', {\n        method:'POST', headers:{'Content-Type':'application/json'},\n        body:JSON.stringify({ action:'verify', email:s?.email || '', password:pw })\n      });\n      const data = await res.json();\n      if (!res.ok || !data.success) throw new Error(data.error || 'Admin verification failed.');\n"""
if old not in s:
    raise SystemExit('admin verification anchor not found')
s = s.replace(old, new, 1)
p.write_text(s)
