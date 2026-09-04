from pathlib import Path
p=Path('admin-dashboard.html')
s=p.read_text()
old='''        <button class="action-btn" onclick="refreshQueueSafe()">↻ Refresh Queue</button>\n\n        <span id="create-cs-status" style="font-size:12px;color:var(--slate);"></span>'''
new='''        <button class="action-btn" onclick="refreshQueueSafe()">↻ Refresh Queue</button>\n\n        <span id="create-cs-status" style="font-size:12px;color:var(--slate);"></span>'''
if old not in s: raise SystemExit('header anchor not found')
# Add panel after section header
anchor='''    </div>\n\n    <div class="summary-row" style="grid-template-columns:repeat(2,1fr);margin-bottom:20px;">'''
panel='''    </div>\n\n    <div id="scan-timing-panel" style="display:none;background:var(--white);border:1px solid var(--border);border-left:4px solid var(--teal);border-radius:8px;padding:12px 16px;margin-bottom:16px;box-shadow:var(--shadow);">\n      <div style="display:flex;justify-content:space-between;gap:12px;align-items:center;flex-wrap:wrap;">\n        <div style="font-size:13px;font-weight:700;color:var(--navy);">⏱️ Last COC Scan Timing</div>\n        <div id="scan-timing-total" style="font-size:16px;font-weight:700;color:var(--teal);"></div>\n      </div>\n      <div id="scan-timing-detail" style="font-size:12px;color:var(--slate);margin-top:7px;line-height:1.7;"></div>\n    </div>\n\n    <div class="summary-row" style="grid-template-columns:repeat(2,1fr);margin-bottom:20px;">'''
if anchor not in s: raise SystemExit('panel anchor not found')
s=s.replace(anchor,panel,1)
# Add renderer before triggerScanNow
anchor2='''  async function triggerScanNow(btn) {'''
renderer='''  function showScanTiming(result) {\n    const panel = document.getElementById('scan-timing-panel');\n    const totalEl = document.getElementById('scan-timing-total');\n    const detailEl = document.getElementById('scan-timing-detail');\n    if (!panel || !result?.timing) return;\n    const t = result.timing;\n    const sec = ms => ((Number(ms)||0)/1000).toFixed(1) + 's';\n    const aiMs = (Number(t.haikuMs)||0) + (Number(t.sonnetRetryMs)||0);\n    const spMs = (Number(t.moveMs)||0) + (Number(t.downloadMs)||0) + (Number(t.reviewQueueMs)||0);\n    totalEl.textContent = `Total ${sec(t.totalMs)}`;\n    detailEl.innerHTML = [\n      `<strong>Azure OCR:</strong> ${sec(t.azureMs)}`,\n      `<strong>AI extraction:</strong> ${sec(aiMs)}${t.sonnetRetryMs ? ' (includes Sonnet recovery)' : ' (Haiku)'}`,\n      `<strong>SharePoint:</strong> ${sec(spMs)}`,\n      `<strong>Move:</strong> ${sec(t.moveMs)}`,\n      `<strong>Download:</strong> ${sec(t.downloadMs)}`,\n      `<strong>Queue write:</strong> ${sec(t.reviewQueueMs)}`\n    ].join(' &nbsp;·&nbsp; ');\n    panel.style.display = 'block';\n  }\n\n  async function triggerScanNow(btn) {'''
if anchor2 not in s: raise SystemExit('trigger anchor not found')
s=s.replace(anchor2,renderer,1)
# Display timing immediately after response
old3='''      const found = data.newScans ?? data.processed ?? data.count ?? 0;\n      if (found > 0) {'''
new3='''      const found = data.newScans ?? data.processed ?? data.count ?? 0;\n      const timedResult = (data.results || []).find(r => r && r.timing);\n      if (timedResult) showScanTiming(timedResult);\n      if (found > 0) {'''
if old3 not in s: raise SystemExit('response anchor not found')
s=s.replace(old3,new3,1)
p.write_text(s)
print('Added scan timing panel to Review Queue UI')