from pathlib import Path
p=Path('admin-dashboard.html')
s=p.read_text()
start=s.find('    <div id="scan-timing-panel"')
if start<0: raise SystemExit('timing panel not found')
end=s.find('    <div class="summary-row"', start)
if end<0: raise SystemExit('summary anchor not found')
s=s[:start]+s[end:]

start=s.find('  function showScanTiming(result) {')
if start<0: raise SystemExit('showScanTiming not found')
end=s.find('  async function triggerScanNow(btn) {', start)
if end<0: raise SystemExit('triggerScanNow anchor not found')
s=s[:start]+s[end:]

s=s.replace("      const timedResult = (data.results || []).find(r => r && r.timing);\n      if (timedResult) showScanTiming(timedResult);\n",'',1)

p.write_text(s)
print('Removed COC scan timing UI')