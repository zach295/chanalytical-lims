from pathlib import Path
p = Path('api/src/approve-scan.js')
s = p.read_text()
old = "async function writeRadonControlSheet(siteId, token, labId, dateDrawn, timeDrawn, context) {"
new = "async function writeRadonControlSheet(siteId, token, labId, dateDrawn, timeDrawn, receivedDate, context) {"
if old not in s:
    raise SystemExit('signature anchor not found')
s = s.replace(old, new, 1)
old2 = """    const todayET = new Date().toLocaleDateString('en-US', {
      timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: 'numeric'
    });
    // Format dateDrawn for Excel (MM/DD/YYYY)
    const drawnFmt = fmtExcel(dateDrawn) || dateDrawn || '';
"""
new2 = """    const todayET = new Date().toLocaleDateString('en-US', {
      timeZone: 'America/New_York', month: '2-digit', day: '2-digit', year: 'numeric'
    });
    // Use the reviewed/corrected dates from the approval card. Previously column G
    // always used today's approval date, so correcting Date Received before approval
    // never reached the Radon Control Sheet.
    const drawnFmt    = fmtExcel(dateDrawn) || dateDrawn || '';
    const receivedFmt = fmtExcel(receivedDate) || receivedDate || todayET;
"""
if old2 not in s:
    raise SystemExit('date block anchor not found')
s = s.replace(old2, new2, 1)
old3 = "body: JSON.stringify({ values: [[newLabId, '', '', '', drawnFmt, timeDrawn || '', todayET]] }),"
new3 = "body: JSON.stringify({ values: [[newLabId, '', '', '', drawnFmt, timeDrawn || '', receivedFmt]] }),"
if old3 not in s:
    raise SystemExit('RCS row anchor not found')
s = s.replace(old3, new3, 1)
old4 = """            dateDrawn || '',
            to24h(timeDrawn) || timeDrawn || '',
            context
"""
new4 = """            dateDrawn || '',
            to24h(timeDrawn) || timeDrawn || '',
            receivedDate || '',
            context
"""
if old4 not in s:
    raise SystemExit('call anchor not found')
s = s.replace(old4, new4, 1)
p.write_text(s)
