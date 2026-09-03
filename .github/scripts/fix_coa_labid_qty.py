from pathlib import Path
p=Path('api/src/approve-scan.js')
s=p.read_text()
s=s.replace("const range = encodeURIComponent(`${SHEETS_TAB}!A:M`);","const range = encodeURIComponent(`${SHEETS_TAB}!A:N`);",1)
s=s.replace("              l.fullId,\n              location   || '',","              l.baseId,\n              location   || '',",1)
s=s.replace("              coaTestName,\n            ]);","              coaTestName,\n              1,\n            ]);",1)
p.write_text(s)
