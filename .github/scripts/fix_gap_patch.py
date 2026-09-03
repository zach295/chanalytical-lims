from pathlib import Path
p = Path('.github/scripts/close_activity_log_gaps.py')
s = p.read_text()
s = s.replace('"const { getToken } = require(\'../shared/graph\');\\n",\n"const { getToken } = require(\'../shared/graph\');\\nconst { writeActivityLog } = require(\'../shared/audit\');\\n",\n\'radon audit import\')', '"const { getToken, createItem, listItems } = require(\'../shared/graph\');\\n",\n"const { getToken, createItem, listItems } = require(\'../shared/graph\');\\nconst { writeActivityLog } = require(\'../shared/audit\');\\n",\n\'radon audit import\')')
p.write_text(s)
