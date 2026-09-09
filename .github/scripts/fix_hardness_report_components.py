from pathlib import Path

p = Path('api/src/generate-report.js')
s = p.read_text()

old = """      const activeParams = PARAM_CONFIG.filter(p => needed.has(p.name));\n      const fhaParams    = PARAM_CONFIG.filter(p => FHA_PARAM_NAMES.includes(p.name));\n"""
new = """      // Hardness is calculated from Calcium and Magnesium. Whenever hardness is\n      // included in a package or ordered by itself, show all three related results\n      // on the report so the calculation inputs are visible to the client.\n      if (needed.has('Hardness by calculation')) {\n        needed.add('Calcium, Total');\n        needed.add('Magnesium, Total');\n        context.log('[gen] Hardness included → forcing Calcium, Total + Magnesium, Total onto report');\n      }\n\n      const activeParams = PARAM_CONFIG.filter(p => needed.has(p.name));\n      const fhaParams    = PARAM_CONFIG.filter(p => FHA_PARAM_NAMES.includes(p.name));\n"""

if old not in s:
    raise SystemExit('Target activeParams block not found')

s = s.replace(old, new, 1)
p.write_text(s)
