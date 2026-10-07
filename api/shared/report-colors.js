'use strict';
// Single source of truth for preview and workbook result indicator colors.
// Exact reporting-limit strings are intentionally not numerically normalized.
function calcFillColor(paramName, displayVal) {
  if (!displayVal && displayVal !== 0) return null;
  const s   = String(displayVal).trim();
  if (!s) return null;
  const exactLimits = {
    'Chloride, Total':'2', 'Fluoride, Total':'0.2', 'Nitrite-Nitrogen, Total':'0.2',
    'Nitrate-Nitrogen, Total':'1', 'Arsenic, Total':'1', 'Lead, Total':'1',
    'Uranium, Total':'1', 'Copper, Total':'0.001', 'Iron, Total':'0.05',
    'Manganese, Total':'0.001', 'Antimony, Total':'0.0005',
    'Cadmium, Total':'0.002', 'Chromium, Total':'0.002',
    'Sulfate':'40', 'Total Coliform':'1', 'E. Coli':'1',
    'Radon Water':'100', 'Turbidity':'1', 'Arsenic III':'1', 'Arsenic V':'1',
    'Arsenic, III':'1', 'Arsenic, V':'1', 'Arsenic, Speciation':'1'
  };
  if (s.startsWith('<')) return s === '<' + exactLimits[paramName] ? '#00CC44' : '#FF0000';
  if (!/^[+-]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)$/.test(s)) return '#FF0000';
  const n = Number(s);
  const num = Number.isFinite(n);
  const rl = false;

  switch (paramName) {
    case 'Chloride, Total':
      // green: <2 (string) OR <250;  blue: >=250
      if (rl || (num && n < 250))  return '#00CC44';
      if (num && n >= 250)          return '#0070C0';
      return null;

    case 'Fluoride, Total':
      // green: <0.2 (string) OR <1.9;  blue: 1.9–3.9;  red: >=4
      if (rl || (num && n < 1.9))            return '#00CC44';
      if (num && n >= 1.9 && n <= 3.9)       return '#0070C0';
      if (num && n >= 4)                      return '#FF0000';
      return null;

    case 'Nitrite-Nitrogen, Total':
      // green: <0.2 (string) OR <1;  red: >=1
      if (rl || (num && n < 1))   return '#00CC44';
      if (num && n >= 1)           return '#FF0000';
      return null;

    case 'Nitrate-Nitrogen, Total':
      // green: <1 (string) OR <10;  red: >=10
      if (rl || (num && n < 10))  return '#00CC44';
      if (num && n >= 10)          return '#FF0000';
      return null;

    case 'Arsenic, Total':
      // green: <1 (string) OR <10;  red: >=10
      if (rl || (num && n < 10))  return '#00CC44';
      if (num && n >= 10)          return '#FF0000';
      return null;

    case 'Lead, Total':
      // green: <1 (string) OR <15;  red: >=15
      if (rl || (num && n < 15))  return '#00CC44';
      if (num && n >= 15)          return '#FF0000';
      return null;

    case 'Uranium, Total':
      // green: <1 (string) OR <30;  red: >=30
      if (rl || (num && n < 30))  return '#00CC44';
      if (num && n >= 30)          return '#FF0000';
      return null;

    case 'Copper, Total':
      // green: <0.001 (string) OR <0.9;  blue: 0.9–1.29;  red: >=1.3
      if (rl || (num && n < 0.9))            return '#00CC44';
      if (num && n >= 0.9 && n <= 1.29)      return '#0070C0';
      if (num && n >= 1.3)                    return '#FF0000';
      return null;

    case 'Iron, Total':
      // green: <0.05 (string) OR <0.3;  blue: >=0.3
      if (rl || (num && n < 0.3))  return '#00CC44';
      if (num && n >= 0.3)          return '#0070C0';
      return null;

    case 'Manganese, Total':
      // green: <0.001 (string) OR <0.05;  blue: >=0.05
      if (rl || (num && n < 0.05))  return '#00CC44';
      if (num && n >= 0.05)          return '#0070C0';
      return null;

    case 'Sodium, Total':
      // blue: >=20 only — no color defined below 20 in template
      if (num && n >= 20)  return '#0070C0';
      return null;

    case 'Antimony, Total':
      // green: <0.0005 (string) OR <0.006;  red: >=0.006
      if (rl || (num && n < 0.006))  return '#00CC44';
      if (num && n >= 0.006)          return '#FF0000';
      return null;

    case 'Cadmium, Total':
      // green: <0.002 (string) OR <0.005;  red: >=0.005
      if (rl || (num && n < 0.005))  return '#00CC44';
      if (num && n >= 0.005)          return '#FF0000';
      return null;

    case 'Chromium, Total':
      // green: <0.002 (string) OR <0.1;  red: >=0.1
      if (rl || (num && n < 0.1))  return '#00CC44';
      if (num && n >= 0.1)          return '#FF0000';
      return null;

    case 'pH Electrometric':
      // green: 6.5–8.5;  blue: outside range
      if (num && n >= 6.5 && n <= 8.5)  return '#00CC44';
      if (num && (n < 6.5 || n > 8.5))  return '#0070C0';
      return null;

    case 'Sulfate':
      // green: <40 (string) OR <250;  red: >=250
      if (rl || (num && n < 250))  return '#00CC44';
      if (num && n >= 250)          return '#FF0000';
      return null;

    case 'Total Coliform':
      // green: <1 strings OR <1;  blue: >=1
      if (rl || (num && n < 1))  return '#00CC44';
      if (num && n >= 1)          return '#0070C0';
      return null;

    case 'E. Coli':
      // green: <1 strings OR <1;  red: >=1  (E.coli is RED, not blue)
      if (rl || (num && n < 1))  return '#00CC44';
      if (num && n >= 1)          return '#FF0000';
      return null;

    case 'Radon Water':
      // green: <100 (string) OR <4000;  blue: >=4000
      if (rl || (num && n < 4000))   return '#00CC44';
      if (num && n >= 4000)           return '#0070C0';
      return null;

    case 'Turbidity':
      // green: <1 string OR <1;  blue: >=1
      if (rl || (num && n < 1))  return '#00CC44';
      if (num && n >= 1)          return '#0070C0';
      return null;
    default:
      return null;
  }
}


module.exports = { calcFillColor };
