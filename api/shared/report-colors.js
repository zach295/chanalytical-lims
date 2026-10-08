'use strict';
// Rules transcribed from Report Templates.xlsx (uploaded October 7, 2026).
// Do not infer a green result from EPA thresholds absent from the Excel CF rules.
// Nonmatching populated values are red by laboratory policy.
function calcFillColor(paramName, displayVal, sheetType = 'lab') {
  const value = String(displayVal ?? '').trim();
  if (!value) return null;
  const key = String(paramName || '').replace(/, Total$/i, ', Total');
  const sheet = String(sheetType || 'lab').toLowerCase();
  const spec = sheet.includes('spec');
  const fha = sheet.includes('fha');
  const radon = sheet.includes('radon');
  const green = '#00FF00', red = '#FF0000', blue = '#0B5394', grey = '#E8E8E8';
  const rules = {
    'Chloride, Total': {lt:250, blueAt:250, text:['<2']},
    'Fluoride, Total': {lt:1.9, blueFrom:1.9, blueTo:3.9, redAt:4, text:['<0.2']},
    'Nitrite-Nitrogen, Total': {lt:1, redAt:1, text:['<0.2']},
    'Nitrate-Nitrogen, Total': {lt:10, redAt:10, text:['<1']},
    'Arsenic, Total': {lt:10, redAt:10, text:['<1']},
    'Lead, Total': {lt:15, redAt:15, text:['<1']},
    'Uranium, Total': {lt:30, redAt:30, text:['<1']},
    'Copper, Total': {lt:0.9, blueFrom:0.9, blueTo:1.29, redAt:1.3, text:['<0.001']},
    'Iron, Total': {lt:0.3, blueAt:0.3, text:['<0.05']},
    'Manganese, Total': {lt:0.05, blueAt:0.05, text:['<0.001']},
    'Sodium, Total': {blueAt:20},
    'Antimony, Total': {lt:0.006, redAt:0.006, text:['<0.0005']},
    'Cadmium, Total': {lt:0.005, redAt:0.005, text:['<0.002']},
    // Excel template C33 green condition uses 0.01 (not 0.1).
    'Chromium, Total': {lt:0.01, redAt:0.1, text:['<0.002']},
    'pH Electrometric': {between:[6.5,8.5], blueOutside:[6.5,8.5]},
    'Sulfate': {lt:250, redAt:250, text:['<40']},
    'Total Coliform': {lt:1, blueAt:1, text:['<1.00','<1.0','<1']},
    'E. Coli': {lt:1, redAt:1, text:['<1.00','<1.0','<1']},
    'Turbidity': {lt:1, blueAt:1, text:['<1']},
    'Radon Water': {lt:4000, blueAbove:4000, text:['<100']}
  };
  let r = rules[key === 'Fluoride' ? 'Fluoride, Total' : key];
  const specArsenic = spec && /^(Arsenic,? (?:III|V))$/.test(key);
  if (specArsenic) {
    // Speciation template only defines a red threshold at 10.
    r = {redAt:10};
  } else if (fha) {
    if (!['Nitrite-Nitrogen, Total','Nitrate-Nitrogen, Total','Lead, Total','Total Coliform','E. Coli'].includes(key)) r = null;
    else r = {...r, text: key === 'Total Coliform' || key === 'E. Coli' ? ['<1','<1.0'] : r.text};
  } else if (spec && key !== 'Arsenic, Total') r = null;
  else if (radon && key !== 'Radon Water') r = null;
  // No conditional formatting: neutral grey, including text and malformed values.
  if (!r) return grey;
  if (value.startsWith('<')) return r.text?.includes(value) ? green : red;
  // Reject malformed text instead of letting parseFloat accept partial values.
  if (!/^[+-]?\d+(?:\.\d+)?$/.test(value)) return red;
  const n = Number(value);
  if (!Number.isFinite(n)) return red;
  if (r.between && n >= r.between[0] && n <= r.between[1]) return green;
  if (r.lt !== undefined && n < r.lt) return green;
  if (r.blueFrom !== undefined && n >= r.blueFrom && n <= r.blueTo) return blue;
  if (r.blueAt !== undefined && n >= r.blueAt) return blue;
  // Sodium's Excel template has blue >= 20, with a neutral grey below 20.
  if (key === 'Sodium, Total' && n < 20) return grey;
  if (r.blueAbove !== undefined && n > r.blueAbove) return blue;
  if (r.blueOutside && (n < r.blueOutside[0] || n > r.blueOutside[1])) return blue;
  return red;
}

module.exports = { calcFillColor };
