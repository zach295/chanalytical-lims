const assert = require('node:assert/strict');
const { calcFillColor } = require('../shared/report-colors');
const green = '#00FF00', red = '#FF0000';
assert.equal(calcFillColor('Fluoride, Total', '<0.2'), green);
assert.equal(calcFillColor('Fluoride, Total', '<0.20'), red);
assert.equal(calcFillColor('Fluoride, Total', '<0.3'), red);
assert.equal(calcFillColor('Fluoride, Total', 'ND'), red);
assert.equal(calcFillColor('Fluoride, Total', 'N/A'), red);
assert.equal(calcFillColor('Fluoride, Total', '0.'), red);
assert.equal(calcFillColor('Fluoride, Total', '.5'), red);
assert.equal(calcFillColor('Fluoride, Total', '0..2'), red);
assert.equal(calcFillColor('Fluoride, Total', '0'), green);
assert.equal(calcFillColor('Fluoride, Total', '0.0'), green);
assert.equal(calcFillColor('Fluoride, Total', '0.5'), green);
assert.equal(calcFillColor('Fluoride, Total', '4'), red);
assert.equal(calcFillColor('Fluoride, Total', '2'), '#0B5394');
assert.equal(calcFillColor('Chromium, Total', '0.05'), red);
assert.equal(calcFillColor('Total Coliform', '<1.0'), green);
assert.equal(calcFillColor('Total Coliform', '<1.00'), green);
assert.equal(calcFillColor('Sodium, Total', '21'), '#0B5394');
assert.equal(calcFillColor('Sodium, Total', '19.9'), '#E8E8E8');
assert.equal(calcFillColor('Sodium, Total', '0'), '#E8E8E8');
assert.equal(calcFillColor('Sodium, Total', '20'), '#0B5394');
for (const param of ['Hardness by calculation', 'Calcium, Total', 'Magnesium, Total', 'Alkalinity']) {
  assert.equal(calcFillColor(param, '1.2'), green, param + ' numeric');
  assert.equal(calcFillColor(param, 'ND'), green, param + ' text');
  assert.equal(calcFillColor(param, '0.'), green, param + ' malformed');
}

assert.equal(calcFillColor('Chloride, Total', '<2'), green);
assert.equal(calcFillColor('Chloride, Total', '<2.00'), red);
assert.equal(calcFillColor('Arsenic, Total', '<1'), green);
assert.equal(calcFillColor('Arsenic, Total', '<1.0'), red);
assert.equal(calcFillColor('E. Coli', '<1'), green);
assert.equal(calcFillColor('E. Coli', '<2'), red);
assert.equal(calcFillColor('pH Electrometric', 'ND'), red);
console.log('Strict reporting-limit color regression checks passed');
