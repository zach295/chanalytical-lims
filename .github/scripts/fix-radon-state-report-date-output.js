const fs = require('fs');
const path = 'admin-dashboard.html';
let s = fs.readFileSync(path, 'utf8');

const oldCalc = `      const toExcelDate = d => Math.floor((new Date(d + 'T00:00:00') - new Date('1899-12-30')) / 86400000);\n\n      const dataRows = rows.map(r => {`;
const newCalc = `      const formatReportDate = d => {\n        const iso = toISO(d);\n        if (!iso) return '';\n        const [y,m,day] = iso.split('-');\n        return m + '/' + day + '/' + y;\n      };\n\n      const dataRows = rows.map(r => {`;
if (!s.includes(oldCalc)) throw new Error('Expected Excel date helper block not found');
s = s.replace(oldCalc, newCalc);

const oldReturn = `        const excelDate = calcDate ? toExcelDate(calcDate) : '';\n        return [\n          'LAB480', 'Water',\n          client.radonLic || '',\n          client.billingZip || '',\n          r.labId || '',\n          (() => { const n = parseFloat(r.rwResults); return isNaN(n) ? r.rwResults : n < 100 ? '<100' : Math.round(n / 100) * 100; })(),\n          r.location || '',\n          town,\n          excelDate || '',\n          '', '',\n        ];`;
const newReturn = `        const reportDateDisplay = calcDate ? formatReportDate(calcDate) : '';\n        return [\n          'LAB480', 'Water',\n          client.radonLic || '',\n          client.billingZip || '',\n          r.labId || '',\n          (() => { const n = parseFloat(r.rwResults); return isNaN(n) ? r.rwResults : n < 100 ? '<100' : Math.round(n / 100) * 100; })(),\n          r.location || '',\n          town,\n          reportDateDisplay,\n          '', '',\n        ];`;
if (!s.includes(oldReturn)) throw new Error('Expected Radon data row date output block not found');
s = s.replace(oldReturn, newReturn);

const oldFormat = `      // Format Report Date column (index 8) as date\n      const dateColLetter = 'I';\n      for (let i = 1; i <= dataRows.length; i++) {\n        const cell = ws[\`${dateColLetter}\${i+1}\`];\n        if (cell && typeof cell.v === 'number') {\n          cell.t = 'n';\n          cell.z = 'MM/DD/YYYY';\n        }\n      }\n\n`;
if (!s.includes(oldFormat)) throw new Error('Expected Excel cell formatting block not found');
s = s.replace(oldFormat, '');

const fnStart = s.indexOf('async function generateRadonStateReport()');
if (fnStart < 0) throw new Error('Radon state report function not found');
const fnChunk = s.slice(fnStart, fnStart + 9000);
if (!fnChunk.includes('formatReportDate') || !fnChunk.includes('reportDateDisplay')) throw new Error('String report date output not present');
if (fnChunk.includes('toExcelDate')) throw new Error('Old Excel serial conversion still present in Radon report function');

fs.writeFileSync(path, s);
console.log('Radon State Report now writes Report Date as MM/DD/YYYY text');
