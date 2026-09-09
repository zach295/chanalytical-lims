const fs = require('fs');
const path = 'admin-dashboard.html';
let s = fs.readFileSync(path, 'utf8');

const oldBlock = `      const formatReportDate = d => {\n        const iso = toISO(d);\n        if (!iso) return '';\n        const [y,m,day] = iso.split('-');\n        return m + '/' + day + '/' + y;\n      };`;

const newBlock = `      const formatReportDate = d => {\n        if (!d) return '';\n        const raw = String(d).trim();\n        const iso = raw.match(/^(\\d{4})-(\\d{1,2})-(\\d{1,2})/);\n        if (iso) return iso[2].padStart(2,'0') + '/' + iso[3].padStart(2,'0') + '/' + iso[1];\n        const us = raw.match(/^(\\d{1,2})\\/(\\d{1,2})\\/(\\d{2}|\\d{4})$/);\n        if (us) {\n          const year = us[3].length === 2 ? '20' + us[3] : us[3];\n          return us[1].padStart(2,'0') + '/' + us[2].padStart(2,'0') + '/' + year;\n        }\n        return raw;\n      };`;

if (!s.includes(oldBlock)) throw new Error('Expected Radon report date formatter not found');
s = s.replace(oldBlock, newBlock);

const fnStart = s.indexOf('async function generateRadonStateReport()');
if (fnStart < 0) throw new Error('Radon state report function not found');
const fnChunk = s.slice(fnStart, fnStart + 8000);
if (!fnChunk.includes("raw.match(/^(\\d{1,2})\\/(\\d{1,2})\\/(\\d{2}|\\d{4})$/)")) throw new Error('MM/DD/YY formatter validation failed');
if (!fnChunk.includes("'20' + us[3]")) throw new Error('Two-digit year expansion missing');

fs.writeFileSync(path, s);
console.log('Radon State Report date formatter now handles ISO and MM/DD/YY dates');
