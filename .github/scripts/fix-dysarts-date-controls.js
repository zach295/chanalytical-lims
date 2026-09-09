const fs = require('fs');
const path = 'admin-dashboard.html';
let s = fs.readFileSync(path, 'utf8');

const oldBlock = `  async function generateDysartsReport() {\n    const fromVal  = document.getElementById('radon-date-from')?.value;\n    const toVal    = document.getElementById('radon-date-to')?.value;\n    const company  = document.getElementById('radon-company')?.value || '';\n    if (!fromVal || !toVal) { showToast('Select a date range first.'); return; }\n    if (!company) { showToast('Select a company first.'); return; }`;

const newBlock = `  async function generateDysartsReport() {\n    const fromVal  = document.getElementById('report-date-from')?.value;\n    const toVal    = document.getElementById('report-date-to')?.value;\n    if (!fromVal || !toVal) { showToast('Select a date range first.'); return; }`;

if (!s.includes(oldBlock)) throw new Error('Expected Dysart report opening block not found');
s = s.replace(oldBlock, newBlock);

const fnStart = s.indexOf('async function generateDysartsReport()');
if (fnStart < 0) throw new Error('Dysart report function missing after patch');
const fnChunk = s.slice(fnStart, fnStart + 3000);
if (!fnChunk.includes("report-date-from") || !fnChunk.includes("report-date-to")) throw new Error('Dysart date controls not present');
if (fnChunk.includes("radon-company") || fnChunk.includes("Select a company first.")) throw new Error('Radon company dependency still present in Dysart function opening');

fs.writeFileSync(path, s);
console.log('Dysart report now uses its own date controls and has no Radon company requirement');
