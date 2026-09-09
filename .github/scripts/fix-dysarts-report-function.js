const fs = require('fs');
const path = 'admin-dashboard.html';
let s = fs.readFileSync(path, 'utf8');

// 1) Repoint only the Dysart's Run card button.
const dysartsCard = `<!-- Dysart's Run Report -->`;
const cardIdx = s.indexOf(dysartsCard);
if (cardIdx < 0) throw new Error("Dysart's Run card anchor not found");
const buttonNeedle = `onclick="generateMonthlyReport()"`;
const buttonIdx = s.indexOf(buttonNeedle, cardIdx);
if (buttonIdx < 0) throw new Error('Dysart report button call not found');
s = s.slice(0, buttonIdx) + `onclick="generateDysartsReport()"` + s.slice(buttonIdx + buttonNeedle.length);

// 2) Find the existing Dysart report implementation by its unique report body,
// then rename only the function that owns that implementation.
const marker = `const dedupedRows = Object.values(kitMap).map(r => {`;
const markerIdx = s.indexOf(marker);
if (markerIdx < 0) throw new Error('Dysart report implementation marker not found');

const before = s.slice(0, markerIdx);
const fnMatches = [...before.matchAll(/async function\s+([A-Za-z0-9_$]+)\s*\(\)\s*\{/g)];
if (!fnMatches.length) throw new Error('No async function found before Dysart implementation');
const owner = fnMatches[fnMatches.length - 1];
const oldFnName = owner[1];
if (oldFnName === 'generateRadonStateReport') throw new Error('Refusing to rename Radon report function');
const ownerStart = owner.index;
const ownerText = owner[0];
const renamed = ownerText.replace(`function ${oldFnName}`, 'function generateDysartsReport');
s = s.slice(0, ownerStart) + renamed + s.slice(ownerStart + ownerText.length);

// Guardrails: verify the renamed function owns the unique Dysart implementation marker
// and that the four-company filter remains present after the function declaration.
const dysFnIdx = s.indexOf('async function generateDysartsReport()');
const newMarkerIdx = s.indexOf(marker);
if (dysFnIdx < 0 || newMarkerIdx < dysFnIdx) throw new Error('Dysart implementation was not renamed correctly');
const between = s.slice(dysFnIdx, newMarkerIdx);
if (!between.includes('NORTHERN_CLIENTS')) throw new Error('Four-company Dysart filter not found in renamed function');

fs.writeFileSync(path, s);
console.log(`Renamed Dysart report function from ${oldFnName} to generateDysartsReport and repointed its button`);
