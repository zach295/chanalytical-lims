const assert = require('node:assert/strict');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const { verifyVisibleReportPdf } = require('../shared/pdf-visual-check');

async function fixture(label, value, color) {
  const doc = await PDFDocument.create();
  const page = doc.addPage([612, 792]);
  const font = await doc.embedFont(StandardFonts.Helvetica);
  page.drawText(label, { x: 70, y: 500, size: 10, font });
  page.drawRectangle({ x: 270, y: 496, width: 20, height: 12, color });
  page.drawText(value, { x: 300, y: 500, size: 10, font });
  return Buffer.from(await doc.save());
}

(async () => {
  for (const [value, color] of [['<0.2', rgb(0,1,0)], ['0.5', rgb(0,1,0)],
    ['2', rgb(11/255,83/255,148/255)], ['4', rgb(1,0,0)]]) {
    const params = [{ name: 'Fluoride, Total', value, previewColor: value === '2' ? '#0B5394' : value === '4' ? '#FF0000' : '#00FF00' }];
    for (const label of ['Fluoride', 'Fluoride, Total']) {
      const result = await verifyVisibleReportPdf(await fixture(label, value, color), params);
      assert.equal(result.verified, 1);
    }
  }
  const params = [{ name: 'Fluoride, Total', value: '<0.2', previewColor: '#00FF00' }];
  const wrongColor = await fixture('Fluoride', '<0.2', rgb(1,0,0));
  await assert.rejects(() => verifyVisibleReportPdf(wrongColor, params), /color mismatch/);
  const wrongValue = await fixture('Fluoride', '<0.3', rgb(0,1,0));
  await assert.rejects(() => verifyVisibleReportPdf(wrongValue, params), /result missing/);
  const wrongName = await fixture('Fluoride, Free', '<0.2', rgb(0,1,0));
  await assert.rejects(() => verifyVisibleReportPdf(wrongName, params), /parameter not found/);
  console.log('Fluoride PDF alias, value and color checks passed');
})().catch(error => { console.error(error); process.exitCode = 1; });
