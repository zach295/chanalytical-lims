const assert=require('node:assert/strict');
const { PDFDocument,StandardFonts,rgb }=require('pdf-lib');
const { verifyVisibleReportPdf, expectedClass, rgbClass }=require('../shared/pdf-visual-check');

assert.equal(rgbClass(255,0,0),'red');
assert.equal(rgbClass(0,255,0),'green');
assert.equal(rgbClass(11,83,148),'blue');
assert.equal(rgbClass(232,232,232),'grey');
assert.equal(expectedClass('#00CC44'),'green');
assert.equal(expectedClass('#0070C0'),'blue');
assert.equal(expectedClass('#E8E8E8'),'grey');

async function fixture(indicatorColor) {
  const doc=await PDFDocument.create();
  const page=doc.addPage([612,792]);
  const font=await doc.embedFont(StandardFonts.Helvetica);
  page.drawText('Parameter',{x:70,y:540,size:10,font});
  page.drawText('Your Result',{x:300,y:540,size:10,font});
  page.drawText('Arsenic, III',{x:70,y:500,size:10,font});
  page.drawRectangle({x:270,y:496,width:20,height:12,color:indicatorColor});
  page.drawText('<1',{x:300,y:500,size:10,font});
  return Buffer.from(await doc.save());
}
(async()=>{
  const param=[{name:'Arsenic, III',value:'<1',sheetType:'spec',previewColor:'#E8E8E8'}];
  const good=await fixture(rgb(232/255,232/255,232/255));
  const verified=await verifyVisibleReportPdf(good,param);
  assert.equal(verified.verified,1);
  const bad=await fixture(rgb(1,0,0));
  await assert.rejects(()=>verifyVisibleReportPdf(bad,param),/Visible PDF color mismatch/);
  console.log('Rendered PDF indicator verification checks passed');
})().catch(e=>{console.error(e);process.exitCode=1});
