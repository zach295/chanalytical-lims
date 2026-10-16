'use strict';
// Render the actual PDF, including Excel conditional formatting, before releasing it.
// Unlike Graph /format/fill, this checks the visible pixels in the exported PDF.
// Any missing/ambiguous indicator or mismatch fails closed.
const { calcFillColor } = require('./report-colors');

function rgbClass(r,g,b) {
  const max=Math.max(r,g,b), min=Math.min(r,g,b);
  if (r>=180 && r>g*1.55 && r>b*1.55) return 'red';
  if (g>=110 && g>r*1.5 && g>b*1.25) return 'green';
  if (b>=85 && b>r*1.25 && b>g*1.08) return 'blue';
  if (max>=195 && max<=245 && max-min<=13) return 'grey';
  return null;
}
function expectedClass(hex) {
  const value=String(hex||'').toUpperCase();
  return ({'#FF0000':'red','#00FF00':'green','#0B5394':'blue','#E8E8E8':'grey'})[value]||null;
}
function normalize(s) {
  return String(s??'').toLowerCase().replace(/[^a-z0-9]/g,'');
}
function pixelCounts(image, xMin, xMax, yCenter) {
  const counts={red:0,green:0,blue:0,grey:0};
  const w=image.width,h=image.height, data=image.data;
  // Scan only the middle of the result's indicator cell, avoiding grid borders.
  const y0=Math.max(0,Math.floor(yCenter-10));
  const y1=Math.min(h-1,Math.ceil(yCenter+3));
  for(let y=y0;y<=y1;y++){
    for(let x=Math.max(0,Math.floor(xMin));x<=Math.min(w-1,Math.ceil(xMax));x++){
      const i=(y*w+x)*4;
      if(data[i+3]<240)continue;
      const kind=rgbClass(data[i],data[i+1],data[i+2]);
      if(kind)counts[kind]++;
    }
  }
  return counts;
}
function classifyIndicator(image, resultX, baselineY) {
  // The indicator is the narrow filled cell immediately LEFT of "Your Result".
  // The right edge of the colored cell borders the result value column.
  // Scan a bounded strip, rather than the entire parameter row/legend.
  const counts=pixelCounts(image,resultX-65,resultX-5,baselineY);
  const entries=Object.entries(counts).sort((a,b)=>b[1]-a[1]);
  const [winner,count]=entries[0];
  if(count<12)return {color:null,counts};
  // If another distinct visible fill occupies a significant area, do not guess.
  if(entries[1][1]>Math.max(15,count*0.6))return {color:null,counts};
  return {color:winner,counts};
}
async function verifyVisibleReportPdf(pdfBuffer, params, options={}) {
  const { createCanvas }=require('@napi-rs/canvas');
  const pdfjs=await import('pdfjs-dist/legacy/build/pdf.mjs');
  const loading=pdfjs.getDocument({
    data:new Uint8Array(pdfBuffer),useSystemFonts:true,
    disableFontFace:true,stopAtErrors:true,
    // Rendering should not load external assets.
    isEvalSupported:false,
  });
  const doc=await loading.promise;
  try {
    const expected=(params||[]).filter(p=>String(p?.value??'').trim() && calcFillColor(p.name,p.value,p.sheetType||'lab'));
    const seen=new Map();
    for(let pageNumber=1;pageNumber<=doc.numPages;pageNumber++){
      const page=await doc.getPage(pageNumber);
      const viewport=page.getViewport({scale:2});
      const canvas=createCanvas(Math.ceil(viewport.width),Math.ceil(viewport.height));
      const ctx=canvas.getContext('2d');
      await page.render({canvasContext:ctx,viewport,canvas}).promise;
      const image=ctx.getImageData(0,0,canvas.width,canvas.height);
      const tc=await page.getTextContent();
      const rows=[];
      for(const item of tc.items){
        if(!item.str || !item.transform)continue;
        const pos=pdfjs.Util.transform(viewport.transform,item.transform);
        const x=pos[4],y=pos[5];
        let row=rows.find(r=>Math.abs(r.y-y)<4);
        if(!row){row={y,items:[]};rows.push(row);}
        row.items.push({x,y,str:String(item.str),width:(item.width||0)*2});
      }
      for(const row of rows){
        row.items.sort((a,b)=>a.x-b.x);
        const line=normalize(row.items.map(t=>t.str).join(' '));
        for(const p of expected){
          const name=normalize(p.name);
          if(!line.startsWith(name))continue;
          const value=String(p.value).trim();
          // Find the result itself, not the EPA limit or method field.
          const nameEnd=row.items.find(t=>normalize(t.str).includes(name))?.x || row.items[0].x;
          const matches=row.items.filter(t=>t.str.trim()===value && t.x>nameEnd+8);
          if(!matches.length)throw new Error(`Visible PDF result missing for ${p.name} (page ${pageNumber})`);
          const result=matches[0];
          const actual=classifyIndicator(image,result.x,row.y);
          const expectedColor=expectedClass(calcFillColor(p.name,p.value,p.sheetType||'lab'));
          if(!actual.color)throw new Error(`Visible PDF indicator cannot be verified for ${p.name} (page ${pageNumber}, pixel counts ${JSON.stringify(actual.counts)})`);
          if(actual.color!==expectedColor)throw new Error(`Visible PDF color mismatch for ${p.name}: preview ${expectedColor}, rendered PDF ${actual.color} (page ${pageNumber})`);
          seen.set(p.name,true);
        }
      }
      page.cleanup();
    }
    for(const p of expected)if(!seen.has(p.name))throw new Error(`Visible PDF parameter not found: ${p.name}`);
    return {verified:seen.size,pages:doc.numPages};
  } finally { await doc.destroy(); }
}
module.exports={verifyVisibleReportPdf,rgbClass,expectedClass,classifyIndicator};
