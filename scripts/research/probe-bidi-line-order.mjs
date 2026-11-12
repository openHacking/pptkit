import { chromium, firefox, webkit } from '@playwright/test';
import { createPresentation } from '../../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';
import { writeFile } from 'node:fs/promises';
const samples = ['ABC (שלום 123) XYZ', 'ABC שלום - 123 XYZ', 'ABC (שלום 123 XYZ) DEF', 'ABC שלום (123) XYZ', 'ABC שלום 123 - DEF', 'ABC مرحبا (123) DEF', 'ABC שלום (ABC 123) XYZ', 'שלום 123 - ABC XYZ', 'ABC שלום - XYZ 123', 'ABC ܫܠܡ - 123 XYZ', 'ABC ދިވެހި - 123 XYZ', 'ABC ߒߞߏ - 123 XYZ'];
const deck = createPresentation(), slide = deck.addSlide(), cases = [];
for (const [i, text] of samples.entries()) for (const width of [60,80,100,120,140,160,180,200]) {
  const id = `t${i}w${width}`; cases.push({ id, text, width });
  slide.addElement({ type:'text', id, box: { x:0,y:0,width,height:600 }, frame: { margin:0,wrap:true }, content: [{ runs:[{ text,style:{fontFamily:'Arial',fontSize:24} }] }] });
}
const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
const engine = { chromium, firefox, webkit }[process.argv[3] ?? 'chromium'];
if (!engine) throw new Error('Expected browser: chromium, firefox or webkit');
const browser=await engine.launch({headless:true});
try {
  const page=await browser.newPage();await page.setContent(svg);
  const rows=await page.evaluate(cases=>cases.map(item=>{
    const group=document.querySelector(`[data-pptkit-element-id="${item.id}"]`), lines=[]; let index=0, lastY;
    const nativeLines=[...group.querySelectorAll('text[data-pptkit-text-line]')];
    for (const container of nativeLines.length ? nativeLines : [group]) {
      if (nativeLines.length) lines.push([]);
      const walker=document.createTreeWalker(container,NodeFilter.SHOW_TEXT);let node;
      while ((node=walker.nextNode())) for (let j=0;j<node.length;j++) {
        const range=document.createRange();range.setStart(node,j);range.setEnd(node,j+1);const rect=range.getBoundingClientRect();
        const y=Math.round(rect.y*100)/100;
        if (!nativeLines.length && lastY !== y) {lines.push([]); lastY=y;}
        lines.at(-1).push({index:index++,char:node.data[j],x:rect.x,width:rect.width});
      }
    }
    return {...item, lines};
  }),cases);
  await writeFile(process.argv[2]??'.tmp/bidi-line-order.json',JSON.stringify({browser:browser.version(),rows},null,2));
} finally {await browser.close()}
