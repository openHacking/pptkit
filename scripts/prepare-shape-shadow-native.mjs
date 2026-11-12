import { writeFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';

const dir=process.argv[2];if(!dir)throw Error('Expected existing evidence directory');
const d=createPresentation({size:{width:960,height:540}});
for(const [i,spec] of [
 {opacity:1,shadow:{color:'000000',opacity:.5,blur:0,distance:60,direction:0}},
 {opacity:.5,shadow:{color:'000000',opacity:.5,blur:0,distance:60,direction:0}},
 {opacity:1,shadow:{color:'000000',opacity:.5,blur:18,distance:60,direction:90}},
].entries())d.addSlide({id:`page-${i}`}).addElement({id:`shape-${i}`,type:'shape',shape:'rect',box:{x:300,y:150,width:240,height:180},opacity:spec.opacity,style:{fill:{type:'solid',color:'FF0000'},stroke:{paint:{type:'none'}},shadow:spec.shadow}});
await writeFile(`${dir}/shadow.pptx`,(await generatePptx(d)).bytes,{flag:'wx'});
const result=await renderPresentationToSvg(d),browser=await chromium.launch();try{const page=await browser.newPage({viewport:{width:960,height:540}});for(const [i,slide] of result.slides.entries()){await page.setContent(`<style>body{margin:0}</style>${slide.svg}`);await page.screenshot({path:`${dir}/svg-${i+1}.png`});}}finally{await browser.close();}
