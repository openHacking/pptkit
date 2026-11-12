import {writeFile} from 'node:fs/promises';
import {createPresentation} from '../packages/core/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {renderPresentationToSvg} from '../packages/svg-renderer/dist/index.js';
import {chromium} from '@playwright/test';
const dir=process.argv[2];if(!dir)throw Error('Expected evidence directory');
const d=createPresentation({size:{width:960,height:540}});
for(const [i,pattern] of ['horizontal','vertical','cross'].entries())d.addSlide({id:`page-${i}`}).addElement({id:`shape-${i}`,type:'shape',shape:'rect',box:{x:180,y:120,width:600,height:300},style:{fill:{type:'pattern',pattern,foreground:'FF0000',background:'0000FF',foregroundOpacity:.5,backgroundOpacity:1},stroke:{paint:{type:'none'}}}});
await writeFile(`${dir}/pattern.pptx`,(await generatePptx(d)).bytes,{flag:'wx'});const svgs=(await renderPresentationToSvg(d)).slides,browser=await chromium.launch(),page=await browser.newPage({viewport:{width:960,height:540}});for(const [i,slide] of svgs.entries()){await page.setContent(`<style>body{margin:0}</style>${slide.svg}`);await page.screenshot({path:`${dir}/svg-${i+1}.png`});}await browser.close();
