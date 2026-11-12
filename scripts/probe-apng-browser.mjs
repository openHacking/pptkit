import { readFile, writeFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
const dir=process.argv[2];if(!dir)throw Error('Expected existing output directory');
const svg=await readFile(`${dir}/preview.svg`,'utf8'),apng=await readFile('tests/fixtures/two-frame-apng.png'),browser=await chromium.launch();
try{
 for(const mode of ['inline','svg-image','direct-image']){
  const page=await browser.newPage({viewport:{width:960,height:540},deviceScaleFactor:1});
  const body=mode==='inline'?svg:mode==='svg-image'?`<img width="960" height="540" src="data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}">`:`<img style="position:absolute;left:210px;top:90px;width:540px;height:360px" src="data:image/png;base64,${apng.toString('base64')}">`;
  await page.setContent(`<style>body{margin:0;background:white}img{image-rendering:pixelated}</style>${body}`);const started=Date.now(),samples=[];
  for(let i=0;i<10;i++){await page.waitForTimeout(60);const captured=Date.now()-started;await page.screenshot({path:`${dir}/${mode}-${i}.png`});samples.push(captured)}
  await writeFile(`${dir}/${mode}-times.json`,JSON.stringify(samples),{flag:'wx'});await page.close();
 }
}finally{await browser.close();}
