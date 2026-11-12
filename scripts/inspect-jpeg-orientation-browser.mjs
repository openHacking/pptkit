import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { chromium } from '@playwright/test';
const dir=process.argv[2]; if(!dir || process.argv.length!==3) throw Error('Expected sample directory.');
const svg=await readFile(join(dir,'preview.svg'),'utf8'),browser=await chromium.launch();
try {const page=await browser.newPage({viewport:{width:960,height:540},deviceScaleFactor:1});let requests=[];await page.route('**/*',route=>{requests.push(route.request().url());return route.abort()});await page.setContent(`<style>html,body{margin:0}</style>${svg}`);await page.evaluate(async()=>Promise.all([...document.images].map(i=>i.decode())));await page.screenshot({path:join(dir,'browser.png')});await writeFile(join(dir,'browser.json'),JSON.stringify({version:browser.version(),requests},null,2)+'\n');} finally{await browser.close();}
