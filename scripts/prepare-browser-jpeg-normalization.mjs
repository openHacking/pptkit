import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { createHash } from 'node:crypto';
import { chromium } from '@playwright/test';
const dir=process.argv[2];if(!dir || process.argv.length!==3)throw Error('Expected sample directory.');
const source=await readFile(join(dir,'orientation6.jpg'));
const bundle=await readFile('packages/pptx-exporter/dist/browser/dom.js');
const adapter='data:text/javascript;base64,'+bundle.toString('base64');
const browser=await chromium.launch();
try {
  const page=await browser.newPage(),requests=[];await page.route('**/*',route=>{requests.push(route.request().url());return route.abort()});
  const result=await page.evaluate(async({source,adapter})=>{
    const api=await import(adapter),output=await api.normalizeJpegOrientationInBrowser(Uint8Array.from(source));
    return {...output,bytes:[...output.bytes]};
  },{source:[...source],adapter});
  const bytes=Uint8Array.from(result.bytes);await writeFile(join(dir,'browser-normalized.png'),bytes);
  await writeFile(join(dir,'browser-normalized.json'),JSON.stringify({browserVersion:browser.version(),requests,sourceSha256:createHash('sha256').update(source).digest('hex'),outputSha256:createHash('sha256').update(bytes).digest('hex'),width:result.width,height:result.height,sourceOrientation:result.sourceOrientation,mimeType:result.mimeType,orientationNormalization:result.orientationNormalization},null,2)+'\n');
  console.log(JSON.stringify({width:result.width,height:result.height,sourceOrientation:result.sourceOrientation}));
}finally{await browser.close()}
