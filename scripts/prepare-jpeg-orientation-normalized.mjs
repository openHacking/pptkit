import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { createHash } from 'node:crypto';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
const dir=process.argv[2];if(!dir || process.argv.length!==3) throw Error('Expected sample directory.');
const jpeg=await readFile(join(dir,'orientation6.jpg')),png=await readFile(join(dir,'normalized.png'));
const deck=createPresentation({size:{width:960,height:540}});
deck.registerAsset({id:'exif',kind:'image',source:{type:'bytes',value:jpeg},mimeType:'image/jpeg',width:300,height:200});
deck.registerAsset({id:'normalized',kind:'image',source:{type:'bytes',value:png},mimeType:'image/png',width:200,height:300});
const slide=deck.addSlide({id:'orientation'});slide.addElement({type:'text',id:'title',content:'EXIF 6 / normalized control',box:{x:40,y:25,width:850,height:60}});
for(const [id,x] of [['exif',100],['normalized',500]]) slide.addElement({type:'image',id:`image-${id}`,assetId:id,fit:'stretch',box:{x,y:120,width:200,height:300}});
const output=await generatePptx(deck,{jpegOrientationDiagnostics:'warn',resolveAsset:asset=>asset.id==='exif'
  ? {bytes:png,mimeType:'image/png',orientationNormalization:'pixels'} : {bytes:png,mimeType:'image/png'}});
if(output.warnings.length!==1 || output.warnings[0].code!=='asset-orientation-normalized')throw Error('Unexpected warning.');
const hash=b=>createHash('sha256').update(b).digest('hex');await writeFile(join(dir,'normalized-candidate.pptx'),output.bytes);
await writeFile(join(dir,'normalized-producer.json'),JSON.stringify({pptxSha256:hash(output.bytes),sourceJpegSha256:hash(jpeg),normalizedPngSha256:hash(png),warnings:output.warnings},null,2)+'\n');console.log(JSON.stringify({sha256:hash(output.bytes)}));
