import { readFile, writeFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';

const dir=process.argv[2];if(!dir)throw Error('Expected existing output directory');
const data=await readFile('tests/fixtures/two-frame-apng.png'),d=createPresentation({size:{width:960,height:540}});
d.registerAsset({id:'animated',kind:'image',mimeType:'image/apng',source:{type:'bytes',value:data},width:3,height:2});d.addSlide({id:'slide'}).addElement({id:'image',type:'image',assetId:'animated',fit:'contain',box:{x:180,y:90,width:600,height:360}});
const pptx=await generatePptx(d,{pngDiagnostics:'warn'}),svg=await renderPresentationToSvg(d);await writeFile(`${dir}/source.pptx`,pptx.bytes,{flag:'wx'});await writeFile(`${dir}/preview.svg`,svg.slides[0].svg,{flag:'wx'});await writeFile(`${dir}/warnings.json`,JSON.stringify({pptx:pptx.warnings,svg:svg.warnings},null,2),{flag:'wx'});
