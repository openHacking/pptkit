import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';
import { mkdir,writeFile } from 'node:fs/promises';
const dir=process.argv[2]??'.tmp/bidi-line-native';await mkdir(dir,{recursive:true});const p=createPresentation(),s=p.addSlide();
for(const [i,width] of [80,120,160].entries())for(const [j,runs] of [[{text:'ABC שלום - 123 XYZ'}],[{text:'ABC שלום - '},{text:'123 XYZ'}]].entries())s.addElement({type:'text',id:`w${width}-${j}`,box:{x:50+i*210,y:40+j*220,width,height:200},frame:{margin:0,wrap:true},content:[{runs:runs.map(r=>({...r,style:{fontFamily:'Arial',fontSize:24}}))}]});
const [pptx,svg]=await Promise.all([generatePptx(p),renderPresentationToSvg(p)]);await writeFile(`${dir}/source.pptx`,pptx.bytes);await writeFile(`${dir}/preview.svg`,svg.slides[0].svg);console.log(JSON.stringify(svg.warnings));
