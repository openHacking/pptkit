import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { mkdir,writeFile } from 'node:fs/promises';
const dir=process.argv[2]??'.tmp/bidi-direction-native';const authored=process.argv[3]==='--authored';await mkdir(dir,{recursive:true});
const p=createPresentation(),slide=p.addSlide();
const samples=[['latin-first','ABC (שלום 123) XYZ'],['rtl-first','שלום (ABC 123) סוף']];
for(const [row,[id,text]] of samples.entries()) for(const [col,dir] of ['ltr','rtl'].entries())slide.addElement({type:'text',id:`${id}-${dir}`,box:{x:40+col*400,y:50+row*160,width:340,height:100},frame:{margin:0,wrap:false},content:[{style:{align:'left',...(authored?{direction:dir}:{})},runs:[{text,style:{fontFamily:'Arial',fontSize:32}}]}]});
await writeFile(`${dir}/base.pptx`,(await generatePptx(p)).bytes);
