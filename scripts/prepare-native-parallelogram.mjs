import { mkdir, writeFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
const deck=createPresentation({size:{width:600,height:360}}),slide=deck.addSlide({id:'s'});
for(const [id,box,transform] of [['wide',{x:30,y:30,width:220,height:80},{}],['tall',{x:310,y:30,width:80,height:220},{}],['rotated',{x:40,y:190,width:220,height:80},{rotation:30,flipH:true}]]) slide.addElement({id,name:id,type:'shape',shape:'parallelogram',box,transform,style:{fill:{type:'solid',color:'FF0000'},stroke:{paint:{type:'none'},width:0}}});
await mkdir('.tmp/native-parallelogram',{recursive:true});
await writeFile('.tmp/native-parallelogram/source.pptx',(await generatePptx(deck)).bytes,{flag:'wx'});
