import {mkdir,writeFile} from 'node:fs/promises';
import {createPresentation} from '../packages/core/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
const d=createPresentation({size:{width:600,height:360}}),s=d.addSlide({id:'s'});
s.addElement({id:'table',type:'table',name:'direction-table',box:{x:40,y:50,width:480,height:220},columns:[120,120,120,120],rows:[{height:220,cells:[undefined,'rotate90','rotate270','horizontal'].map((direction,i)=>({content:[{runs:[{text:['DEFAULT','CLOCKWISE','COUNTER','HORIZONTAL'][i],style:{fontFamily:'Arial',fontSize:18,color:'000000'}}]}],style:{fill:{type:'solid',color:['EEEEEE','F2E5D0','DCECF5','E0F1E0'][i]},margin:8,verticalAlign:'top',...(direction?{textDirection:direction}:{})}}))}]});
await mkdir('.tmp/native-table-text-direction',{recursive:true});
await writeFile('.tmp/native-table-text-direction/source.pptx',(await generatePptx(d)).bytes,{flag:'wx'});
