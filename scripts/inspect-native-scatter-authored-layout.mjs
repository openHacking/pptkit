import { createHash } from 'node:crypto';
import { mkdir,mkdtemp,writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
const p=createPresentation({id:'scatter-authored',size:{width:960,height:540}}),data=[{name:'RED',xValues:[-8,null,2,8],yValues:[4,3,-4,6],color:'#FF0000'},{name:'BLUE',xValues:[7,-1,3,-7],yValues:[-6,4,2,-2],color:'#0000FF'}];
for(const [name,mode,lines] of [['markers-gap','gap',false],['line-gap','gap',true],['line-zero','zero',true],['line-span','span',true]]){
 const slide=p.addSlide({id:name});slide.addElement({id:`title-${name}`,type:'text',box:{x:30,y:10,width:800,height:45},content:name});
 slide.addElement({id:`chart-${name}`,name:`chart-${name}`,type:'chart',chartType:'scatter',blankDisplay:mode,box:{x:40,y:60,width:800,height:420},legend:{visible:false},axes:{x:{scale:{min:-10,max:10,majorUnit:5}},y:{scale:{min:-10,max:10,majorUnit:5}}},series:data.map(d=>({...d,line:lines?{paint:{type:'solid',color:d.color},width:2}:false,marker:{shape:'circle',size:10,fill:{type:'solid',color:d.color},stroke:{paint:{type:'none'},width:1}}}))});
}
const {resolveLayout}=await import('../packages/layout/dist/index.js');
const {readFile}=await import('node:fs/promises');
const bytes=await readFile('.tmp/native-scatter-authored/powerpoint-PmQkqF/scatter-authored.pptx');
const slides=resolveLayout(p).slides.map(s=>{const e=s.elements.find(e=>e.type==='chart');return {box:e.box,plotBox:e.chartLayout.plotBox,xScale:e.chartLayout.xScale,yScale:e.chartLayout.valueScale,scatter:e.chartLayout.scatter};});
await writeFile('.tmp/native-scatter-authored/powerpoint-PmQkqF/layout.json',JSON.stringify({sourceSha256:createHash('sha256').update(bytes).digest('hex'),slides},null,2)+'\n');
