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
await mkdir('.tmp/native-scatter-authored',{recursive:true});const dir=path.resolve(await mkdtemp('.tmp/native-scatter-authored/powerpoint-')),result=await generatePptx(p);await writeFile(path.join(dir,'scatter-authored.pptx'),result.bytes,{flag:'wx'});await writeFile(path.join(dir,'manifest.json'),JSON.stringify({sha256:createHash('sha256').update(result.bytes).digest('hex'),warnings:result.warnings},null,2)+'\n',{flag:'wx'});console.log(dir);
