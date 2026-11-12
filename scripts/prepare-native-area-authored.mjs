import { createHash } from 'node:crypto';
import { mkdir,mkdtemp,writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
const p=createPresentation({id:'area-authored',size:{width:960,height:540}});
for(const [group,blankDisplay] of [['standard','gap'],['standard','zero'],['stacked','zero'],['percentStacked','zero']]){
 const id=`${group}-${blankDisplay}`,slide=p.addSlide({id});slide.addElement({id:`title-${id}`,type:'text',box:{x:30,y:10,width:800,height:45},content:id});slide.addElement({id:`chart-${id}`,name:`chart-${id}`,type:'chart',chartType:'area',grouping:group,blankDisplay,box:{x:40,y:60,width:800,height:420},categories:['A','B','C','D','E'],series:[{name:'RED',values:[4,2,null,-2,3],color:'#FF0000'},{name:'BLUE',values:[2,-3,3,4,1],color:'#0000FF'}],legend:{visible:false},axes:{value:{scale:group==='percentStacked'?{min:-1,max:1,majorUnit:.5}:{min:-10,max:10,majorUnit:5}}}});
}
await mkdir('.tmp/native-area-authored',{recursive:true});const dir=path.resolve(await mkdtemp('.tmp/native-area-authored/powerpoint-'));const result=await generatePptx(p);await writeFile(path.join(dir,'area-authored.pptx'),result.bytes,{flag:'wx'});await writeFile(path.join(dir,'manifest.json'),JSON.stringify({sha256:createHash('sha256').update(result.bytes).digest('hex'),warnings:result.warnings},null,2)+'\n',{flag:'wx'});console.log(dir);
