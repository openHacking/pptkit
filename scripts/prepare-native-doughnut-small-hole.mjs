import { createHash } from "node:crypto";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
await mkdir('.tmp/native-doughnut-small-hole',{recursive:true});const dir=path.resolve(await mkdtemp('.tmp/native-doughnut-small-hole/powerpoint-'));
const p=createPresentation({id:'small-doughnut-holes',size:{width:960,height:540}});
for(const holeSize of [1,5,9,10]){const slide=p.addSlide({id:`hole-${holeSize}`});slide.addElement({id:`title-${holeSize}`,type:'text',box:{x:30,y:10,width:800,height:45},content:`HOLE ${holeSize}%`});slide.addElement({id:`chart-${holeSize}`,name:`chart-${holeSize}`,type:'chart',chartType:'doughnut',box:{x:80,y:70,width:700,height:420},holeSize,firstSliceAngle:0,categories:['A','B'],series:[{name:'S',values:[1,3],pointColors:['#FF0000','#0000FF']}],legend:{visible:false}});}
const output=await generatePptx(p);await writeFile(path.join(dir,'small-holes.pptx'),output.bytes,{flag:'wx'});await writeFile(path.join(dir,'manifest.json'),JSON.stringify({sha256:createHash('sha256').update(output.bytes).digest('hex'),warnings:output.warnings},null,2)+'\n',{flag:'wx'});console.log(dir);
