import { createHash } from "node:crypto";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
const p=createPresentation({id:'area-research',size:{width:960,height:540}});
for(const [group,blankDisplay] of [['standard','gap'],['standard','zero'],['stacked','zero'],['percentStacked','zero']]) {
 const id=`${group}-${blankDisplay}`, slide=p.addSlide({id});
 slide.addElement({id:`title-${id}`,type:'text',box:{x:30,y:10,width:800,height:45},content:id});
 slide.addElement({id:`chart-${id}`,name:`chart-${id}`,type:'chart',chartType:'bar',grouping:group==='standard'?'clustered':group,blankDisplay,box:{x:40,y:60,width:800,height:420},categories:['A','B','C','D','E'],series:[{name:'RED',values:[4,2,null,-2,3],color:'#FF0000'},{name:'BLUE',values:[2,-3,3,4,1],color:'#0000FF'}],legend:{visible:false},axes:{value:{scale:group==='percentStacked'?{min:-100,max:100,majorUnit:50}:{min:-10,max:10,majorUnit:5}}}});
}
const original=(await generatePptx(p)).bytes,parts=readGeneratedZip(original);
for(let i=1;i<=4;i++){const key=`ppt/charts/chart${i}.xml`;let xml=parts.get(key).toString();xml=xml.replaceAll('barChart','areaChart').replace(/<c:barDir val="(?:bar|col)"\/>/,'').replace('<c:grouping val="clustered"/>','<c:grouping val="standard"/>').replace(/<c:gapWidth val="[^"]+"\/>/,'').replace(/<c:overlap val="[^"]+"\/>/,'').replace('<c:crossBetween val="between"/>','<c:crossBetween val="midCat"/>');parts.set(key,Buffer.from(xml));}
const bytes=zip([...parts].map(([name,body])=>({name,body}))).bytes;await mkdir('.tmp/native-area-research',{recursive:true});const dir=path.resolve(await mkdtemp('.tmp/native-area-research/powerpoint-'));await writeFile(path.join(dir,'area-research.pptx'),bytes,{flag:'wx'});await writeFile(path.join(dir,'manifest.json'),JSON.stringify({sha256:createHash('sha256').update(bytes).digest('hex'),sourceSha256:createHash('sha256').update(original).digest('hex'),groups:['standard-gap','standard-zero','stacked-zero','percentStacked-zero']},null,2)+'\n',{flag:'wx'});console.log(dir);
