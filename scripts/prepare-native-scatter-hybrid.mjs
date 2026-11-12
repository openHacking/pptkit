// Controlled hybrid: independent python-pptx chart/axes with PPTKit XY series and workbook.
import { readFile, writeFile, mkdtemp, mkdir } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
import { zip } from '../packages/pptx-parser/test/zip-fixture.mjs';
import { scatterWorkbook } from '../packages/pptx-exporter/dist/ooxml/scatter-workbook.js';
import { scatterSeriesXml } from '../packages/pptx-exporter/dist/ooxml/scatter-series.js';

const source='.tmp/native-scatter-missing-x/powerpoint-tp3m_1ti/scatter-missing-x.pptx', bytes=await readFile(source),parts=readGeneratedZip(bytes);
const paint=color=>({type:'solid',color,opacity:1}),stroke=(color,width=2)=>({paint:paint(color),width,dash:'solid',beginArrow:'none',endArrow:'none'}),noStroke={paint:{type:'none'},width:1,dash:'solid',beginArrow:'none',endArrow:'none'};
const data=[{name:'RED',xValues:[-8,null,2,8],yValues:[4,3,-4,6],color:'#FF0000'},{name:'BLUE',xValues:[7,-1,3,-7],yValues:[-6,4,2,-2],color:'#0000FF'}];
const workbook=scatterWorkbook(data);
for(let i=1;i<=4;i++){
 const chart=`ppt/charts/chart${i}.xml`,rel=`ppt/charts/_rels/chart${i}.xml.rels`,xml=parts.get(chart).toString(),rels=parts.get(rel).toString();
 const match=rels.match(/<Relationship\b[^>]*Type="[^"]*\/package"[^>]*Target="([^"]+)"/);if(!match)throw new Error(`Missing workbook relationship in ${rel}`);
 const workbookPart=path.posix.normalize(path.posix.join('ppt/charts',match[1]));if(!parts.has(workbookPart))throw new Error(`Missing ${workbookPart}`);
 const series=data.map(d=>({...d,line:i===1?false:stroke(d.color),marker:{shape:'circle',size:10,fill:paint(d.color),stroke:noStroke}}));
 const before=xml.match(/<c:ser>.*?<\/c:ser><c:ser>.*?<\/c:ser>/s);if(!before)throw new Error(`Expected exactly adjacent chart series in ${chart}`);
 parts.set(chart,Buffer.from(xml.replace(before[0],scatterSeriesXml(series))));parts.set(workbookPart,Buffer.from(workbook));
}
const parent='.tmp/native-scatter-hybrid';await mkdir(parent,{recursive:true});const dir=await mkdtemp(path.join(parent,'powerpoint-')),result=zip([...parts].map(([name,body])=>({name,body}))).bytes;
const file=path.join(dir,'scatter-hybrid.pptx');await writeFile(file,result);await writeFile(path.join(dir,'manifest.json'),JSON.stringify({scope:'independent chart/axes with PPTKit scatter series and workbook substituted',source,sourceSha256:createHash('sha256').update(bytes).digest('hex'),sha256:createHash('sha256').update(result).digest('hex')},null,2)+'\n');
console.log(JSON.stringify({file,dir}));
