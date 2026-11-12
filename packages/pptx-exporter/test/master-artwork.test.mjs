import assert from 'node:assert/strict';
import test from 'node:test';
import {createPresentation} from '@pptkit/core';
import {generatePptx} from '../dist/index.js';
import {readGeneratedZip} from '../../../tests/support/read-generated-zip.mjs';

test('master charts/tables/groups keep their owning part and complete chart/workbook relationships',async()=>{
  const d=createPresentation();d.defineSlideMaster({id:'m',name:'M',elements:[{id:'group',type:'group',box:{x:0,y:0,width:500,height:200},coordinateSize:{width:500,height:200},children:[
    {id:'chart',type:'chart',chartType:'line',box:{x:0,y:0,width:200,height:150},categories:['A','B'],series:[{name:'Series',values:[1,2]}]},
    {id:'table',type:'table',box:{x:210,y:0,width:200,height:50},columns:[200],rows:[{cells:[{content:'Shared cell'}]}],style:{wholeTable:{text:{run:{bold:true}}}}},
  ]}]});d.defineSlideLayout({id:'l',name:'L',masterId:'m'});d.addSlide({layoutId:'l'});d.addSlide({layoutId:'l'});
  const out=await generatePptx(d),parts=readGeneratedZip(out.bytes),master=parts.get('ppt/slideMasters/slideMaster2.xml').toString(),rels=parts.get('ppt/slideMasters/_rels/slideMaster2.xml.rels').toString();
  assert.match(master,/<c:chart/);assert.match(master,/<a:tbl>/);assert.match(master,/Shared cell/);assert.match(rels,/charts\/chart1.xml/);
  assert.ok(parts.has('ppt/charts/chart1.xml'));assert.ok(parts.has('ppt/charts/_rels/chart1.xml.rels'));assert.ok(parts.has('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'));
  assert.match(parts.get('[Content_Types].xml').toString(),/\/ppt\/charts\/chart1.xml/);
  for(const i of [1,2]){const xml=parts.get(`ppt/slides/slide${i}.xml`).toString();assert.doesNotMatch(xml,/Shared cell|<c:chart|<a:tbl>/);}
  assert.equal(out.warnings.find(w=>w.code==='table-style-materialized').masterId,'m');
});

test('failed master image loading is explicit and located without crossing owner relationship scopes',async()=>{
  const d=createPresentation();d.registerAsset({id:'bad',kind:'image',source:{type:'bytes',encoding:'base64',value:'AA=='},mimeType:'image/png'});
  d.defineSlideMaster({id:'m',name:'M',elements:[{id:'image',type:'image',assetId:'bad',box:{x:0,y:0,width:20,height:20}}]});d.defineSlideLayout({id:'l',name:'L',masterId:'m'});d.addSlide({layoutId:'l'});
  const out=await generatePptx(d),parts=readGeneratedZip(out.bytes);assert.ok(out.warnings.some(w=>w.code==='asset-read-failed'&&w.assetId==='bad'));
  assert.ok(out.warnings.some(w=>w.code==='image-omitted'&&w.masterId==='m'&&w.assetId==='bad'));assert.doesNotMatch(parts.get('ppt/slideMasters/_rels/slideMaster2.xml.rels').toString(),/\/image"/);
  assert.doesNotMatch(parts.get('ppt/slideMasters/slideMaster2.xml').toString(),/<p:pic>/);
});
