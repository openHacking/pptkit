import assert from 'node:assert/strict';
import test from 'node:test';
import {readFile} from 'node:fs/promises';
import {chromium} from '@playwright/test';
import {createPresentation,normalizePresentation,resolvePresentationInheritance} from '../packages/core/dist/index.js';
import {resolveLayout} from '../packages/layout/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {renderLayoutToSvg} from '../packages/svg-renderer/dist/index.js';
import {masterAdvancedDocument,editMasterHeading} from '../tests/fixtures/master-advanced.mjs';
import {readGeneratedZip} from '../tests/support/read-generated-zip.mjs';


async function snapshot(api,factory,edit){const d=factory(api.createPresentation);if(edit)edit(d);const layout=api.resolveLayout(d),out=await api.generatePptx(d,{connectorMode:'native'});return JSON.parse(JSON.stringify({normalized:api.normalizePresentation(d),origins:api.resolvePresentationInheritance(d).origins,layout,svg:await api.renderLayoutToSvg(layout),bytes:[...out.bytes],warnings:out.warnings}));}

test('advanced master owners and shared heading replacement agree in Node/ESM/global with independent native XML controls',async t=>{
  const api={createPresentation,normalizePresentation,resolvePresentationInheritance,resolveLayout,generatePptx,renderLayoutToSvg},before=await snapshot(api,masterAdvancedDocument),after=await snapshot(api,masterAdvancedDocument,editMasterHeading);
  const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});
  const imports={};for(const name of ['core','layout','pptx-exporter','svg-renderer'])imports['@pptkit/'+name]='data:text/javascript;base64,'+Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`,'utf8')).toString('base64');
  await page.setContent(`<script type="importmap">${JSON.stringify({imports})}</script>`);
  for(const format of ['esm','global']){
    if(format==='global')for(const name of ['core','layout','svg-renderer','pptx-exporter'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
    const actual=await page.evaluate(async({format,factory,snapshot,edit})=>{const api=format==='global'?PPTKit:Object.assign({},...await Promise.all(['core','layout','pptx-exporter','svg-renderer'].map(n=>import('@pptkit/'+n))));const f=Function(`return (${factory})`)(),s=Function(`return (${snapshot})`)(),e=Function(`return (${edit})`)();return{before:await s(api,f),after:await s(api,f,e)};},{format,factory:masterAdvancedDocument.toString(),snapshot:snapshot.toString(),edit:editMasterHeading.toString()});
    assert.deepEqual(actual,{before,after});
  }
  const oldParts=readGeneratedZip(Uint8Array.from(before.bytes)),newParts=readGeneratedZip(Uint8Array.from(after.bytes));assert.deepEqual([...newParts.keys()],[...oldParts.keys()]);assert.deepEqual([...oldParts.keys()].filter(k=>!oldParts.get(k).equals(newParts.get(k))),['ppt/slideMasters/slideMaster2.xml']);
  for(const i of [1,3])assert.equal(after.svg.slides[i].svg,before.svg.slides[i].svg);for(const i of [0,2]){assert.notEqual(after.svg.slides[i].svg,before.svg.slides[i].svg);assert.ok(after.svg.slides[i].svg.includes('MASTER red 2'));}
  const result=await page.evaluate(({parts})=>{
    const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',C='http://schemas.openxmlformats.org/drawingml/2006/chart',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',REL='http://schemas.openxmlformats.org/package/2006/relationships';
    const parse=name=>{const doc=new DOMParser().parseFromString(parts[name],'application/xml');if(doc.getElementsByTagName('parsererror').length)throw Error('Invalid generated XML');return doc;};
    const records=[];
    for(const[n,owner]of [[2,'red'],[3,'blue']]){
      const part=`ppt/slideMasters/slideMaster${n}.xml`,doc=parse(part),rels=parse(`ppt/slideMasters/_rels/slideMaster${n}.xml.rels`),table=doc.getElementsByTagNameNS(A,'tbl')[0],cells=[...table.getElementsByTagNameNS(A,'tc')],pr=cells[0].getElementsByTagNameNS(A,'tcPr')[0],charts=[...doc.getElementsByTagNameNS(C,'chart')],edges=[...rels.getElementsByTagNameNS(REL,'Relationship')];
      if(charts.length!==3||doc.getElementsByTagNameNS(A,'tbl').length!==1)throw Error('Wrong owner collection');
      const heading=[...doc.getElementsByTagNameNS(A,'t')].map(n=>n.textContent).filter(s=>s.startsWith('MASTER '));
      const types=charts.map(node=>{const rid=node.getAttributeNS(R,'id'),edge=edges.find(e=>e.getAttribute('Id')===rid),target=edge.getAttribute('Target').replace('../','ppt/'),chart=parse(target);return ['barChart','lineChart','pieChart'].find(kind=>chart.getElementsByTagNameNS(C,kind).length);});
      const connector=doc.getElementsByTagNameNS(P,'cxnSp')[0],start=connector.getElementsByTagNameNS(A,'stCxn')[0],end=connector.getElementsByTagNameNS(A,'endCxn')[0],ids=new Set([...doc.getElementsByTagNameNS(P,'cNvPr')].map(n=>n.getAttribute('id')));
      if(!ids.has(start.getAttribute('id'))||!ids.has(end.getAttribute('id')))throw Error('Dangling owner-local connector');
      records.push({owner,heading,types,margin:pr.getAttribute('marL'),children:[...pr.children].map(n=>n.localName),gridSpan:cells[0].getAttribute('gridSpan'),continuation:cells[1].getAttribute('hMerge')});
    }
    for(let i=1;i<=4;i++){const d=parse(`ppt/slides/slide${i}.xml`);if(d.getElementsByTagNameNS(A,'tbl').length||d.getElementsByTagNameNS(C,'chart').length||d.getElementsByTagNameNS(P,'cxnSp').length)throw Error('Flattened inherited object');}
    return records;
  },{parts:Object.fromEntries([...newParts].filter(([n])=>n.endsWith('.xml')||n.endsWith('.rels')).map(([n,b])=>[n,b.toString()]))});
  assert.deepEqual(result,['red','blue'].map(owner=>({owner,heading:['MASTER '+owner+(owner==='red'?' 2':'')],types:['barChart','lineChart','pieChart'],margin:'50800',children:['lnL','lnR','lnT','lnB','solidFill'],gridSpan:'2',continuation:'1'})));
  assert.deepEqual(before.warnings,after.warnings);assert.equal(before.warnings.length,2);assert.ok(before.warnings.every(w=>w.code==='table-style-materialized'&&['red','blue'].includes(w.masterId)));assert.deepEqual(requests,[]);
});
