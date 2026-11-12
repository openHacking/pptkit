// Optional independent consumer experiment; requires soffice, pdftoppm and Pillow.
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import assert from 'node:assert/strict';
import { createPresentation } from '../../packages/core/dist/index.js';
import { planNestedVisibleImageReplacement } from '../../packages/layout/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';

const dir = await mkdtemp(join(tmpdir(), 'pptkit-nested-image-'));
function run(command, args) { const result=spawnSync(command,args,{encoding:'utf8',timeout:60000});assert.equal(result.status,0,`${command}: ${result.error?.message ?? result.stderr}`);return result.stdout.trim(); }
try {
  const paths=['red','blue','green'].map(name=>join(dir,`${name}.png`));
  run('python3',['-c',"from PIL import Image; import sys; colors=[(255,0,0),(0,0,255),(0,180,0)]; [Image.new('RGB',(8,8),c).save(p) for c,p in zip(colors,sys.argv[1:])]",...paths]);
  const [red,blue,green]=await Promise.all(paths.map(path=>readFile(path))),deck=createPresentation({size:{width:400,height:200}});
  for(const [id,bytes] of [['red',red],['blue',blue],['green',green]])deck.registerAsset({id,kind:'image',mimeType:'image/png',width:8,height:8,source:{type:'bytes',value:bytes}});
  const group=id=>({type:'group',id:`outer-${id}`,box:{x:50,y:30,width:120,height:120},coordinateSize:{width:100,height:100},transform:{rotation:180,flipH:true},children:[{type:'group',id:`inner-${id}`,box:{x:0,y:0,width:100,height:100},coordinateSize:{width:100,height:100},children:[{type:'image',id:`target-${id}`,assetId:'red',box:{x:10,y:10,width:30,height:30}},{type:'image',id:`sibling-${id}`,assetId:'blue',box:{x:60,y:10,width:30,height:30}}]}]});
  const first=deck.addSlide({id:'first',elements:[group('first')]}),second=deck.addSlide({id:'second',elements:[group('second')]});
  const rules=[{type:'constraints',container:{kind:'slide',id:'first'},elements:[{elementId:'outer-first',x:{value:200},width:{value:160}}]}];
  const plan=planNestedVisibleImageReplacement(deck,'first','target-first','green','stretch',{composition:rules,commitAncestorGeometry:true});
  assert.deepEqual(plan.committedAncestorIds,['outer-first']);first.replaceElement(plan.rootElementId,plan.replacement);assert.equal(second.elements[0].children[0].children[0].assetId,'red');
  const result=await generatePptx(deck);assert.deepEqual(result.warnings,[]);let source=join(dir,'input.pptx');await writeFile(source,result.bytes);
  const rounds=[];
  for(let round=0;round<=3;round++){
    const out=join(dir,`round-${round}`);await mkdir(out);
    if(round){run('soffice',['--headless','--convert-to','pptx','--outdir',out,source]);source=join(out,'input.pptx');}
    const inventory=JSON.parse(run('python3',['-c',`import hashlib,io,json,posixpath,sys,zipfile,xml.etree.ElementTree as E
from PIL import Image
out=[]
with zipfile.ZipFile(sys.argv[1]) as z:
 for n in [1,2]:
  slide=E.fromstring(z.read(f'ppt/slides/slide{n}.xml'));rels=E.fromstring(z.read(f'ppt/slides/_rels/slide{n}.xml.rels'))
  targets={r.get('Id'):posixpath.normpath(posixpath.join('ppt/slides',r.get('Target'))) for r in rels}; ns='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
  items=[]
  for blip in slide.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
   target=targets[blip.get(ns+'embed')];data=z.read(target);im=Image.open(io.BytesIO(data)).convert('RGB');items.append({'target':target,'sha256':hashlib.sha256(data).hexdigest(),'center':list(im.getpixel((im.width//2,im.height//2)))})
  grp=[]
  for node in slide.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}grpSp'):
   xfrm=node.find('{http://schemas.openxmlformats.org/presentationml/2006/main}grpSpPr/{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
   if xfrm is not None:
    grp.append({'attributes':dict(xfrm.attrib),'children':{e.tag.split('}')[-1]:dict(e.attrib) for e in xfrm}})
  out.append({'images':items,'groups':grp})
print(json.dumps(out))`,source]));
    run('soffice',['--headless','--convert-to','pdf','--outdir',out,source]);run('pdftoppm',['-png','-r','72',join(out,'input.pdf'),join(out,'page')]);
    const pixels=JSON.parse(run('python3',['-c',`from PIL import Image
import json,sys
images=[Image.open(p).convert('RGB') for p in sys.argv[1:]]
points=[[(240,120),(320,120),(80,120)],[(80,120),(140,120),(240,120)]]
print(json.dumps([[list(im.getpixel(p)) for p in ps] for im,ps in zip(images,points)]))`,join(out,'page-1.png'),join(out,'page-2.png')]));
    assert.deepEqual(pixels,[[[0,180,0],[0,0,255],[255,255,255]],[[255,0,0],[0,0,255],[255,255,255]]]);
    assert.deepEqual(inventory.map(page=>page.images.map(item=>item.center)),[[[0,180,0],[0,0,255]],[[255,0,0],[0,0,255]]]);
    assert.equal(inventory[0].images[1].sha256,inventory[1].images[1].sha256);
    assert.equal(inventory[0].images[1].target,inventory[1].images[1].target);
    const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
    assert.deepEqual(inventory.map(page=>page.images.map(item=>item.sha256)),[[sha(green),sha(blue)],[sha(red),sha(blue)]]);
    assert.deepEqual(inventory.map(page=>page.groups.length),[2,2]);
    if(round===0){assert.deepEqual(inventory[0].groups[0].attributes,{rot:'10800000',flipH:'1'});assert.equal(inventory[0].groups[0].children.off.x,'2540000');assert.equal(inventory[0].groups[0].children.ext.cx,'2032000');}
    const bytes=await readFile(source),session=await importPptx(bytes,{projection:'content'});
    assert.deepEqual((await session.save((await session.prepareSave({mode:'no-op'})).plan)).bytes,new Uint8Array(bytes));session.dispose();
    rounds.push({round,inventory,pixels,parserNoOpExact:true});
  }
  console.log(JSON.stringify({version:run('soffice',['--version']),rounds},null,2));
}finally{await rm(dir,{recursive:true,force:true});}
