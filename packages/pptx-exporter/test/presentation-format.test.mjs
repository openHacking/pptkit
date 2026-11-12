import assert from 'node:assert/strict';
import test from 'node:test';
import {readFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {createPresentation} from '@pptkit/core';
import {generatePptx} from '../dist/index.js';
import {appPropertiesXml} from '../dist/ooxml/package-parts.js';
import {readGeneratedZip} from '../../../tests/support/read-generated-zip.mjs';
import {textOriginsDocument} from '../../../tests/fixtures/text-origins.mjs';
import {masterArtworkDocument} from '../../../tests/fixtures/master-artwork.mjs';
const metadata={title:'Format',author:'PPTKit',language:'en-US',keywords:[],revision:1};
const label=(width,height)=>appPropertiesXml(1,metadata,{width,height,unit:'pt'}).match(/<PresentationFormat>([^<]+)<\/PresentationFormat>/)[1];

test('format metadata follows exact serialized screen ratios while portrait/square/paper/custom stay custom',()=>{
  for(const[w,h,expected]of [[960,540,'16:9'],[720,405,'16:9'],[720,540,'4:3'],[800,600,'4:3'],[960,600,'16:10'],[320,200,'16:10']])assert.equal(label(w,h),`On-screen Show (${expected})`);
  for(const[w,h]of [[540,960],[540,720],[600,600],[500,400],[842,595],[595,842],[123.45,67.89]])assert.equal(label(w,h),'Custom');
});

test('classification uses EMU rounding without float epsilon matching or unsafe cross multiplication',()=>{
  assert.equal(label(960+0.4/12700,540),'On-screen Show (16:9)');
  assert.equal(label(960+0.6/12700,540),'Custom');
  assert.equal(label(720-0.4/12700,540),'On-screen Show (4:3)');
  assert.equal(label(720-0.6/12700,540),'Custom');
  assert.equal(label(1e-20,1e-20),'Custom');assert.equal(label(Number.MAX_VALUE,1),'Custom');
  // Both serialized dimensions are exact safe integers; their products by 16 are not.
  const x=9007199254740976,y=5066549580791799;
  assert.ok(Number.isSafeInteger(x)&&Number.isSafeInteger(y));
  assert.equal(Math.round(x/12700*12700),x);assert.equal(Math.round(y/12700*12700),y);
  assert.equal(label(x/12700,y/12700),'On-screen Show (16:9)');
  assert.equal(label((x+1)/12700,y/12700),'Custom');
});

test('fresh exports transport the actual normalized size without altering dimensions, counts or escaped properties',async()=>{
  for(const[width,height,expected]of [[720,540,'On-screen Show (4:3)'],[960,600,'On-screen Show (16:10)'],[500,500,'Custom'],[540,960,'Custom']]){
    const d=createPresentation({size:{width,height},metadata:{company:'A & <B>'}});d.addSlide();d.addSlide({hidden:true});
    const parts=readGeneratedZip((await generatePptx(d)).bytes),xml=parts.get('docProps/app.xml').toString(),presentation=parts.get('ppt/presentation.xml').toString();
    assert.ok(xml.includes(`<PresentationFormat>${expected}</PresentationFormat>`));assert.match(xml,/<Company>A &amp; &lt;B&gt;<\/Company>/);assert.match(xml,/<Slides>2<\/Slides>/);
    assert.ok(presentation.includes(`<p:sldSz cx="${Math.round(width*12700)}" cy="${Math.round(height*12700)}"/>`));
  }
});

test('fresh export enforces PresentationML slide bounds after EMU rounding',async()=>{
  for(const width of [72,4032,72-0.4/12700,4032+0.4/12700]){
    const d=createPresentation({size:{width,height:72}});d.addSlide();
    const parts=readGeneratedZip((await generatePptx(d)).bytes), xml=parts.get('ppt/presentation.xml').toString();
    assert.ok(xml.includes(`<p:sldSz cx="${Math.round(width*12700)}" cy="914400"/>`));
    const master=parts.get('ppt/notesMasters/notesMaster1.xml').toString();
    assert.doesNotMatch(master,/NaN|Infinity/);assert.match(master,/<a:prstGeom prst="rect">/);
  }
  for(const value of [72-0.6/12700,4032+0.6/12700,1e-300,Number.MAX_VALUE])for(const dimension of ['width','height']){
    const d=createPresentation({size:{width:720,height:540,[dimension]:value}});d.addSlide();
    await assert.rejects(generatePptx(d),e=>e instanceof RangeError && /914400–51206400 EMU/.test(e.message));
  }
  const d=createPresentation({size:{width:1,height:540}});d.registerAsset({id:'asset',kind:'image',source:{type:'path',value:'not-read.png'}});d.addSlide({elements:[{type:'image',assetId:'asset',box:{x:0,y:0,width:10,height:10}}]});
  let calls=0;await assert.rejects(generatePptx(d,{resolveAsset:()=>{calls++;throw Error('unexpected asset load');}}),RangeError);assert.equal(calls,0);
});

test('the admitted default 16:9 source stays byte-exact and the 4:3 artwork correction changes only one metadata string',async()=>{
  const legacy=textOriginsDocument(createPresentation),group=legacy.slides[0].elements.find(e=>e.id==='group'),{placeholderKey:_binding,...ordinary}=group;legacy.slides[0].replaceElement(group.id,ordinary);
  const hash=bytes=>createHash('sha256').update(bytes).digest('hex');
  assert.equal(hash((await generatePptx(legacy)).bytes),'ef3d886dcf09d80e23b58cbbafb67cbb26be84fbb2a5c71797c2eb0a3219cd2d');
  const before=JSON.parse(await readFile(new URL('../../../tests/fixtures/master-artwork-format-baseline.json',import.meta.url),'utf8')),after=readGeneratedZip((await generatePptx(masterArtworkDocument(createPresentation))).bytes);
  assert.deepEqual([...after.keys()].sort(),before.parts.map(p=>p.name).sort());
  for(const part of before.parts) if(part.name!=='docProps/app.xml') {assert.equal(hash(after.get(part.name)),part.sha256,part.name);assert.equal(after.get(part.name).length,part.bytes,part.name);}
  assert.equal(after.get('docProps/app.xml').toString(),before.appProperties.replace('<PresentationFormat>On-screen Show (16:9)</PresentationFormat>','<PresentationFormat>On-screen Show (4:3)</PresentationFormat>'));
});
