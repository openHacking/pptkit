import {mkdir,mkdtemp,writeFile} from 'node:fs/promises';
import path from 'node:path';
import {createPresentation,prepareLayoutRebind} from '../packages/core/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {layoutRebindDocument} from '../tests/fixtures/layout-rebind.mjs';
await mkdir('.tmp/placeholder-metrics',{recursive:true});const dir=path.resolve(await mkdtemp('.tmp/placeholder-metrics/run-'));
for(const family of ['Georgia','Times New Roman','Courier New']){
 const d=layoutRebindDocument(init=>createPresentation({...init,theme:{fonts:{heading:family,body:family}},textStylePresets:{local:{run:{fontFamily:'Georgia',fontSize:24,color:'CC2200',bold:false},paragraph:{align:'right'},frame:{wrap:false}}}}));
 for(let i=0;i<3;i++){const s=d.slides[i],e=s.elements[1];s.replaceElement(e.id,{...e,textStylePreset:'local',frame:{wrap:false},content:[{style:{align:'right'},runs:[{text:`Native reset page ${i}`,style:{fontFamily:'Arial',color:'CC2200',bold:false}}]}]});}
 for(const [slideId,options] of [['page-0',{localAppearance:'adopt',reset:['text.fontFamily','text.color','text.bold','text.paragraph.align','text.frame.wrap']}],['page-1',{localAppearance:'adopt'}]])d.applyLayoutRebind(prepareLayoutRebind(d,[{slideId,layoutId:'target'}],options));
 for(const mode of ['materialize','inherit']){const out=await generatePptx(d,{nativePlaceholderText:mode});await writeFile(path.join(dir,`${family.replaceAll(' ','-')}-${mode}.pptx`),out.bytes);}
}
console.log(dir);
