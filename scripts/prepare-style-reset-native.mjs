import {mkdir,mkdtemp,writeFile} from 'node:fs/promises';
import path from 'node:path';
import {createPresentation,normalizePresentation,prepareLayoutRebind,resolvePresentationInheritance} from '../packages/core/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {layoutRebindDocument} from '../tests/fixtures/layout-rebind.mjs';
await mkdir('.tmp/style-reset-native',{recursive:true});const directory=path.resolve(await mkdtemp('.tmp/style-reset-native/run-'));
const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{run:{fontFamily:'Georgia',fontSize:24,color:'CC2200',bold:false},paragraph:{align:'right'},frame:{wrap:false}}}}));
for(let i=0;i<3;i++){const s=d.slides[i],body=s.elements[1];s.replaceElement(body.id,{...body,textStylePreset:'local',frame:{wrap:false},content:[{style:{align:'right'},runs:[{text:`Native reset page ${i}`,style:{fontFamily:'Arial',color:'CC2200',bold:false}}]}]});}
const before=await generatePptx(d);const plans=[];for(const [slideId,options] of [['page-0',{localAppearance:'adopt',reset:['text.fontFamily','text.color','text.bold','text.paragraph.align','text.frame.wrap']}],['page-1',{localAppearance:'adopt'}]]){const plan=prepareLayoutRebind(d,[{slideId,layoutId:'target'}],options);if(plan.status!=='ready')throw Error('control needs review');d.applyLayoutRebind(plan);plans.push(plan);}const after=await generatePptx(d,{nativePlaceholderText:"inherit"}),materialized=await generatePptx(d);
for(const [name,value] of Object.entries({'before.pptx':before.bytes,'input.pptx':after.bytes,'default-materialized.pptx':materialized.bytes,'plans.json':plans,'normalized.json':normalizePresentation(d),'origins.json':resolvePresentationInheritance(d).origins,'warnings.json':after.warnings}))await writeFile(path.join(directory,name),value instanceof Uint8Array?value:JSON.stringify(value,null,2)+'\n');console.log(directory);
