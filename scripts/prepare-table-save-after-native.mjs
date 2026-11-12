import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir, mkdtemp } from 'node:fs/promises';
import path from 'node:path';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
const input=process.argv[2];if(!input)throw Error('Usage: node scripts/prepare-table-save-after-native.mjs native.pptx');
await mkdir('.tmp/table-after-native',{recursive:true});const directory=await mkdtemp(path.resolve('.tmp/table-after-native/run-'));let source=await readFile(input);await writeFile(path.join(directory,'source.pptx'),source);
for(let round=1;round<=3;round++){
 const session=await importPptx(source,{projection:'content',editSaveProfile:'text-runs-v3'});
 try {
  const slide=session.document.slides[0],previous=round===1?'Anchor 123':`Anchor after native ${round-1}`,previousRight=round===1?'Cell 456':`Cell after native ${round-1}`;
  const table=slide.elements.find(e=>e.type==='table'&&Array.isArray(e.rows[0]?.cells[0]?.content)&&e.rows[0].cells[0].content[0]?.runs[0]?.text===previous);assert.ok(table);
  const modified=structuredClone(table),anchor=modified.rows[0].cells[0],right=modified.rows[1].cells[0];assert.equal(anchor.rowSpan,2);assert.equal(anchor.colSpan,2);assert.equal(right.content[0].runs[0].text,previousRight);
  anchor.content[0].runs[0].text=`Anchor after native ${round}`;right.content[0].runs[0].text=`Cell after native ${round}`;slide.replaceElement(table.id,modified);
  const prepared=await session.prepareSave({mode:'edit'});assert.equal(prepared.status,'ready',JSON.stringify(prepared.report));const result=await session.save(prepared.plan);source=result.bytes;await writeFile(path.join(directory,`round${round}.pptx`),source);await writeFile(path.join(directory,`round${round}.json`),JSON.stringify(result.report,null,2)+'\n');
 }finally{session.dispose()}
}
console.log(directory);
