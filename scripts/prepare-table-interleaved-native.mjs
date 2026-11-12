import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { readFile, writeFile, mkdir, mkdtemp } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
const [input,executable='soffice']=process.argv.slice(2);if(!input||process.argv.length>4)throw Error('Usage: node scripts/prepare-table-interleaved-native.mjs native.pptx [soffice]');
const execute=promisify(execFile),options={timeout:45000,maxBuffer:8*1024*1024};await mkdir('.tmp/table-interleaved-native',{recursive:true});const directory=await mkdtemp(path.resolve('.tmp/table-interleaved-native/run-'));await mkdir(path.join(directory,'native'));await mkdir(path.join(directory,'pdf'));let source=await readFile(input);await writeFile(path.join(directory,'source.pptx'),source);const profile=`-env:UserInstallation=${pathToFileURL(path.join(directory,'profile')).href}`,version=(await execute(executable,['--version'],options)).stdout.trim(),runs=[];
for(let round=1;round<=3;round++){
 const session=await importPptx(source,{projection:'content',editSaveProfile:'text-runs-v3'});let bytes;
 try{
  const slide=session.document.slides[0],previous=round===1?'Anchor 123':`Interleaved anchor ${round-1}`,previousRight=round===1?'Cell 456':`Interleaved cell ${round-1}`;
  const table=slide.elements.find(e=>e.type==='table'&&Array.isArray(e.rows[0]?.cells[0]?.content)&&e.rows[0].cells[0].content[0]?.runs[0]?.text===previous);assert.ok(table);
  const modified=structuredClone(table),anchor=modified.rows[0].cells[0],right=modified.rows[1].cells[0];assert.equal(anchor.rowSpan,2);assert.equal(anchor.colSpan,2);assert.equal(right.content[0].runs[0].text,previousRight);anchor.content[0].runs[0].text=`Interleaved anchor ${round}`;right.content[0].runs[0].text=`Interleaved cell ${round}`;slide.replaceElement(table.id,modified);
  const prepared=await session.prepareSave({mode:'edit'});assert.equal(prepared.status,'ready',JSON.stringify(prepared.report));const result=await session.save(prepared.plan);bytes=result.bytes;await writeFile(path.join(directory,`round${round}.pptx`),bytes);await writeFile(path.join(directory,`round${round}.json`),JSON.stringify(result.report,null,2)+'\n');
 }finally{session.dispose()}
 const candidate=path.join(directory,`round${round}.pptx`),native=path.join(directory,'native'),pdf=path.join(directory,'pdf');const a=await execute(executable,[profile,'--headless','--convert-to','pptx:Impress Office Open XML','--outdir',native,candidate],options);const b=await execute(executable,[profile,'--headless','--convert-to','pdf:impress_pdf_Export','--outdir',pdf,candidate],options);source=await readFile(path.join(native,`round${round}.pptx`));runs.push({round,nativeStdout:a.stdout,nativeStderr:a.stderr,pdfStdout:b.stdout,pdfStderr:b.stderr});
}
await writeFile(path.join(directory,'headless-run.json'),JSON.stringify({version,runs},null,2)+'\n');console.log(directory);
