import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
const [directory,executable='soffice']=process.argv.slice(2);if(!directory||process.argv.length>4)throw Error('Usage: node scripts/inspect-table-native-reopens.mjs directory [soffice]');
const execute=promisify(execFile),options={timeout:45000,maxBuffer:8*1024*1024},root=path.resolve(directory),out=path.join(root,'reopened-pdf');await mkdir(out,{recursive:true});const profile=`-env:UserInstallation=${pathToFileURL(path.join(root,'reopen-profile')).href}`,version=(await execute(executable,['--version'],options)).stdout.trim(),rounds=[];
for(let i=1;i<=3;i++){const input=path.join(root,'native',`round${i}.pptx`);const result=await execute(executable,[profile,'--headless','--convert-to','pdf:impress_pdf_Export','--outdir',out,input],options);rounds.push({round:i,stdout:result.stdout,stderr:result.stderr});}
await writeFile(path.join(root,'reopen-run.json'),JSON.stringify({version,rounds},null,2)+'\n');console.log(root);
