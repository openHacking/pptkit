import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
const [directory, executable='soffice']=process.argv.slice(2);if(!directory || process.argv.length>4)throw Error('Usage: node scripts/inspect-table-save-after-native.mjs directory [soffice]');
const execute=promisify(execFile), options={timeout:45000,maxBuffer:8*1024*1024},root=path.resolve(directory),native=path.join(root,'native'),pdf=path.join(root,'pdf');await mkdir(native,{recursive:true});await mkdir(pdf,{recursive:true});
const profile=`-env:UserInstallation=${pathToFileURL(path.join(root,'profile')).href}`,version=(await execute(executable,['--version'],options)).stdout.trim(),rounds=[];
for(let i=1;i<=3;i++){
 const input=path.join(root,`round${i}.pptx`);const a=await execute(executable,[profile,'--headless','--convert-to','pptx:Impress Office Open XML','--outdir',native,input],options);const b=await execute(executable,[profile,'--headless','--convert-to','pdf:impress_pdf_Export','--outdir',pdf,input],options);rounds.push({round:i,nativeStdout:a.stdout,pdfStdout:b.stdout,nativeStderr:a.stderr,pdfStderr:b.stderr});
}
await writeFile(path.join(root,'headless-run.json'),JSON.stringify({version,rounds},null,2)+'\n');console.log(root);
