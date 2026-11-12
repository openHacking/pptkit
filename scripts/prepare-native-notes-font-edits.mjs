import assert from "node:assert/strict";
import { readFile, writeFile, mkdir, mkdtemp, copyFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";

if(process.argv.length!==3)throw Error("Usage: node scripts/prepare-native-notes-font-edits.mjs <native-notes-font-run-directory>");
const original=path.resolve(process.argv[2]),manifest=JSON.parse(await readFile("docs/testing/native-notes-font-evidence.json","utf8")),sha=b=>createHash("sha256").update(b).digest("hex");
await mkdir(".tmp/native-notes-font-edits",{recursive:true});const directory=await mkdtemp(path.resolve(".tmp/native-notes-font-edits/run-"));
for(const variant of ["A","B"]){
  const from=path.join(original,variant),to=path.join(directory,variant);await mkdir(to);
  const expected=JSON.parse(await readFile(path.join(from,"expected.json"),"utf8")),evidence=manifest.groups.find(g=>g.variant===variant).independent.rounds[2];
  const file=path.join(from,"native-round3.pptx"),pdf=path.join(from,"native-round3-notes.pdf");
  assert.equal(sha(await readFile(file)),evidence.package.sha256);assert.equal(sha(await readFile(pdf)),evidence.pdf.sha256);
  await copyFile(file,path.join(to,"baseline.pptx"));await copyFile(pdf,path.join(to,"baseline-notes.pdf"));
  await writeFile(path.join(to,"expected.json"),JSON.stringify({...expected,nativeBaseline:{directory:from,file,sha256:evidence.package.sha256,pdfSha256:evidence.pdf.sha256},editTargets:[3,5,6,7,8,10,11,12,13,14,15]},null,2)+"\n");
}
await writeFile(".tmp/native-notes-font-edits/latest.txt",directory);console.log(directory);
