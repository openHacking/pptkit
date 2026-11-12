import { createHash } from "node:crypto";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
const source = await readFile(".tmp/native-doughnut/powerpoint-TFHkIb/doughnut.pptx"), parts = readGeneratedZip(source);
for (const [i, tags] of [[1,["firstSliceAng"]],[2,["holeSize"]],[3,["firstSliceAng","holeSize"]]]) {
 const key=`ppt/charts/chart${i}.xml`; let xml=parts.get(key).toString(); for(const tag of tags){const pattern=new RegExp(`<c:${tag} val="[0-9]+"/>`,"g"); if([...xml.matchAll(pattern)].length!==1) throw Error(`Expected ${tag}`); xml=xml.replace(pattern,`<c:${tag}/>`);} parts.set(key,Buffer.from(xml));
}
const bytes=zip([...parts].map(([name,body])=>({name,body}))).bytes; await mkdir('.tmp/native-doughnut-empty-geometry',{recursive:true});const dir=path.resolve(await mkdtemp('.tmp/native-doughnut-empty-geometry/powerpoint-'));
await writeFile(path.join(dir,'empty.pptx'),bytes,{flag:'wx'});await writeFile(path.join(dir,'manifest.json'),JSON.stringify({sourceSha256:createHash('sha256').update(source).digest('hex'),sha256:createHash('sha256').update(bytes).digest('hex')},null,2)+'\n',{flag:'wx'});console.log(dir);
