import assert from "node:assert/strict";
import { readFile,writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { join } from "node:path";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
const dir=process.argv[2];if(!dir||process.argv.length!==3)throw Error("Expected evidence directory.");
const bytes=await readFile(join(dir,"master-image-native-edit.pptx"));const session=await importPptx(bytes,{projection:"content"});
try {
 const masters=session.document.masters??[];assert.equal(masters.length,0);
 const ready=await session.prepareSave({mode:"no-op"});assert.equal(ready.status,"ready");const saved=await session.save(ready.plan);assert.deepEqual(Buffer.from(saved.bytes),bytes);
 const report={inputSha256:createHash("sha256").update(bytes).digest("hex"),projectedMasterCount:masters.length,noOpByteIdentical:true,scope:"content projection does not reconstruct source master artwork"};
 await writeFile(join(dir,"native-edit-parser.json"),JSON.stringify(report,null,2)+"\n");console.log(JSON.stringify(report));
} finally {session.dispose()}
