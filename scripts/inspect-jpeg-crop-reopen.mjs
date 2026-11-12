import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { join } from "node:path";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
const dir = process.argv[2]; if (!dir || process.argv.length !== 3) throw Error("Expected evidence directory.");
const input = await readFile(join(dir,"crop-native-round1.pptx"));
const session = await importPptx(input,{projection:"content"});
try {
  const images=session.document.slides[0].elements.filter(element=>element.type==="image");
  assert.equal(images.length,2);
  const mapped=images.find(image=>image.name==="mapped"), control=images.find(image=>image.name==="control");
  assert.ok(mapped && control);
  assert.equal(mapped.fit,"crop"); assert.deepEqual(mapped.crop,{left:0,top:0,right:0,bottom:0.33333});
  assert.deepEqual(mapped.box,{x:120,y:130,width:240,height:240});
  assert.equal(control.fit,"stretch"); assert.deepEqual(control.box,{x:600,y:130,width:240,height:240});
  for(const [image,file,width,height] of [[mapped,"browser-normalized.png",200,300],[control,"crop-control.png",200,200]]) {
    const asset=session.document.assets.find(value=>value.id===image.assetId);
    assert.ok(asset);assert.equal(asset.width,width);assert.equal(asset.height,height);assert.equal(asset.mimeType,"image/png");
    assert.deepEqual(Buffer.from(asset.source.value,"base64"),await readFile(join(dir,file)));
  }
  const prepared=await session.prepareSave({mode:"no-op"});assert.equal(prepared.status,"ready");
  const saved=await session.save(prepared.plan);assert.deepEqual(Buffer.from(saved.bytes),input);
  const hash = data => createHash("sha256").update(data).digest("hex");
  const report={inputSha256:hash(input),mappedCrop:mapped.crop,mappedBox:mapped.box,controlBox:control.box,assets:session.document.assets.map(({width,height,mimeType})=>({width,height,mimeType})),noOpByteIdentical:true};
  await writeFile(join(dir,"crop-native-parser.json"),JSON.stringify(report,null,2)+"\n");console.log(JSON.stringify(report));
} finally { session.dispose(); }
