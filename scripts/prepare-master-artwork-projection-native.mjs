import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";

const [sourcePath,outputPath]=process.argv.slice(2);if(!sourcePath||!outputPath||process.argv.length!==4)throw Error("Expected source and output PPTX paths.");
const source=await readFile(sourcePath),session=await importPptx(source,{projection:"content",masterArtwork:"artwork-v1"});
const master=session.document.masters.find(m=>m.elements.some(e=>e.name==="logo"));assert.ok(master);assert.deepEqual(master.elements.map(e=>e.name),["blue","logo"]);
assert.equal(master.elements[1].crop.left,.2087);const result=await generatePptx(session.document);await writeFile(outputPath,result.bytes,{flag:"wx"});
console.log(JSON.stringify({sourceSha256:createHash("sha256").update(source).digest("hex"),outputSha256:createHash("sha256").update(result.bytes).digest("hex"),output:outputPath,slides:session.document.slides.length,warnings:result.warnings},null,2));session.dispose();
