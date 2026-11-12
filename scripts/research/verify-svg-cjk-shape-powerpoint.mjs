import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
const sha = bytes => createHash("sha256").update(bytes).digest("hex");
const evidence = JSON.parse(await readFile(process.argv[2] ?? "docs/testing/tab-svg-cjk-shape-powerpoint-evidence.json", "utf8"));
execFileSync("python3", ["scripts/research/verify-cjk-shape-tabs-powerpoint.py"], { stdio: "pipe" });
assert.equal(evidence.source, "docs/testing/tab-cjk-shape-powerpoint-evidence.json");
const bytes = await readFile(evidence.source), source = JSON.parse(bytes); assert.equal(sha(bytes), evidence.sourceSha256);
assert.equal(sha(await readFile("scripts/inspect-svg-cjk-shape-powerpoint.mjs")), evidence.runnerSha256);
assert.deepEqual(evidence.bundles.map(({name,file})=>[name,file]), ["core","layout","svg-renderer"].map(name=>[`@pptkit/${name}`,`packages/${name}/dist/browser/index.js`]));
for (const b of evidence.bundles) assert.equal(sha(await readFile(b.file)), b.sha256);
assert.equal(evidence.font.file, "/System/Library/Fonts/Hiragino Sans GB.ttc"); assert.equal(sha(await readFile(evidence.font.file)), evidence.font.sha256);
assert.deepEqual(evidence.selectedFonts.map(row=>row.map(f=>f.postScriptName)), [["HiraginoSansGB-W3"],["HiraginoSansGB-W6"]]);
assert.ok(evidence.selectedFonts.every(row=>row.length===1 && !row[0].isCustomFont && row[0].glyphCount>0));
const expected=source.pages.flatMap(page=>page.rows.map(row=>({align:page.kind,word:row.word,bold:row.bold,characters:row.characters})));
assert.equal(expected.length,12); assert.equal(evidence.items.length,12);
for(const [i,item] of evidence.items.entries()) {
 const target=expected[i];for(const key of ["align","word","bold"]) assert.equal(item[key],target[key]);assert.equal(item.chars.length,target.characters.length);
 for(const [j,char] of item.chars.entries()) {assert.deepEqual(char.native,target.characters[j]);assert.equal(char.char,char.native.char);assert.equal(char.svg.char,char.char);for(const key of ["left","right","baseline"]) {assert.ok(Number.isFinite(char.svg[key]));assert.equal(char[`${key}Difference`],char.svg[key]-char.native[key]);}}
}
const chars=evidence.items.flatMap(i=>i.chars),horizontal=Math.max(...chars.flatMap(c=>[Math.abs(c.leftDifference),Math.abs(c.rightDifference)])),baseline=Math.max(...chars.map(c=>Math.abs(c.baselineDifference)));
assert.equal(chars.length,32);assert.equal(evidence.maxHorizontal,horizontal);assert.equal(evidence.maxBaseline,baseline);assert.ok(horizontal<.34 && baseline<.25);
console.log(JSON.stringify({verified:true,samples:12,characters:chars.length,maxHorizontal:horizontal,maxBaseline:baseline}));
