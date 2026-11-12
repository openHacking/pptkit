import assert from "node:assert/strict";
import { readFile, writeFile, mkdir, mkdtemp } from "node:fs/promises";
import path from "node:path";
import { createHash } from "node:crypto";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";

const hash = b => createHash("sha256").update(b).digest("hex"), inputs = [];
for (const [name, stage, nativeName] of [["B", "styles", "regular-round3"], ["override", "themes", "native-baseline-round3"]]) {
  const manifestPath = "docs/testing/native-notes-" + stage + "-evidence.json", manifestBytes = await readFile(manifestPath), manifest = JSON.parse(manifestBytes);
  const group = manifest.groups.find(g => (g.variant ?? g.group) === name), folder = path.join(manifest.root, name);
  const bytes = await readFile(path.join(folder, nativeName + ".pptx")), pdf = await readFile(path.join(folder, nativeName + "-notes.pdf"));
  const input = group.inputs.find(i => name === "B" ? i.stage === "regular" && i.round === 3 : i.name === nativeName);
  const recordedPdf = name === "B" ? group.independent.stages.find(s => s.stage === "regular").rounds[2].pdf.sha256 : group.independent.baselineRounds[2].pdf.sha256;
  assert.equal(hash(bytes), input.sha256); assert.equal(hash(pdf), recordedPdf);
  const controls = name === "B" ? JSON.parse(await readFile(path.join(folder, "expected.json"))).controls.map(c => ({ ...c, variant: "B" })) : JSON.parse(await readFile(path.join(folder, "controls.json")));
  const parts = readGeneratedZip(bytes), master = "ppt/notesMasters/notesMaster1.xml", original = parts.get(master).toString();
  let matches = 0;
  const changed = original.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, shape => {
    if (!shape.includes('type="sldImg"')) return shape;
    matches++; assert.ok(!shape.includes("<a:prstGeom"));
    return shape.replace("</p:spPr>", '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>');
  });
  assert.equal(matches, 1); parts.set(master, Buffer.from(changed));
  inputs.push({ name, bytes, pdf, rect: zip([...parts].map(([name, body]) => ({ name, body }))).bytes, controls, provenance: { source: path.join(folder, nativeName + ".pptx"), sourceSha256: hash(bytes), sourcePdfSha256: hash(pdf), manifest: path.resolve(manifestPath), manifestSha256: hash(manifestBytes), operation: "Only insert master slide-image rect; all other payloads retained exactly. Existing native B/independent per-notes override baseline, not fresh author output." } });
}
const root = path.resolve(".tmp/notes-image-domains"); await mkdir(root, { recursive: true });
const directory = await mkdtemp(path.join(root, "run-"));
for (const input of inputs) {
  const folder = path.join(directory, input.name); await mkdir(folder);
  await writeFile(path.join(folder, "baseline.pptx"), input.bytes); await writeFile(path.join(folder, "baseline-notes.pdf"), input.pdf);
  await writeFile(path.join(folder, "rect.pptx"), input.rect);
  await writeFile(path.join(folder, "expected.json"), JSON.stringify({ name: input.name, controls: input.controls, ...input.provenance, rectSha256: hash(input.rect) }, null, 2) + "\n");
}
await writeFile(path.join(root, "latest.txt"), directory + "\n"); console.log(directory);
