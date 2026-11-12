import { createHash } from "node:crypto";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";

const source = await readFile(".tmp/native-doughnut/powerpoint-TFHkIb/doughnut.pptx"), parts = readGeneratedZip(source);
for (const [index, names] of [[1, ["firstSliceAng"]], [2, ["holeSize"]], [3, ["firstSliceAng", "holeSize"]]]) {
  const key = `ppt/charts/chart${index}.xml`; let xml = parts.get(key).toString();
  for (const name of names) { const pattern = new RegExp(`<c:${name} val="[0-9]+"/>`, "g"); if ([...xml.matchAll(pattern)].length !== 1) throw new Error(`Expected one ${name} in chart ${index}`); xml = xml.replace(pattern, ""); }
  parts.set(key, Buffer.from(xml));
}
const bytes = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
await mkdir(".tmp/native-doughnut-defaults", { recursive: true }); const dir = path.resolve(await mkdtemp(".tmp/native-doughnut-defaults/powerpoint-"));
await writeFile(path.join(dir, "defaults.pptx"), bytes, { flag: "wx" });
await writeFile(path.join(dir, "manifest.json"), JSON.stringify({ sourceSha256: createHash("sha256").update(source).digest("hex"), sha256: createHash("sha256").update(bytes).digest("hex"), removed: [["firstSliceAng"], ["holeSize"], ["firstSliceAng", "holeSize"]] }, null, 2) + "\n", { flag: "wx" });
console.log(dir);
