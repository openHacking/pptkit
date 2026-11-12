import { createHash } from "node:crypto";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
const source = await readFile(".tmp/native-doughnut/powerpoint-TFHkIb/doughnut.pptx"), parts = readGeneratedZip(source), part = "ppt/charts/chart1.xml", xml = parts.get(part).toString();
const token = '<c:holeSize val="50"/>'; if (xml.split(token).length !== 2) throw new Error("Expected a unique 50% hole"); parts.set(part, Buffer.from(xml.replace(token, '<c:holeSize val="050%"/>')));
const bytes = zip([...parts].map(([name, body]) => ({ name, body }))).bytes; await mkdir(".tmp/native-doughnut-percent", { recursive: true }); const dir = path.resolve(await mkdtemp(".tmp/native-doughnut-percent/powerpoint-"));
await writeFile(path.join(dir, "percent.pptx"), bytes, { flag: "wx" }); await writeFile(path.join(dir, "manifest.json"), JSON.stringify({ sourceSha256: createHash("sha256").update(source).digest("hex"), sha256: createHash("sha256").update(bytes).digest("hex") }, null, 2) + "\n", { flag: "wx" }); console.log(dir);
