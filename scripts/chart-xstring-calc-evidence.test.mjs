import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const verify = file => execFileSync("python3", ["scripts/research/verify-chart-xstring-calc.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] });
const hash = bytes => createHash("sha256").update(bytes).digest("hex");

test("retained Calc evidence verifies actual files and rejects forged observations and rehashed package damage", async t => {
  const result = JSON.parse(verify()); assert.equal(result.verified, true); assert.equal(result.records, 4); assert.equal(result.changedCell, "A4");
  const directory = await mkdtemp(path.resolve(".tmp/chart-xstring-calc/check-")); t.after(() => rm(directory, { recursive: true, force: true }));
  const evidence = JSON.parse(await readFile("docs/testing/chart-xstring-calc-evidence.json", "utf8"));
  const write = async (name, value) => { const target = path.join(directory, name); await writeFile(target, JSON.stringify(value)); return target; };
  for (const mutate of [x => { x.records[1].cells.A4.decodedOnce = x.records[0].cells.A4.decodedOnce; }, x => { x.records[2].sha256 = "0".repeat(64); }, x => { x.records.pop(); }, x => { x.logs[0].exitCode = 1; }]) {
    const altered = structuredClone(evidence); mutate(altered); const file = await write("bad.json", altered); assert.throws(() => verify(file));
  }
  const altered = structuredClone(evidence), entries = readGeneratedZip(await readFile(altered.records[0].path));
  const sheet = entries.get("xl/worksheets/sheet1.xml").toString(); assert.ok(sheet.includes("plain &amp; 中文"));
  entries.set("xl/worksheets/sheet1.xml", Buffer.from(sheet.replace("plain &amp; 中文", "tampered")));
  const bytes = zip([...entries].map(([name, body]) => ({ name, body }))).bytes, damaged = path.join(directory, "damaged.xlsx"); await writeFile(damaged, bytes);
  altered.records[0].path = path.relative(process.cwd(), damaged); altered.records[0].sha256 = hash(bytes);
  const file = await write("damaged.json", altered); assert.throws(() => verify(file));
});
