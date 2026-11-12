import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { compareOpcPackages } from "../packages/pptx-parser/dist/index.js";

if (process.argv.length !== 4) throw new Error("Usage: node scripts/inspect-package-comparison.mjs before.pptx after.pptx");
const inputs = await Promise.all(process.argv.slice(2).map((file) => readFile(file)));
const sha256 = (bytes) => createHash("sha256").update(bytes).digest("hex");
const profiles = {};
for (const profile of ["opc-payload-v1", "opc-xml-v1"]) {
  profiles[profile] = await compareOpcPackages(inputs[0], inputs[1], { profile });
}
console.log(JSON.stringify({
  schemaVersion: 1,
  scope: "Independent source comparison only; no application was launched or native fidelity established.",
  beforeSha256: sha256(inputs[0]),
  afterSha256: sha256(inputs[1]),
  profiles,
}, null, 2));
