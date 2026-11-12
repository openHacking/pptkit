import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { inspectOpcClosure } from "../packages/pptx-parser/dist/index.js";

if (process.argv.length < 4) throw new Error("Usage: node scripts/inspect-preservation-closure.mjs source.pptx /logical/root [...]");
const bytes = await readFile(process.argv[2]);
const inspection = await inspectOpcClosure(bytes, { roots: process.argv.slice(3) });
console.log(JSON.stringify({
  schemaVersion: 1,
  scope: "Source relationship inspection only; no native application launch, extraction, cleanup or editing.",
  sourceSha256: createHash("sha256").update(bytes).digest("hex"),
  inspection,
}, null, 2));
