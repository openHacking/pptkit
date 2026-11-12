import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "@pptkit/core";

function fixture() {
  const document = createPresentation({ size: { width: 200, height: 100 } });
  document.addSlide({ id: "slide" }).addElement({ type: "text", id: "label", content: "Measured text", box: { x: 190, y: 0, width: 50, height: 20 }, frame: { margin: 0 } });
  return document;
}
const layout = { textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 60, height: 30 } }) }, diagnostics: { textOverflow: true, slideBounds: true } };

import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { generatePptx } from "../dist/index.js";
import { generatePptx as generateNode, writePptx } from "../dist/node.js";

test("analysis warnings propagate through byte and file exports without changing package content", async () => {
  const document = fixture();
  const before = normalizePresentation(document);
  const plain = await generatePptx(document);
  const analyzed = await generatePptx(document, { layout });
  assert.deepEqual(analyzed.bytes, plain.bytes);
  assert.equal(analyzed.status, "generated-with-warnings");
  assert.deepEqual(analyzed.warnings.map(({ code }) => code), ["slide-overflow", "text-overflow", "slide-overflow"]);
  assert.ok(analyzed.warnings.every((warning) => warning.slideId === "slide" && warning.elementId === "label" && warning.path.startsWith("slides.0.elements.0.")));
  const native = await generateNode(document, { layout });
  assert.deepEqual(native, analyzed);
  const directory = await mkdtemp(join(tmpdir(), "pptkit-layout-"));
  try {
    const output = join(directory, "analysis.pptx");
    const written = await writePptx(document, { output, layout });
    assert.equal(written.status, "written-with-warnings");
    assert.deepEqual(written.warnings, analyzed.warnings);
    assert.deepEqual(new Uint8Array(await readFile(output)), analyzed.bytes);
  } finally { await rm(directory, { recursive: true, force: true }); }
  assert.deepEqual(normalizePresentation(document), before);
});

test("measurement failure is recoverable but invalid configuration rejects export", async () => {
  const result = await generatePptx(fixture(), { layout: { textMeasurement: { measure() { throw new Error("offline"); } } } });
  assert.equal(result.warnings[0].code, "text-measurement-failed");
  assert.ok(result.bytes.length > 0);
  await assert.rejects(generatePptx(fixture(), { layout: { diagnostics: { tolerance: -1 } } }), /tolerance/);
});
