import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { importPptx, describeImportCapabilities } from "../packages/pptx-parser/dist/index.js";
import { diagramTextFixture } from "../packages/pptx-parser/test/diagram-text-fixture.mjs";

test("Node/browser ESM/global capability tables retain public source evidence and complete operations", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage(); const requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile(new URL("../packages/pptx-parser/dist/browser/global.js", import.meta.url), "utf8") });
  const bundle = await readFile(new URL("../packages/pptx-parser/dist/browser/index.js", import.meta.url), "utf8");
  await page.evaluate(async source => { globalThis.parser = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`); }, bundle);
  for (const strict of [false, true]) {
    const fixture = diagramTextFixture(strict), bytes = fixture.bytes(), options = { projection: "content", diagramPreview: "cache", editSaveProfile: "text-runs-v3" }, session = await importPptx(bytes, options);
    try {
      const expected = describeImportCapabilities(session.report), result = await page.evaluate(async ({ bytes, options }) => {
        const outputs = [];
        for (const api of [globalThis.parser, globalThis.PPTKit]) {
          const session = await api.importPptx(Uint8Array.from(bytes), options);
          try {
            const before = JSON.stringify(session.report), capability = api.describeImportCapabilities(session.report);
            const grouped = capability.subjects.flatMap(subject => subject.operations.flatMap(operation => operation.observations.map(observation => ({ feature: subject.feature, source: subject.source, slideId: subject.slideId, elementId: subject.elementId, operation: operation.operation, ...observation }))));
            const key = f => JSON.stringify([f.feature, f.source.part, f.source.path ?? null, f.source.sourceObjectId ?? null, f.source.relationshipId ?? null, f.source.elementIndex ?? null, f.slideId ?? null, f.elementId ?? null, f.operation, f.id, f.tier, f.reason]);
            const expectedKeys = session.report.findings.map(key).sort(), actualKeys = grouped.map(key).sort();
            if (JSON.stringify(actualKeys) !== JSON.stringify(expectedKeys)) throw new Error("Finding provenance lost or invented");
            if (!capability.subjects.every(s => s.operations.length === 7)) throw new Error("Incomplete operation vectors");
            if (JSON.stringify(session.report) !== before) throw new Error("Input report mutated");
            const plan = await session.prepareSave({ mode: "no-op" }), saved = await session.save(plan.plan);
            if (saved.bytes.some((value, i) => value !== bytes[i])) throw new Error("No-op source changed");
            outputs.push(capability);
          } finally { session.dispose(); }
        }
        return outputs;
      }, { bytes: [...bytes], options });
      for (const actual of result) assert.deepEqual(actual, expected);
    } finally { session.dispose(); }
  }
  const features = ['first-slide-number', 'master-artwork.shape', 'artwork.background', 'artwork.background.unmapped-property', 'artwork.backgroundish'];
  const report = { projection: 'content', findings: features.map((feature, index) => ({ id: `finding-${index}`, feature, source: { part: '/fixture.xml' }, operation: 'import', tier: 'fallback', reason: 'Observed in fixture' })) };
  const expectedReferences = describeImportCapabilities(report);
  const browserReferences = await page.evaluate(report => [globalThis.parser, globalThis.PPTKit].map(api => api.describeImportCapabilities(report)), report);
  for (const actual of browserReferences) {
    // Sorting comparison counts are runtime-dependent; the public rows are canonical.
    const { workUnits: _browserWork, ...browserView } = actual;
    const { workUnits: _nodeWork, ...nodeView } = expectedReferences;
    assert.deepEqual(browserView, nodeView);
  }
  assert.equal(expectedReferences.subjects.find(subject => subject.feature === 'artwork.backgroundish').references, undefined);
  assert.deepEqual(requests, []);
});
