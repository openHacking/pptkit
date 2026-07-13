import test from "node:test";
import assert from "node:assert/strict";

import { createPresentation } from "../dist/index.js";

test("asset registration deduplicates immutable assets", () => {
  const presentation = createPresentation();
  const first = presentation.registerAsset({
    kind: "image",
    source: { type: "url", value: "https://example.com/hero.png" },
    dedupeKey: "hero",
    accessibility: { description: "Hero" },
  });
  const second = presentation.registerAsset({
    kind: "image",
    source: { type: "url", value: "https://example.com/hero.png" },
    dedupeKey: "hero",
    accessibility: { description: "Hero" },
  });
  assert.equal(first, second);
  assert.equal(presentation.assets.length, 1);
  assert.throws(() => { first.source.value = "changed"; }, TypeError);
});

test("asset registration rejects conflicting metadata", () => {
  const presentation = createPresentation();
  presentation.registerAsset({ kind: "image", id: "logo", source: { type: "path", value: "./logo.png" } });
  assert.throws(
    () => presentation.registerAsset({ kind: "image", id: "logo", source: { type: "path", value: "./other.png" } }),
    /already registered with different metadata/,
  );
});
