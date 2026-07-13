import test from "node:test";
import assert from "node:assert/strict";

import { AssetRegistry } from "../dist/document/asset-registry.js";

test("AssetRegistry deduplicates by dedupe key and source", () => {
  const registry = new AssetRegistry();
  const first = registry.register({
    kind: "image",
    source: { type: "url", value: "https://example.com/hero.png" },
    dedupeKey: "hero",
  });
  const second = registry.register({
    kind: "image",
    source: { type: "url", value: "https://example.com/hero.png" },
    dedupeKey: "hero",
  });

  assert.equal(second, first);
  assert.equal(registry.assets.length, 1);
});

test("AssetRegistry rejects duplicate identifiers with different metadata", () => {
  const registry = new AssetRegistry();
  registry.register({ kind: "image", id: "logo", source: { type: "path", value: "./logo.png" } });

  assert.throws(
    () => registry.register({ kind: "image", id: "logo", source: { type: "path", value: "./other.png" } }),
    /already registered with different metadata/,
  );
});
