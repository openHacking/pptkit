import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import {
  authorPresentation,
  extractSource,
  inspectPptxPackage,
  measureImageDimensions,
  parseDeckSession,
  validateDeckSpec,
} from "../dist/index.js";

function deck(themeId = "clean-business") {
  return {
    brief: {
      title: "Browser workflow",
      audience: "Reviewers",
      purpose: "Exercise the portable runtime",
      language: "en-US",
      slideCountRange: [3, 3],
      themeId,
      imagePolicy: "Embedded assets",
      constraints: [],
    },
    slides: [
      { id: "cover", role: "cover", title: "Browser-first decks", subtitle: "Preview before download" },
      { id: "process", role: "process", title: "Review loop", steps: ["Import", "Preview", "Revise", "Download"] },
      { id: "closing", role: "closing", title: "Ready", message: "Generate only when approved." },
    ],
  };
}

test("authors the same portable deck across every theme", () => {
  for (const theme of ["clean-business", "swiss-grid", "editorial-story"]) {
    const spec = deck(theme);
    assert.deepEqual(validateDeckSpec(spec), []);
    const presentation = authorPresentation(spec);
    assert.equal(presentation.slides.length, 3);
    assert.equal(presentation.metadata.title, "Browser workflow");
  }
});

test("extracts text bytes without filesystem APIs", async () => {
  const source = await extractSource({ name: "brief.md", mimeType: "text/markdown", bytes: new TextEncoder().encode("# Evidence") }, 0);
  assert.equal(source.id, "src-01-brief");
  assert.equal(source.content, "# Evidence");
});

test("measures browser-neutral PNG, GIF, and SVG bytes", () => {
  const png = new Uint8Array(24);
  new DataView(png.buffer).setUint32(16, 640);
  new DataView(png.buffer).setUint32(20, 360);
  assert.deepEqual(measureImageDimensions({ name: "image.png", mimeType: "image/png", bytes: png }), { width: 640, height: 360 });
  const gif = new Uint8Array(10);
  new DataView(gif.buffer).setUint16(6, 320, true);
  new DataView(gif.buffer).setUint16(8, 180, true);
  assert.deepEqual(measureImageDimensions({ name: "image.gif", mimeType: "image/gif", bytes: gif }), { width: 320, height: 180 });
  const svg = new TextEncoder().encode('<svg viewBox="0 0 1200 675"></svg>');
  assert.deepEqual(measureImageDimensions({ name: "image.svg", mimeType: "image/svg+xml", bytes: svg }), { width: 1200, height: 675 });
});

test("validates session schema and inline asset limits", () => {
  const now = new Date().toISOString();
  const session = { schemaVersion: 1, id: "session-1", revision: 1, createdAt: now, updatedAt: now, deck: deck(), sources: [], assets: [] };
  assert.equal(parseDeckSession(session).id, "session-1");
  assert.throws(() => parseDeckSession({ ...session, schemaVersion: 2 }), /Unsupported/);
});

test("package inspection reports invalid bytes instead of throwing", () => {
  const result = inspectPptxPackage(new Uint8Array([1, 2, 3]));
  assert.equal(result.status, "checked");
  assert.equal(result.valid, false);
});

test("compiled runtime has no Node runtime imports", async () => {
  for (const file of ["author.js", "extraction.js", "session.js", "verify.js", "index.js"]) {
    const source = await readFile(new URL(`../dist/${file}`, import.meta.url), "utf8");
    assert.doesNotMatch(source, /(?:^|["'])node:/);
    assert.doesNotMatch(source, /\bprocess(?:\.|\[)/);
  }
});
