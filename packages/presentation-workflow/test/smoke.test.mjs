import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { normalizePresentation } from "@pptkit/core";

import {
  authorDeck,
  extractSource,
  inspectPptxPackage,
  inspectStructure,
  measureImageDimensions,
  parseDeckSession,
  planDeckLayout,
  validateDeckSpec,
} from "../dist/index.js";
import { RECIPE_REGISTRY } from "../dist/authoring/registry.js";

function deck(themeId = "clean-business") {
  return {
    design: { theme: { id: themeId }, seed: "browser-workflow", variation: "balanced" },
    brief: {
      title: "Browser workflow",
      audience: "Reviewers",
      purpose: "Exercise the portable runtime",
      language: "en-US",
      slideCountRange: [3, 3],
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

function allRolesDeck(themeId = "clean-business") {
  return {
    design: { theme: { id: themeId }, seed: "all-roles", variation: "balanced" },
    brief: {
      title: "Design system fixture",
      audience: "Reviewers",
      purpose: "Exercise every deterministic role recipe",
      language: "en-US",
      slideCountRange: [10, 10],
      imagePolicy: "No images",
      constraints: [],
    },
    slides: [
      { id: "cover", role: "cover", title: "A clearer presentation system", subtitle: "Hierarchy, rhythm, and provenance" },
      { id: "agenda", role: "agenda", title: "The argument", items: ["Frame the decision", "Show the evidence", "Make the next move clear"] },
      { id: "section", role: "section", title: "01 / Direction", message: "A design system should make the message easier to see." },
      { id: "statement", role: "statement", title: "The core judgment", message: "Structure turns evidence into action.", items: ["One message", "One reading path", "One deliberate composition"] },
      { id: "image", role: "image", title: "Evidence belongs in the layout", message: "Images support the argument.", items: ["Preserve aspect ratio", "Use a stable slot"] },
      { id: "kpi", role: "kpi", title: "Signals that matter", kpis: [{ value: "+18%", label: "Activation", detail: "Quarter over quarter" }, { value: "72%", label: "Retention" }, { value: "3.4×", label: "Velocity" }] },
      { id: "comparison", role: "comparison", title: "Two approaches", comparison: { left: { heading: "Before", items: ["Repeated cards", "Weak hierarchy"] }, right: { heading: "After", items: ["Theme-specific structure", "Readable scale"] } } },
      { id: "process", role: "process", title: "A deliberate workflow", steps: ["Frame", "Outline", "Compose", "Review"] },
      { id: "table", role: "table", title: "Delivery ledger", table: { headers: ["Artifact", "Role"], rows: [["Brief", "Locks the purpose"], ["Outline", "Locks the argument"], ["Preview", "Reveals layout issues"]] } },
      { id: "closing", role: "closing", title: "Make the next decision obvious.", message: "Review the evidence, then act." },
    ],
  };
}

function saasHuntDeck() {
  const sourceRefs = [{ id: "src-01-slides", label: "slides.md" }];
  return {
    design: { theme: { id: "swiss-grid" }, seed: "saas-hunt", variation: "balanced" },
    brief: {
      title: "SaaS Hunt", audience: "Founders and product teams",
      purpose: "Explain continuous product discovery", language: "en-US",
      slideCountRange: [6, 6], imagePolicy: "No images", constraints: [],
    },
    slides: [
      { id: "cover", role: "cover", title: "SaaS Hunt", subtitle: "Launch and grow your SaaS earlier", sourceRefs },
      { id: "why", role: "agenda", title: "Why SaaS Hunt?", items: ["Gain visibility before launch", "Build an audience early", "Generate long-term organic traffic", "Help users discover innovative SaaS products"], sourceRefs },
      { id: "platform", role: "statement", title: "A platform for discovery before and after release", message: "SaaS Hunt helps people launch, discover, and promote SaaS products through one continuous discovery loop.", sourceRefs },
      { id: "features", role: "table", title: "Core features", table: { headers: ["Feature", "Role"], rows: [["Product Launch Pages", "Present the launch story"], ["Coming Soon Listings", "Collect interest early"], ["Product Categories", "Create browsing paths"], ["Search & Discovery", "Find relevant SaaS quickly"], ["Product Showcase", "Keep products visible"]] }, sourceRefs },
      { id: "difference", role: "comparison", title: "How it differs", comparison: { left: { heading: "Product Hunt", items: ["Launch-day momentum", "Short visibility spike"] }, right: { heading: "SaaS Hunt", items: ["Continuous discovery", "Long-term visibility"] } }, sourceRefs },
      { id: "close", role: "closing", title: "Build momentum before launch, not after.", message: "saashunt.pro", sourceRefs },
    ],
  };
}

function presentationOf(spec, resolveAsset) {
  return authorDeck(spec, resolveAsset).presentation;
}

test("authors the same portable deck across every theme", () => {
  for (const theme of ["clean-business", "swiss-grid", "editorial-story"]) {
    const spec = deck(theme);
    assert.deepEqual(validateDeckSpec(spec), []);
    const presentation = presentationOf(spec);
    assert.equal(presentation.slides.length, 3);
    assert.equal(presentation.metadata.title, "Browser workflow");
  }
});

test("registers multiple stable recipes for every semantic role", () => {
  for (const role of ["cover", "agenda", "section", "statement", "image", "kpi", "comparison", "process", "table", "closing"]) {
    const recipes = RECIPE_REGISTRY.filter((candidate) => candidate.role === role);
    assert.ok(recipes.length >= 2, `${role} should expose at least two internal recipes`);
    assert.equal(new Set(recipes.map((candidate) => candidate.id)).size, recipes.length);
  }
  for (const role of ["statement", "image", "process", "table"]) {
    assert.ok(RECIPE_REGISTRY.filter((candidate) => candidate.role === role).length >= 3, `${role} should expose at least three recipes`);
  }
});

test("plans deterministic, explainable layouts from seed and explicit intent", () => {
  const spec = allRolesDeck("clean-business");
  spec.design = { ...spec.design, seed: "stable-seed", variation: "expressive" };
  assert.deepEqual(planDeckLayout(spec), planDeckLayout(structuredClone(spec)));
  const statement = spec.slides.find((slide) => slide.id === "statement");
  statement.composition = "split";
  const decision = planDeckLayout(spec).find((candidate) => candidate.slideId === "statement");
  assert.equal(decision.composition, "split");
  assert.match(decision.reason, /Explicit split/);

  const recipeIds = new Set();
  for (let index = 0; index < 24; index += 1) {
    const varied = deck();
    varied.design = { ...varied.design, seed: `seed-${index}`, variation: "expressive" };
    recipeIds.add(planDeckLayout(varied)[0].recipeId);
  }
  assert.ok(recipeIds.size > 1, "different seeds should explore compatible cover recipes");
  const coverRecipeIds = new Set(RECIPE_REGISTRY.filter((candidate) => candidate.role === "cover").map((candidate) => candidate.id));
  for (const recipeId of recipeIds) assert.ok(coverRecipeIds.has(recipeId));
});

test("rejects incompatible compositions and unsafe theme overrides", () => {
  const spec = deck();
  spec.slides[1].composition = "image-hero";
  spec.design = { ...spec.design, theme: { ...spec.design.theme, overrides: { colors: { text: "777777", background: "888888" }, fonts: { heading: "" } } } };
  const issues = validateDeckSpec(spec);
  assert.ok(issues.some((issue) => issue.code === "incompatible-composition"));
  assert.throws(
    () => authorDeck(spec),
    /No layout recipe supports process slide process with composition image-hero density balanced/,
  );
  assert.ok(issues.some((issue) => issue.code === "theme-contrast"));
  assert.ok(issues.some((issue) => issue.code === "invalid-theme-font"));
});

test("authors every role without structural failures across every theme", () => {
  for (const theme of ["clean-business", "swiss-grid", "editorial-story"]) {
    const spec = allRolesDeck(theme);
    assert.deepEqual(validateDeckSpec(spec), []);
    const presentation = presentationOf(spec);
    assert.deepEqual(inspectStructure(presentation), []);
    assert.equal(presentation.slides.length, 10);
  }
});

test("authors every registered recipe without geometry errors across themes", () => {
  for (const theme of ["clean-business", "swiss-grid", "editorial-story"]) {
    const source = allRolesDeck(theme);
    for (const recipe of RECIPE_REGISTRY) {
      const base = source.slides.find((slide) => slide.role === recipe.role);
      const slide = structuredClone(base);
      slide.composition = recipe.composition;
      slide.density = "balanced";
      if (recipe.id === "image-hero") slide.items = undefined;
      if (recipe.role === "image") slide.image = { assetId: "fixture", alt: "Fixture", width: 1200, height: 675 };
      if (recipe.id === "table-split-chart") {
        slide.table = undefined;
        slide.chart = { type: "bar", categories: ["A", "B", "C"], series: [{ name: "Value", values: [1, 2, 3] }] };
      }
      const spec = { ...source, brief: { ...source.brief, slideCountRange: [1, 1] }, slides: [slide], design: { ...source.design, seed: recipe.id } };
      assert.deepEqual(validateDeckSpec(spec, new Set(["fixture"])).filter((issue) => issue.severity === "error"), [], `${theme}/${recipe.id}`);
      assert.equal(planDeckLayout(spec)[0].recipeId, recipe.id, `${theme}/${recipe.id}`);
      const presentation = presentationOf(spec, () => ({
        source: { type: "url", value: "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='1200' height='675'/%3E" },
        mimeType: "image/svg+xml",
      }));
      assert.deepEqual(inspectStructure(presentation).filter((issue) => issue.severity === "error"), [], `${theme}/${recipe.id}`);
    }
  }
});

test("renders the closing title exactly once across every theme", () => {
  for (const theme of ["clean-business", "swiss-grid", "editorial-story"]) {
    const spec = deck(theme);
    const normalized = normalizePresentation(presentationOf(spec));
    const closing = normalized.slides.at(-1);
    const titleElements = closing.elements.filter(
      (element) => element.type === "text" && element.plainText === "Ready",
    );

    assert.deepEqual(titleElements.map((element) => element.name), ["Closing message"]);
  }
});

test("keeps provenance in speaker notes and out of visible slide text", () => {
  const spec = deck("swiss-grid");
  spec.slides[1].notes = "Presenter context.";
  spec.slides[1].sourceRefs = [{ id: "src-01-slides", label: "slides.md" }];
  const normalized = normalizePresentation(presentationOf(spec));
  const visible = normalized.slides[1].elements
    .filter((element) => element.type === "text")
    .map((element) => element.plainText)
    .join("\n");
  const notes = normalized.slides[1].notes.flatMap((paragraph) => paragraph.runs.map((run) => run.text)).join("\n");
  assert.doesNotMatch(visible, /Sources:|src-01-slides|slides\.md/);
  assert.match(notes, /Presenter context/);
  assert.match(notes, /Source references \(provenance only/);
  assert.match(notes, /src-01-slides/);
  assert.match(notes, /slides\.md/);
});

test("uses presentation-scale font floors for ordinary visible copy", () => {
  const normalized = normalizePresentation(presentationOf(allRolesDeck("clean-business")));
  for (const slide of normalized.slides) {
    for (const element of slide.elements) {
      if (element.type === "text" && !element.name.startsWith("Metadata")) {
        for (const paragraph of element.content) {
          for (const run of paragraph.runs) assert.ok(run.style.fontSize >= 15, `${slide.id}/${element.name} uses ${run.style.fontSize} pt`);
        }
      }
      if (element.type === "table") {
        for (const row of element.rows) for (const cell of row.cells) for (const paragraph of cell.content) for (const run of paragraph.runs) {
          assert.ok(run.style.fontSize >= 15, `${slide.id}/table uses ${run.style.fontSize} pt`);
        }
      }
    }
  }
});

test("inherits PowerPoint list indentation instead of overriding it in the workflow", () => {
  const normalized = normalizePresentation(presentationOf(allRolesDeck("clean-business")));
  const bulletParagraphs = normalized.slides
    .flatMap((slide) => slide.elements)
    .filter((element) => element.type === "text")
    .flatMap((element) => element.content)
    .filter((paragraph) => paragraph.style.bullet.type === "bullet");
  assert.ok(bulletParagraphs.length > 0);
  for (const paragraph of bulletParagraphs) {
    assert.equal(paragraph.style.indent, 27);
    assert.equal(paragraph.style.hanging, 27);
  }
});

test("keeps Clean Business geometry square and restrained", () => {
  const normalized = normalizePresentation(presentationOf(allRolesDeck("clean-business")));
  const roundedShapes = normalized.slides
    .flatMap((slide) => slide.elements)
    .filter((element) => element.type === "shape" && element.shape === "roundRect");
  assert.deepEqual(roundedShapes, []);
});

test("keeps chart categories at the 15 pt data-detail floor", () => {
  const spec = deck("clean-business");
  spec.slides[1] = {
    id: "chart",
    role: "table",
    title: "Adoption trend",
    chart: { type: "bar", categories: ["Q1", "Q2", "Q3"], series: [{ name: "Teams", values: [12, 18, 27] }] },
  };
  const normalized = normalizePresentation(presentationOf(spec));
  const labels = normalized.slides[1].elements.filter((element) => element.name === "Chart category");
  assert.equal(labels.length, 3);
  for (const label of labels) for (const paragraph of label.content) for (const run of paragraph.runs) {
    assert.ok(run.style.fontSize >= 15);
  }
  assert.deepEqual(inspectStructure(normalized), []);
});

test("warns instead of silently shrinking dense or internal workflow copy", () => {
  const spec = deck();
  spec.slides[1].title = "A very long title ".repeat(8);
  spec.slides[1].items = ["src-01-report: report.md", "Input slides.md", "A paragraph-like item ".repeat(8)];
  const issues = validateDeckSpec(spec);
  assert.ok(issues.some((issue) => issue.code === "title-density"));
  assert.ok(issues.some((issue) => issue.code === "item-density"));
  assert.ok(issues.some((issue) => issue.code === "internal-metadata-visible"));
});

test("allows a human-readable URL citation without treating it as an internal filename", () => {
  const spec = deck();
  spec.slides[1].items = ["Research note: https://example.com/report.pdf"];
  assert.ok(!validateDeckSpec(spec).some((issue) => issue.code === "internal-metadata-visible"));
});

test("uses the deterministic image-hero composition when an image has no support list", () => {
  const spec = deck("editorial-story");
  spec.slides[1] = {
    id: "image",
    role: "image",
    title: "Evidence, given room to lead",
    message: "One image. One editorial judgment.",
    image: { assetId: "hero", alt: "Abstract editorial composition" },
  };
  assert.deepEqual(validateDeckSpec(spec, new Set(["hero"])), []);
  const presentation = normalizePresentation(presentationOf(spec, () => ({
    source: { type: "url", value: "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='1200' height='675'/%3E" },
    mimeType: "image/svg+xml",
  })));
  assert.ok(presentation.slides[1].elements.some((element) => element.name === "Image hero message"));
  assert.deepEqual(inspectStructure(presentation), []);
});

test("regresses the SaaS Hunt case without visible source footers or weak structure", () => {
  const spec = saasHuntDeck();
  assert.deepEqual(validateDeckSpec(spec), []);
  const normalized = normalizePresentation(presentationOf(spec));
  assert.deepEqual(inspectStructure(normalized), []);
  assert.equal(normalized.slides.length, 6);
  for (const [index, slide] of normalized.slides.entries()) {
    const visible = slide.elements.flatMap((element) => element.type === "text" ? [element.plainText] : []).join("\n");
    assert.doesNotMatch(visible, /Sources:|src-01-slides|slides\.md/);
    assert.match(slide.notes.flatMap((paragraph) => paragraph.runs.map((run) => run.text)).join("\n"), /src-01-slides/);
    if (index > 0) assert.ok(slide.elements.length >= 6, `${slide.id} should use a composed layout`);
  }
});

test("extracts text bytes without filesystem APIs", async () => {
  const source = await extractSource({ name: "brief.md", mimeType: "text/markdown", bytes: new TextEncoder().encode("# Evidence") }, 0);
  assert.equal(source.id, "src-01-brief");
  assert.equal(source.content, "# Evidence");
});

test("extracts PPTX content through the same source parser contract", async () => {
  const input = { name: "reference.pptx", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", bytes: new Uint8Array([1, 2, 3]) };
  const source = await extractSource(input, 0, { pptx: async () => ({ content: "Slide 1\nKeep this message" }) });
  assert.equal(source.content, "Slide 1\nKeep this message");
  assert.deepEqual(source.warnings, []);
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

test("validates session schema and external asset metadata", () => {
  const now = new Date().toISOString();
  const session = { schemaVersion: 2, id: "session-1", revision: 1, createdAt: now, updatedAt: now, deck: deck(), sources: [], assets: [] };
  assert.equal(parseDeckSession(session).id, "session-1");
  assert.throws(() => parseDeckSession({ ...session, schemaVersion: 1 }), /Unsupported/);
  assert.throws(() => parseDeckSession({ ...session, assets: [{ id: "asset", name: "asset.png", mimeType: "image/png", dataUrl: "data:image/png;base64,AA==" }] }), /must not contain inline dataUrl/);
  assert.throws(() => parseDeckSession({ ...session, assets: [{ id: "asset", name: "asset.mp4", mimeType: "video/mp4", byteLength: 1, sha256: "a".repeat(64) }] }), /Unsupported session asset MIME type/);
  assert.throws(() => parseDeckSession({ ...session, assets: [{ id: "asset", name: "asset.png", mimeType: "image/png", byteLength: 0, sha256: "a".repeat(64) }] }), /positive byteLength/);
  assert.throws(() => parseDeckSession({ ...session, assets: [{ id: "asset", name: "asset.png", mimeType: "image/png", byteLength: 1, sha256: "invalid" }] }), /valid SHA-256/);
  const imageDeck = deck();
  imageDeck.slides.push({ id: "missing-image", role: "image", title: "Missing", image: { assetId: "missing", alt: "Missing" } });
  assert.throws(() => parseDeckSession({ ...session, deck: imageDeck }), /references undeclared asset/);
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
