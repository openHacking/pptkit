import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

const verify = file => execFileSync("python3", ["scripts/research/verify-tab-stops.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
test("native shape tab observations and SVG comparisons reject altered evidence", async t => {
  const native = file => execFileSync("python3", ["scripts/research/verify-native-shape-tabs.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const browser = file => execFileSync("node", ["scripts/research/verify-svg-shape-tab-powerpoint.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.equal(JSON.parse(native()).samples, 12); assert.equal(JSON.parse(browser()).comparisons, 12);
  const dir = await mkdtemp(path.join(tmpdir(), 'pptkit-shape-tab-evidence-')); t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, 'bad.json');
  for (const [path, check, mutations] of [
    ['docs/testing/tab-shape-powerpoint-evidence.json', native, [e => { e.pdfSha256 = '0'.repeat(64); }, e => { e.slides[1].rows[1].word[0]++; }, e => { e.rounds[2].pptxSha256 = '0'.repeat(64); }]],
    ['docs/testing/tab-svg-shape-browser-powerpoint-evidence.json', browser, [e => { e.items[0].svgFill = '#FFFFFF'; }, e => { e.items[2].svgLeft++; }, e => { e.bundles[2].sha256 = '0'.repeat(64); }, e => { e.sourceSha256 = '0'.repeat(64); }]],
  ]) {
    const original = JSON.parse(await readFile(path, 'utf8'));
    for (const mutate of mutations) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
  }
});
test("multiword SVG comparison binds the fitting native sample and rejects altered character evidence", async t => {
  const check = file => execFileSync("node", ["scripts/research/verify-svg-multiword-powerpoint.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const result = JSON.parse(check()); assert.equal(result.characters, 10); assert.ok(result.maxDifference < 0.5);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-svg-multiword-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-svg-multiword-browser-powerpoint-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.chars[1].pdfX++; }, e => { e.chars[2].svgX++; }, e => { e.chars[3].char = "X"; }, e => { e.fonts[0].sha256 = "0".repeat(64); }, e => { e.bundles[2].sha256 = "0".repeat(64); }, e => { e.parts[1] = "I WIDE"; }, e => { e.sourceSha256 = "0".repeat(64); }]) {
    const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file));
  }
});
test("tab-stop observation is bound to its input packages and measured PDFs", async t => {
  const result = JSON.parse(verify()); assert.equal(result.verified, true); assert.equal(result.variants, 3); assert.equal(result.pagesPerVariant, 4);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tabs-evidence-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-stops-libreoffice-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.variants[0].pages[0][0].bbox[0] += 1; }, e => { e.variants[1].sourceSha256 = "0".repeat(64); }, e => { e.variants[2].pdfSha256 = "0".repeat(64); }, e => { e.runnerSha256 = "0".repeat(64); }, e => { e.variants[0].source = "../../outside.pptx"; }]) {
    const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => verify(file));
  }
});

test("indent isolation evidence detects changed source and PDF observations", async t => {
  const check = file => execFileSync("python3", ["scripts/research/verify-tab-stop-indent.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(check()), { verified: true, variants: 6, pagesPerVariant: 4 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tabs-indent-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-stop-indent-libreoffice-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.variants[1].pages[0][0].bbox[0] += 1; }, e => { e.variants[2].sourceSha256 = "0".repeat(64); }, e => { e.variants[4].pdfSha256 = "0".repeat(64); }, e => { e.runnerSha256 = "0".repeat(64); }]) {
    const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file));
  }
});

test("paragraph-coordinate evidence checks ODF controls and PDF geometry", async t => {
  const check = file => execFileSync("python3", ["scripts/research/verify-tab-stop-paragraph.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(check()), { verified: true, cases: 4, pagesPerCase: 4 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tabs-paragraph-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-stop-paragraph-libreoffice-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.cases[1].margin = 35; }, e => { e.cases[2].pages[0].samples[0].bbox[0] += 1; }, e => { e.cases[3].sourceSha256 = "0".repeat(64); }, e => { e.runnerSha256 = "0".repeat(64); }]) {
    const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file));
  }
});

test("soft-wrap evidence checks controls, package hashes and measured PDF geometry", async t => {
  const check = file => execFileSync("python3", ["scripts/research/verify-tab-stop-wrap.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(check()), { verified: true, cases: 4, pagesPerCase: 4 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tabs-wrap-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-stop-wrap-libreoffice-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.cases[0].width++; }, e => { e.cases[1].pages[0][0].bbox[0]++; }, e => { e.cases[2].sourceSha256 = "0".repeat(64); }, e => { e.cases[3].pdfSha256 = "0".repeat(64); }, e => { e.runnerSha256 = "0".repeat(64); }]) {
    const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file));
  }
});

test("soft-wrap cross-application evidence is bound to the observations and fonts", async t => {
  const check = file => execFileSync("node", ["scripts/research/verify-tab-wrap-cross-app.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const result = JSON.parse(check()); assert.equal(result.verified, true); assert.equal(result.comparisons, 36); assert.ok(result.maxRightDifference < 0.5);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tabs-wrap-cross-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-stop-wrap-browser-libreoffice-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.items[0].browserRight++; }, e => { e.items[0].browserWrapped = !e.items[0].browserWrapped; }, e => { e.fonts[0].sha256 = "0".repeat(64); }, e => { e.sourceSha256 = "0".repeat(64); }, e => { e.runnerSha256 = "0".repeat(64); }]) {
    const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file));
  }
});

test("multiword controls and cross-application evidence reject altered observations", async t => {
  const native = file => execFileSync("python3", ["scripts/research/verify-tab-stop-multiword.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const cross = file => execFileSync("node", ["scripts/research/verify-tab-multiword-cross-app.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(native()), { verified: true, cases: 3, samples: 27 });
  assert.equal(JSON.parse(cross()).comparisons, 27);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tabs-multi-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, "bad.json");
  for (const [source, check, mutations] of [
    ["docs/testing/tab-stop-multiword-libreoffice-evidence.json", native, [e => { e.cases[0].pages[0][0][0].bbox[0]++; }, e => { e.cases[0].sourceSha256 = "0".repeat(64); }, e => { e.priorSha256 = "0".repeat(64); }]],
    ["docs/testing/tab-stop-multiword-browser-libreoffice-evidence.json", cross, [e => { e.items[0].browserRight++; }, e => { e.items[0].browserLines++; }, e => { e.fonts[0].sha256 = "0".repeat(64); }]],
  ]) {
    const original = JSON.parse(await readFile(source, "utf8"));
    for (const mutate of mutations) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
  }
});

test("PowerPoint first-open and three saved XML observations reject tampering", async t => {
  const check = file => execFileSync("python3", ["scripts/research/verify-native-tab-stops.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(check()), { verified: true, slides: 4, samples: 12 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tabs-powerpoint-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-stop-powerpoint-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.slides[0].rows[0].anchor++; }, e => { e.rounds[2].pptxSha256 = "0".repeat(64); }, e => { e.pdfSha256 = "0".repeat(64); }]) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
});

test("PowerPoint narrow-frame observation binds source and per-character PDF geometry", async t => {
  const check = file => execFileSync("python3", ["scripts/research/verify-native-tab-wrap.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(check()), { verified: true, slides: 4, samples: 12 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tabs-powerpoint-wrap-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-stop-wrap-powerpoint-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.observations[0][1][1].chars[0].bbox[0]++; }, e => { e.pdfSha256 = "0".repeat(64); }, e => { e.sourceSha256 = "0".repeat(64); }, e => { e.generatorSha256 = "0".repeat(64); }]) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
});

test("PowerPoint width matrix and browser comparison reject changed evidence", async t => {
  const native = file => execFileSync("python3", ["scripts/research/verify-native-tab-wrap-matrix.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const cross = file => execFileSync("node", ["scripts/research/verify-tab-wrap-matrix-cross-app.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(native()), { verified: true, slides: 16, samples: 48 });
  const result = JSON.parse(cross()); assert.equal(result.comparisons, 36); assert.equal(result.chars, 316); assert.ok(result.maxRightDifference < 1); assert.ok(result.maxCharDifference < 1);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tabs-width-matrix-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, "bad.json");
  for (const [source, check, mutations] of [
    ["docs/testing/tab-stop-wrap-matrix-powerpoint-evidence.json", native, [e => { e.observations[0].rows[1][1].chars[0].bbox[0]++; }, e => { e.sourceSha256 = "0".repeat(64); }, e => { e.pdfSha256 = "0".repeat(64); }]],
    ["docs/testing/tab-stop-wrap-matrix-browser-powerpoint-evidence.json", cross, [e => { e.items[0].browserRight++; }, e => { e.items[0].browserLines++; }, e => { e.items[0].chars[0].browserX++; }, e => { e.items[0].chars[0].browserLine++; }, e => { e.fonts[0].sha256 = "0".repeat(64); }]],
  ]) {
    const original = JSON.parse(await readFile(source, "utf8"));
    for (const mutate of mutations) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
  }
});

test("tab-position LibreOffice control checks the only permitted XML change and PDF geometry", async t => {
  const check = file => execFileSync("python3", ["scripts/research/verify-tab-position-libreoffice.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(check()), { verified: true, slides: 12, samples: 36 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tab-position-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-position-libreoffice-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.pages[0].rows[0][0].chars[0].bbox[0]++; }, e => { e.variantSha256 = "0".repeat(64); }, e => { e.pdfSha256 = "0".repeat(64); }, e => { e.sourceSha256 = "0".repeat(64); }]) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
});

test("tab-position PowerPoint and browser evidence reject changed inputs and character records", async t => {
  const native = file => execFileSync("python3", ["scripts/research/verify-native-tab-position.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const cross = file => execFileSync("node", ["scripts/research/verify-tab-position-cross-app.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(native()), { verified: true, slides: 12, samples: 36 });
  const result = JSON.parse(cross()); assert.equal(result.comparisons, 36); assert.equal(result.chars, 316); assert.ok(result.maxCharDifference < 1);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tab-native-position-")); t.after(() => rm(dir, { recursive: true, force: true })); const file = path.join(dir, "bad.json");
  for (const [source, check, mutations] of [
    ["docs/testing/tab-position-powerpoint-evidence.json", native, [e => { e.pages[0].rows[0][0].chars[0].bbox[0]++; }, e => { e.sourceSha256 = "0".repeat(64); }, e => { e.pdfSha256 = "0".repeat(64); }]],
    ["docs/testing/tab-position-browser-powerpoint-evidence.json", cross, [e => { e.items[0].chars[0].browserX++; }, e => { e.items[0].chars[0].browserLine++; }, e => { e.fonts[0].sha256 = "0".repeat(64); }]],
  ]) { const original = JSON.parse(await readFile(source, "utf8")); for (const mutate of mutations) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); } }
});

test("equal-authored-body cross-application comparison detects altered line groups", async t => {
  const check = file => execFileSync("python3", ["scripts/research/verify-tab-shared-source-apps.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(check()), { verified: true, samples: 6, differentLineGroupings: 3 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tab-apps-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-shared-source-apps-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.cases[0].libreofficeLines[0] += "X"; }, e => { e.powerpointEvidenceSha256 = "0".repeat(64); }, e => { e.libreofficeEvidenceSha256 = "0".repeat(64); }]) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
});

test("PowerPoint three-item phrase evidence binds source, PDF and browser character geometry", async t => {
  const native = file => execFileSync("python3", ["scripts/research/verify-native-tab-phrases.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const cross = file => execFileSync("node", ["scripts/research/verify-tab-phrases-cross-app.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(native()), { verified: true, slides: 8, samples: 24 });
  const result = JSON.parse(cross()); assert.equal(result.comparisons, 24); assert.equal(result.chars, 304); assert.ok(result.maxCharDifference < 1);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tab-phrases-")); t.after(() => rm(dir, { recursive: true, force: true })); const file = path.join(dir, "bad.json");
  for (const [source, check, mutations] of [
    ["docs/testing/tab-phrases-powerpoint-evidence.json", native, [e => { e.observations[0].rows[0][0].chars[0].bbox[0]++; }, e => { e.sourceSha256 = "0".repeat(64); }, e => { e.pdfSha256 = "0".repeat(64); }]],
    ["docs/testing/tab-phrases-browser-powerpoint-evidence.json", cross, [e => { e.items[0].chars[0].browserX++; }, e => { e.items[0].chars[0].browserLine++; }, e => { e.fonts[0].sha256 = "0".repeat(64); }]],
  ]) { const original = JSON.parse(await readFile(source, "utf8")); for (const mutate of mutations) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); } }
});

test("native SVG tab anchors compare against PowerPoint and reject altered evidence", async t => {
  const check = file => execFileSync("node", ["scripts/research/verify-svg-tab-powerpoint.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const result = JSON.parse(check()); assert.equal(result.verified, true); assert.equal(result.comparisons, 12); assert.ok(result.maxDifference < 0.5);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tab-svg-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-svg-browser-powerpoint-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.items[0].svgLeft++; }, e => { e.items[0].pdfRight++; }, e => { e.bundles[2].sha256 = "0".repeat(64); }, e => { e.fonts[0].sha256 = "0".repeat(64); }, e => { e.sourceSha256 = "0".repeat(64); }]) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
});

test("mixed-run SVG anchors retain native sample styles and reject altered evidence", async t => {
  const check = file => execFileSync("node", ["scripts/research/verify-svg-mixed-tab-powerpoint.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const result = JSON.parse(check()); assert.equal(result.verified, true); assert.equal(result.comparisons, 12); assert.ok(result.maxDifference < 0.5);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tab-svg-mixed-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-svg-mixed-browser-powerpoint-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.items[0].svgPrefixRight++; }, e => { e.items[0].svgText = "PI"; }, e => { e.items[0].svgSuffixWeights[0] = "700"; }, e => { e.bundles[2].sha256 = "0".repeat(64); }, e => { e.fonts[0].sha256 = "0".repeat(64); }, e => { e.sourceSha256 = "0".repeat(64); }]) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
});

test("cross-application tab evidence binds source geometry, fonts and measured arithmetic", async t => {
  const check = file => execFileSync("node", ["scripts/research/verify-tab-cross-app.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const verified = JSON.parse(check()); assert.equal(verified.verified, true); assert.equal(verified.comparisons, 48); assert.ok(verified.maxRightDifference < 0.5);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tabs-cross-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/tab-stop-browser-libreoffice-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const mutate of [e => { e.sourceSha256 = "0".repeat(64); }, e => { e.fonts[0].sha256 = "0".repeat(64); }, e => { e.items[0].pdfRight++; }, e => { e.items[0].browserRight++; }, e => { e.selectedFonts[0][0].postScriptName = "ArialMT"; }, e => { e.runnerSha256 = "0".repeat(64); }]) {
    const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file));
  }
});

test("list/tab evidence rejects changes to source geometry and cross-application records", async t => {
  const checkNative = file => execFileSync("python3", ["scripts/research/verify-tab-lists.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const checkCross = file => execFileSync("node", ["scripts/research/verify-tab-list-cross-app.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(checkNative()), { verified: true, cases: 36 });
  const cross = JSON.parse(checkCross()); assert.equal(cross.verified, true); assert.equal(cross.comparisons, 36); assert.ok(cross.maxRightDifference < 0.5);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-tabs-list-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, "bad.json");
  for (const [source, check, mutations] of [
    ["docs/testing/tab-list-libreoffice-evidence.json", checkNative, [e => { e.observations[0].marker[0].bbox[0]++; }, e => { e.sourceSha256 = "0".repeat(64); }, e => { e.observations[0].indent++; }]],
    ["docs/testing/tab-list-browser-libreoffice-evidence.json", checkCross, [e => { e.items[0].browserRight++; }, e => { e.font.sha256 = "0".repeat(64); }, e => { e.selectedFonts[0].postScriptName = "ArialMT"; }]],
  ]) {
    const original = JSON.parse(await readFile(source, "utf8"));
    for (const mutate of mutations) { const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
  }
});

test("third-reopen shape PDF evidence rejects altered coordinates and bindings", async t => {
  const check = file => execFileSync("python3", ["scripts/research/verify-shape-tabs-reopened.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(check()), { verified: true, pages: 4, samples: 12 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-shape-reopened-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, "bad.json"), original = JSON.parse(await readFile("docs/testing/tab-shape-reopened-powerpoint-evidence.json", "utf8"));
  for (const mutate of [e => { e.firstEvidenceSha256 = "0".repeat(64); }, e => { e.thirdSavedPptxSha256 = "0".repeat(64); }, e => { e.reopenedPdfSha256 = "0".repeat(64); }, e => { e.pages[3].rows[2].anchor++; }, e => { e.pages[1].fills[0][0][0]++; }]) {
    const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file));
  }
});

test("PowerPoint shape-text edit evidence binds its saved package, print and observation", async t => {
  const check = file => execFileSync("python3", ["scripts/research/verify-shape-tab-native-edit.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(check()), { verified: true, slides: 4, samples: 12, edited: 1 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-shape-edit-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, "bad.json"), original = JSON.parse(await readFile("docs/testing/tab-shape-native-edit-powerpoint-evidence.json", "utf8"));
  for (const mutate of [e => { e.priorEvidenceSha256 = "0".repeat(64); }, e => { e.editedPptxSha256 = "0".repeat(64); }, e => { e.editedPdfSha256 = "0".repeat(64); }, e => { e.reopenedAxSha256 = "0".repeat(64); }, e => { e.changedWord.after[0]++; }]) {
    const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file));
  }
});

test("PowerPoint mixed-size shape observation binds baseline measurements and inputs", async t => {
  const check = file => execFileSync("python3", ["scripts/research/verify-shape-mixed-size-powerpoint.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const result = JSON.parse(check()); assert.equal(result.samples, 12); assert.ok(result.baselineDifferenceRangePt[0] > -1 && result.baselineDifferenceRangePt[1] < 1);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-shape-mixed-size-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, "bad.json"), original = JSON.parse(await readFile("docs/testing/tab-shape-mixed-size-powerpoint-evidence.json", "utf8"));
  for (const mutate of [e => { e.sourceSha256 = "0".repeat(64); }, e => { e.pdfSha256 = "0".repeat(64); }, e => { e.svgSourceSha256 = "0".repeat(64); }, e => { e.pages[1].rows[1].baselineDifferencePt++; }]) {
    const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file));
  }
});

test("mixed-size shape browser and reopen evidence reject changed measurements and inputs", async t => {
  const browser = file => execFileSync("node", ["scripts/research/verify-svg-shape-mixed-size-powerpoint.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const reopen = file => execFileSync("python3", ["scripts/research/verify-shape-mixed-size-reopen.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const comparison = JSON.parse(browser()); assert.equal(comparison.characters, 44); assert.ok(comparison.maxHorizontal < .60 && comparison.maxBaseline < .85);
  assert.deepEqual(JSON.parse(reopen()), { verified: true, pages: 4, samples: 12, reopenRounds: 1 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-mixed-shape-evidence-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, "bad.json");
  for (const [source, check, changes] of [
    ["docs/testing/tab-svg-shape-mixed-size-powerpoint-evidence.json", browser, [e => { e.sourceSha256 = "0".repeat(64); }, e => { e.bundles[2].sha256 = "0".repeat(64); }, e => { e.items[0].chars[0].svg.left++; }, e => { e.maxBaseline++; }]],
    ["docs/testing/tab-shape-mixed-size-reopen-powerpoint-evidence.json", reopen, [e => { e.savedPptxSha256 = "0".repeat(64); }, e => { e.reopenedPdfSha256 = "0".repeat(64); }, e => { e.firstEvidenceSha256 = "0".repeat(64); }]],
  ]) {
    const original = JSON.parse(await readFile(source, "utf8"));
    for (const change of changes) { const changed = structuredClone(original); change(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
  }
});

test("three mixed-size shape reopen rounds bind all packages and printed endpoints", async t => {
  const check = file => execFileSync("python3", ["scripts/research/verify-shape-mixed-size-three-reopen.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(check()), { verified: true, pages: 4, samples: 12, reopenRounds: 3 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-mixed-shape-three-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, "bad.json"), original = JSON.parse(await readFile("docs/testing/tab-shape-mixed-size-three-reopen-powerpoint-evidence.json", "utf8"));
  for (const mutate of [e => { e.rounds[1].pptxSha256 = "0".repeat(64); }, e => { e.rounds[2].axSha256 = "0".repeat(64); }, e => { e.reopenedPdfSha256[1] = "0".repeat(64); }]) {
    const changed = structuredClone(original); mutate(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file));
  }
});

test("Chinese shape tab PowerPoint and browser evidence bind glyphs, fonts and measurements", async t => {
  const native = file => execFileSync("python3", ["scripts/research/verify-cjk-shape-tabs-powerpoint.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  const browser = file => execFileSync("node", ["scripts/research/verify-svg-cjk-shape-powerpoint.mjs", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
  assert.deepEqual(JSON.parse(native()), { verified: true, pages: 4, samples: 12, characters: 32 });
  const summary = JSON.parse(browser()); assert.equal(summary.characters, 32); assert.ok(summary.maxHorizontal < .34 && summary.maxBaseline < .25);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-cjk-shape-")); t.after(() => rm(dir, { recursive: true, force: true })); const file = path.join(dir, "bad.json");
  for (const [path, check, changes] of [
    ["docs/testing/tab-cjk-shape-powerpoint-evidence.json", native, [e => { e.pdfSha256 = "0".repeat(64); }, e => { e.pages[1].rows[1].characters[0].left++; }, e => { e.sourceSha256 = "0".repeat(64); }]],
    ["docs/testing/tab-svg-cjk-shape-powerpoint-evidence.json", browser, [e => { e.font.sha256 = "0".repeat(64); }, e => { e.sourceSha256 = "0".repeat(64); }, e => { e.items[2].chars[0].svg.left++; }, e => { e.maxBaseline++; }]],
  ]) {
    const original = JSON.parse(await readFile(path, "utf8")); for (const change of changes) { const changed = structuredClone(original); change(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => check(file)); }
  }
});
