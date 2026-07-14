import test from "node:test";
import assert from "node:assert/strict";
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const skillRoot = path.join(repoRoot, "skills", "pptkit-presentation");
const initScript = path.join(skillRoot, "scripts", "init-project.mjs");
const tsxLoader = path.join(repoRoot, "examples", "dev-app", "node_modules", "tsx", "dist", "loader.mjs");

test("presentation skill requires progressive native interaction and an approval gate", () => {
  const skill = readFileSync(path.join(skillRoot, "SKILL.md"), "utf8");
  const workflow = readFileSync(path.join(skillRoot, "references", "workflow.md"), "utf8");
  const guide = readFileSync(path.join(repoRoot, "docs", "guides", "presentation-skill.md"), "utf8");

  assert.match(skill, /one at a time in this order: purpose and audience, theme, then page count and asset strategy/i);
  assert.match(skill, /request_user_input/i);
  assert.match(skill, /Approve and generate.*Change the plan.*Cancel/is);
  assert.match(skill, /Do not initialize a project, install dependencies, copy sources, or generate a PPTX before/i);
  assert.match(workflow, /# Interaction capability/);
  assert.match(workflow, /Choose one option for <decision>:/);
  assert.match(workflow, /custom in-chat plugin form/i);
  assert.match(workflow, /Treat every outcome except \*\*Approve and generate\*\* as a stop/);
  assert.match(guide, /one at a time/i);
  assert.match(guide, /Approve and generate.*Change the plan.*Cancel/is);
});

function linkDirectory(target, link) {
  mkdirSync(path.dirname(link), { recursive: true });
  symlinkSync(target, link, process.platform === "win32" ? "junction" : "dir");
}

function wireWorkspace(project) {
  linkDirectory(path.join(repoRoot, "packages", "core"), path.join(project, "node_modules", "@pptkit", "core"));
  linkDirectory(path.join(repoRoot, "packages", "layout"), path.join(project, "node_modules", "@pptkit", "layout"));
  linkDirectory(path.join(repoRoot, "packages", "pptx-exporter"), path.join(project, "node_modules", "@pptkit", "pptx-exporter"));
  linkDirectory(realpathSync(path.join(repoRoot, "packages", "pptx-exporter", "node_modules", "fflate")), path.join(project, "node_modules", "fflate"));
}

function runTypeScript(project, entry, args = [], env = process.env) {
  return spawnSync(process.execPath, ["--import", tsxLoader, entry, ...args], { cwd: project, encoding: "utf8", env });
}

function fixtureSpec(themeId) {
  return `import type { DeckSpec } from "./contracts.js";

export const deckSpec: DeckSpec = {
  brief: {
    title: "PPTKit Skill Fixture",
    audience: "Cross-functional teams",
    purpose: "Validate the editable presentation generation workflow",
    language: "en-US",
    slideCountRange: [12, 12],
    themeId: ${JSON.stringify(themeId)},
    imagePolicy: "Local assets only",
    constraints: ["No unsupported warnings"],
    author: "PPTKit",
  },
  slides: [
    { id: "cover", role: "cover", title: "From source material to editable PPTX", subtitle: "A complete, local, testable workflow" },
    { id: "agenda", role: "agenda", title: "The path ahead", items: ["Clarify the goal", "Organize evidence", "Author the slides", "Validate delivery"] },
    { id: "section", role: "section", title: "01 / Why", message: "Clarify the problem before choosing the format." },
    { id: "statement", role: "statement", title: "Core judgment", message: "Structured content and controlled layouts turn generation quality into a verifiable engineering result.", items: ["Local assets", "Native objects", "Explicit diagnostics"], sourceRefs: [{ id: "src-01-report" }] },
    { id: "image", role: "image", title: "Visual evidence", message: "Images fill controlled slots instead of determining the page structure.", items: ["Preserve aspect ratio", "Choose contain or cover", "Provide alt text"], image: { path: "assets/fixture.svg", alt: "Theme fixture", width: 1200, height: 675, fit: "cover" } },
    { id: "kpi", role: "kpi", title: "Key metrics", kpis: [{ value: "+18%", label: "Activation", detail: "Quarter over quarter" }, { value: "72%", label: "Retention" }, { value: "3.4×", label: "Velocity" }] },
    { id: "comparison", role: "comparison", title: "Two paths", comparison: { left: { heading: "Before", items: ["Repeated coordinates", "Implicit defaults", "Hard to inspect"] }, right: { heading: "After", items: ["Semantic plan", "Theme tokens", "Automated report"] } } },
    { id: "process", role: "process", title: "The workflow", steps: ["Intake", "Outline", "Author", "Validate", "Export"] },
    { id: "table", role: "table", title: "Delivery matrix", table: { headers: ["Artifact", "Owner", "Status"], rows: [["Brief", "Agent", "Ready"], ["PPTX", "PPTKit", "Ready"], ["Review", "User", "Open"]] } },
    { id: "bar", role: "table", title: "Quarterly adoption", chart: { type: "bar", categories: ["Q1", "Q2", "Q3", "Q4"], series: [{ name: "Team A", values: [22, 35, 48, 70] }, { name: "Team B", values: [18, 30, 44, 58] }] } },
    { id: "line", role: "table", title: "Delivery speed", chart: { type: "line", categories: ["Jan", "Feb", "Mar", "Apr", "May"], series: [{ name: "Cycle time", values: [8, 7, 5, 4, 3] }] } },
    { id: "closing", role: "closing", title: "Next step", message: "Run the same workflow with real source material.", items: ["Open the PPTX", "Check fonts and wrapping", "Continue editing"] },
  ],
};
`;
}

test("initializer creates an isolated starter and applies theme", () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "pptkit-skill-init-"));
  const project = path.join(root, "deck");
  try {
    const result = spawnSync(process.execPath, [initScript, "--output", project, "--title", "demo-deck", "--theme", "editorial-story", "--no-install"], { encoding: "utf8" });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(readFileSync(path.join(project, "package.json"), "utf8")).name, "demo-deck");
    assert.match(readFileSync(path.join(project, "src", "deck-spec.ts"), "utf8"), /themeId: "editorial-story"/);
    assert.ok(existsSync(path.join(project, "deck-brief.md")));
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test("text source extraction preserves provenance without optional parsers", () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "pptkit-skill-extract-"));
  const project = path.join(root, "deck");
  try {
    assert.equal(spawnSync(process.execPath, [initScript, "--output", project, "--no-install"], { encoding: "utf8" }).status, 0);
    wireWorkspace(project);
    const source = path.join(root, "report.txt");
    writeFileSync(source, "Revenue increased 18%.\nRetention is the next focus.\n");
    const result = runTypeScript(project, "src/extract-sources.ts", [source]);
    assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
    const extracted = JSON.parse(readFileSync(path.join(project, "content", "sources.json"), "utf8"));
    assert.equal(extracted.sources[0].type, "text");
    assert.match(extracted.sources[0].content, /Revenue increased 18%/);
    assert.match(extracted.sources[0].id, /^src-01-report$/);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

for (const themeId of ["clean-business", "swiss-grid", "editorial-story"]) {
  test(`end-to-end fixture builds all roles with ${themeId}`, () => {
    const root = mkdtempSync(path.join(os.tmpdir(), `pptkit-skill-${themeId}-`));
    const project = path.join(root, "deck");
    try {
      const initialized = spawnSync(process.execPath, [initScript, "--output", project, "--theme", themeId, "--no-install"], { encoding: "utf8" });
      assert.equal(initialized.status, 0, initialized.stderr);
      wireWorkspace(project);
      copyFileSync(path.join(skillRoot, "assets", "previews", `${themeId}.svg`), path.join(project, "assets", "fixture.svg"));
      writeFileSync(path.join(project, "src", "deck-spec.ts"), fixtureSpec(themeId));
      const result = runTypeScript(project, "src/build.ts");
      assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
      const report = JSON.parse(readFileSync(path.join(project, "output", "build-report.json"), "utf8"));
      assert.equal(report.slideCount, 12);
      assert.deepEqual(report.diagnostics, []);
      assert.deepEqual(report.exportWarnings, []);
      assert.deepEqual(report.structuralIssues, []);
      assert.equal(report.packageChecks.valid, true);
      assert.equal(report.packageChecks.slideParts, 12);
      assert.ok(existsSync(path.join(project, "output", "deck.pptx")));
      if (themeId === "clean-business") {
        const rendered = runTypeScript(project, "src/render.ts", [], { ...process.env, PATH: "" });
        assert.equal(rendered.status, 0, `${rendered.stdout}\n${rendered.stderr}`);
        const renderReport = JSON.parse(readFileSync(path.join(project, "output", "build-report.json"), "utf8"));
        assert.equal(renderReport.renderStatus, "skipped");
      }
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });
}

test("missing image is reported as an export failure", () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "pptkit-skill-missing-image-"));
  const project = path.join(root, "deck");
  try {
    assert.equal(spawnSync(process.execPath, [initScript, "--output", project, "--no-install"], { encoding: "utf8" }).status, 0);
    wireWorkspace(project);
    writeFileSync(path.join(project, "src", "deck-spec.ts"), `import type { DeckSpec } from "./contracts.js";
export const deckSpec: DeckSpec = {
  brief: { title: "Missing asset", audience: "QA", purpose: "Failure test", language: "en-US", slideCountRange: [1, 1], themeId: "clean-business", imagePolicy: "Local", constraints: [] },
  slides: [{ id: "image", role: "image", title: "Missing image", image: { path: "assets/does-not-exist.png", alt: "Missing fixture", width: 100, height: 100 } }],
};
`);
    const result = runTypeScript(project, "src/build.ts");
    assert.equal(result.status, 1);
    const report = JSON.parse(readFileSync(path.join(project, "output", "build-report.json"), "utf8"));
    assert.ok(report.exportWarnings.length >= 1);
    assert.match(report.exportWarnings.map((warning) => warning.message).join(" "), /does-not-exist|ENOENT|read/i);
    assert.ok(report.structuralIssues.some((issue) => issue.severity === "error"));
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
