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
const testFallbackArgs = [
  "--fallback-reason",
  "unattended-local-output",
  "--browser-check",
  "not-required",
  "--browser-step",
  "user-requirement",
  "--fallback-evidence",
  "Automated test requires isolated local output",
];

test("presentation skill requires progressive native interaction and an approval gate", () => {
  const skill = readFileSync(path.join(skillRoot, "SKILL.md"), "utf8");
  const workflow = readFileSync(path.join(skillRoot, "references", "workflow.md"), "utf8");
  const browserWorkflow = readFileSync(path.join(skillRoot, "references", "browser-workflow.md"), "utf8");
  const nodeWorkflow = readFileSync(path.join(skillRoot, "references", "node-workflow.md"), "utf8");
  const runtimeRouting = readFileSync(path.join(skillRoot, "references", "runtime-routing.md"), "utf8");
  const designSystem = readFileSync(path.join(skillRoot, "references", "design-system.md"), "utf8");
  const quality = readFileSync(path.join(skillRoot, "references", "quality.md"), "utf8");
  const guide = readFileSync(path.join(repoRoot, "docs", "guides", "presentation-skill.md"), "utf8");

  assert.match(skill, /one at a time in this order: purpose and audience, theme, then page count and asset strategy/i);
  assert.match(skill, /request_user_input/i);
  assert.match(skill, /Approve and generate.*Change the plan.*Cancel/is);
  assert.match(skill, /Do not create artifacts, open a preview, install dependencies, or generate PPTX bytes before/i);
  assert.match(skill, /Prefer the browser workflow/i);
  assert.match(skill, /node_repl js/);
  assert.match(skill, /explicitly select the `iab` browser/i);
  assert.match(skill, /Do not infer unavailability from the initial tool list/i);
  assert.match(skill, /runtime-routing\.md/);
  assert.match(skill, /guarded initializer/i);
  assert.match(skill, /Generate & download PPTX/i);
  assert.match(workflow, /# Interaction capability/);
  assert.match(workflow, /Choose one option for <decision>:/);
  assert.match(workflow, /custom in-chat plugin form/i);
  assert.match(workflow, /Treat every outcome except \*\*Approve and generate\*\* as a stop/);
  assert.match(browserWorkflow, /PPTKIT_PREVIEW_URL/);
  assert.match(browserWorkflow, /https:\/\/openhacking\.github\.io\/pptkit\//);
  assert.match(browserWorkflow, /explicitly supplied[\s\S]*PPTKIT_PREVIEW_URL[\s\S]*official PPTKit preview application/i);
  assert.match(browserWorkflow, /resolved URL is unreachable or incompatible/i);
  assert.match(browserWorkflow, /Do not give up solely because the initial tool list omits browser controls/i);
  assert.match(browserWorkflow, /successful open or focus operation as proof/i);
  assert.match(browserWorkflow, /marked `deliverable`/i);
  assert.match(browserWorkflow, /Do not download automatically/);
  assert.match(browserWorkflow, /explicitly asks the agent to trigger the export\/download/);
  assert.match(browserWorkflow, /IndexedDB/);
  assert.match(nodeWorkflow, /State the fallback reason/i);
  assert.match(nodeWorkflow, /runtime-decision\.json/i);
  assert.match(runtimeRouting, /Do not read or execute `node-workflow\.md` while the decision is unresolved/i);
  assert.match(runtimeRouting, /Do not claim a browser failure without a tool result/i);
  assert.match(designSystem, /sourceRefs.*provenance metadata/is);
  assert.match(designSystem, /46–60 pt/);
  assert.match(designSystem, /hero.*split.*ledger.*timeline/is);
  assert.match(quality, /visible internal source IDs/i);
  assert.match(guide, /customers do not need to configure a preview URL/i);
  assert.match(guide, /does not treat an abbreviated initial tool list as evidence that no browser exists/i);
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
  linkDirectory(path.join(repoRoot, "packages", "presentation-workflow"), path.join(project, "node_modules", "@pptkit", "presentation-workflow"));
  linkDirectory(realpathSync(path.join(repoRoot, "packages", "pptx-exporter", "node_modules", "fflate")), path.join(project, "node_modules", "fflate"));
  linkDirectory(realpathSync(path.join(repoRoot, "examples", "presentation-preview", "node_modules", "mammoth")), path.join(project, "node_modules", "mammoth"));
  linkDirectory(realpathSync(path.join(repoRoot, "examples", "presentation-preview", "node_modules", "pdfjs-dist")), path.join(project, "node_modules", "pdfjs-dist"));
  linkDirectory(realpathSync(path.join(repoRoot, "examples", "presentation-preview", "node_modules", "xlsx")), path.join(project, "node_modules", "xlsx"));
  linkDirectory(realpathSync(path.join(repoRoot, "node_modules", "@types", "node")), path.join(project, "node_modules", "@types", "node"));
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
    { id: "image", role: "image", title: "Visual evidence", message: "Images fill controlled slots instead of determining the page structure.", items: ["Preserve aspect ratio", "Choose contain or cover", "Provide alt text"], image: { assetId: "fixture.svg", alt: "Theme fixture", width: 1200, height: 675, fit: "cover" } },
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
    const result = spawnSync(process.execPath, [initScript, "--output", project, "--title", "demo-deck", "--theme", "editorial-story", "--no-install", ...testFallbackArgs], { encoding: "utf8" });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(readFileSync(path.join(project, "package.json"), "utf8")).name, "demo-deck");
    assert.match(readFileSync(path.join(project, "src", "deck-spec.ts"), "utf8"), /themeId: "editorial-story"/);
    assert.ok(existsSync(path.join(project, "deck-brief.md")));
    assert.deepEqual(JSON.parse(readFileSync(path.join(project, "runtime-decision.json"), "utf8")), {
      schemaVersion: 1,
      selectedRuntime: "node",
      reason: "unattended-local-output",
      resolvedPreviewUrl: "https://openhacking.github.io/pptkit/",
      browserCheck: {
        status: "not-required",
        step: "user-requirement",
        evidence: "Automated test requires isolated local output",
      },
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test("initializer rejects missing or contradictory runtime-routing evidence before creating output", () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "pptkit-skill-routing-"));
  try {
    const missingOutput = path.join(root, "missing");
    const missing = spawnSync(process.execPath, [initScript, "--output", missingOutput, "--no-install"], { encoding: "utf8" });
    assert.equal(missing.status, 2);
    assert.match(missing.stderr, /fallback-reason/i);
    assert.equal(existsSync(missingOutput), false);

    const contradictoryOutput = path.join(root, "contradictory");
    const contradictory = spawnSync(
      process.execPath,
      [
        initScript,
        "--output",
        contradictoryOutput,
        "--no-install",
        "--fallback-reason",
        "browser-setup-failed",
        "--browser-check",
        "not-required",
        "--browser-step",
        "setup",
        "--fallback-evidence",
        "No browser control was visible initially",
      ],
      { encoding: "utf8" },
    );
    assert.equal(contradictory.status, 2);
    assert.match(contradictory.stderr, /requires --browser-check failed/i);
    assert.equal(existsSync(contradictoryOutput), false);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test("text source extraction preserves provenance without optional parsers", () => {
  const root = mkdtempSync(path.join(os.tmpdir(), "pptkit-skill-extract-"));
  const project = path.join(root, "deck");
  try {
    assert.equal(spawnSync(process.execPath, [initScript, "--output", project, "--no-install", ...testFallbackArgs], { encoding: "utf8" }).status, 0);
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
      const initialized = spawnSync(process.execPath, [initScript, "--output", project, "--theme", themeId, "--no-install", ...testFallbackArgs], { encoding: "utf8" });
      assert.equal(initialized.status, 0, initialized.stderr);
      wireWorkspace(project);
      copyFileSync(path.join(skillRoot, "assets", "previews", `${themeId}.svg`), path.join(project, "assets", "fixture.svg"));
      writeFileSync(path.join(project, "src", "deck-spec.ts"), fixtureSpec(themeId));
      if (themeId === "clean-business") {
        const typecheck = spawnSync(process.execPath, [path.join(repoRoot, "node_modules", "typescript", "bin", "tsc"), "--noEmit"], { cwd: project, encoding: "utf8" });
        assert.equal(typecheck.status, 0, `${typecheck.stdout}\n${typecheck.stderr}`);
      }
      const result = runTypeScript(project, "src/build.ts");
      const reportText = existsSync(path.join(project, "output", "build-report.json")) ? readFileSync(path.join(project, "output", "build-report.json"), "utf8") : "no report";
      assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}\n${reportText}`);
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
    assert.equal(spawnSync(process.execPath, [initScript, "--output", project, "--no-install", ...testFallbackArgs], { encoding: "utf8" }).status, 0);
    wireWorkspace(project);
    writeFileSync(path.join(project, "src", "deck-spec.ts"), `import type { DeckSpec } from "./contracts.js";
export const deckSpec: DeckSpec = {
  brief: { title: "Missing asset", audience: "QA", purpose: "Failure test", language: "en-US", slideCountRange: [1, 1], themeId: "clean-business", imagePolicy: "Local", constraints: [] },
  slides: [{ id: "image", role: "image", title: "Missing image", image: { assetId: "does-not-exist.png", alt: "Missing fixture", width: 100, height: 100 } }],
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
