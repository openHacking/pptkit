import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { assertTargetVersionsAvailable, bumpVersion, getPublishablePackages, parseArgs, performRelease, validateVersion } from "./release.mjs";
import { assertReleaseNotes, compareVersions, releaseTag } from "./release-policy.mjs";

test("bumpVersion supports patch, minor, and major releases", () => {
  assert.equal(bumpVersion("1.2.3", "patch"), "1.2.4");
  assert.equal(bumpVersion("1.2.3", "minor"), "1.3.0");
  assert.equal(bumpVersion("1.2.3", "major"), "2.0.0");
});

test("version validation accepts numeric and preview.N versions without coercion", () => {
  assert.equal(validateVersion("0.1.0"), true);
  assert.equal(validateVersion("1.2"), false);
  assert.equal(validateVersion("v1.2.3"), false);
  assert.equal(validateVersion("0.2.0-preview.12"), true);
  for (const value of ["01.2.3", "1.02.3", "1.2.03", "1.2.3-preview.01", "1.2.3-preview", "1.2.3-beta.1", "1.2.3+build", "9007199254740992.0.0", "1.2.3\n", undefined, 123]) assert.equal(validateVersion(value), false, String(value));
});

test("package discovery selects only public packages", async (t) => {
  const root = await mkdtemp(path.join(os.tmpdir(), "pptkit-release-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  await mkdir(path.join(root, "public"));
  await mkdir(path.join(root, "private"));
  await writeFile(path.join(root, "public", "package.json"), JSON.stringify({ name: "@pptkit/public", version: "0.1.0", private: false }));
  await writeFile(path.join(root, "private", "package.json"), JSON.stringify({ name: "@pptkit/private", version: "0.1.0", private: true }));

  const packages = getPublishablePackages(root);
  assert.deepEqual(packages.map((pkg) => pkg.name), ["@pptkit/public"]);
});

test("version ordering and preview promotion cannot regress", () => {
  assert.equal(bumpVersion("0.1.10", "preview"), "0.1.11-preview.0");
  assert.equal(bumpVersion("0.1.11-preview.0", "preview"), "0.1.11-preview.1");
  assert.equal(bumpVersion("0.1.11-preview.1", "patch"), "0.1.11");
  assert.equal(bumpVersion("0.1.11-preview.1", "minor"), "0.2.0");
  assert.equal(compareVersions("0.2.0-preview.9", "0.2.0-preview.10"), -1);
  assert.equal(compareVersions("0.2.0", "0.2.0-preview.99"), 1);
  assert.equal(compareVersions("0.2.0-preview.1", "0.1.99"), 1);
  assert.equal(compareVersions("0.2.0", "0.2.0"), 0);
  assert.throws(() => bumpVersion("9007199254740991.0.0", "major"), /range/);
  assert.throws(() => bumpVersion("0.1.0", "unknown"), /release type/);
});

test("release tags prevent provisional APIs from advancing latest", () => {
  assert.equal(releaseTag("0.1.11"), "next");
  assert.equal(releaseTag("1.0.0-preview.0"), "next");
  assert.equal(releaseTag("1.0.0"), "latest");
  assert.equal(releaseTag("1.0.0", "next"), "next");
  assert.throws(() => releaseTag("0.1.11", "latest"), /next tag/);
  assert.throws(() => releaseTag("1.0.0-preview.0", "latest"), /next tag/);
  assert.throws(() => releaseTag("1.0.0", "nightly"), /tag/);
});

test("CLI options reject malformed, repeated, and missing arguments", () => {
  assert.deepEqual(parseArgs(["--dry-run", "--version", "0.2.0-preview.0", "--tag", "next"]), { dryRun: true, version: "0.2.0-preview.0", tag: "next" });
  for (const args of [["--version"], ["--version", "--dry-run"], ["--unknown"], ["--dry-run", "--dry-run"], ["--version", "latest"], ["--tag", "nightly"]]) assert.throws(() => parseArgs(args));
});

test("release notes require the exact release, with Unreleased allowed only for inspection", () => {
  const notes = "# Changelog\n\n## [Unreleased]\nWork in progress.\n\n## [0.2.0]\nCompatibility and verification evidence.\n";
  assert.doesNotThrow(() => assertReleaseNotes(notes, "0.2.0", false));
  assert.doesNotThrow(() => assertReleaseNotes(notes, "0.3.0", true));
  assert.throws(() => assertReleaseNotes(notes, "0.3.0", false), /0.3.0/);
  assert.throws(() => assertReleaseNotes("## [0.2.0]\n\n## [0.1.0]\nOlder release\n", "0.2.0", false), /nonempty/);
});

test("registry absence requires structured E404 rather than arbitrary error text", () => {
  const packages = [{ name: "@pptkit/core" }];
  assert.doesNotThrow(() => assertTargetVersionsAvailable(packages, "0.2.0", () => ({ status: 1, stdout: '{"error":{"code":"E404"}}' })));
  assert.throws(() => assertTargetVersionsAvailable(packages, "0.2.0", () => ({ status: 0, stdout: '"0.2.0"' })), /already exists/);
  for (const result of [
    { status: 1, stdout: "", stderr: "E404 not really registry evidence" },
    { status: 1, stdout: '{"error":{"code":"E401"}}' },
    { status: 1, stdout: '{"error":{"code":"E404"}}', signal: "SIGTERM" },
    { status: null, error: new Error("timeout") },
  ]) assert.throws(() => assertTargetVersionsAvailable(packages, "0.2.0", () => result), /Could not verify/);
});

async function releaseHarness(t) {
  const directory = await mkdtemp(path.join(os.tmpdir(), "pptkit-release-flow-"));
  t.after(() => rm(directory, { recursive: true, force: true }));
  for (const name of ["core", "layout"]) {
    await mkdir(path.join(directory, name));
    await writeFile(path.join(directory, name, "package.json"), JSON.stringify({ name: `@pptkit/${name}`, private: false, version: "0.1.10", description: "Retain unrelated fields" }));
  }
  const packages = getPublishablePackages(directory);
  const versions = async () => Promise.all(packages.map(async (pkg) => JSON.parse(await readFile(pkg.packageJsonPath, "utf8")).version));
  const events = [];
  const commands = {
    log() {},
    checkRegistry() { events.push("registry"); },
    run(_command, args) { events.push(args.includes("publish") ? "publish" : args[0]); },
    async confirm() { events.push("confirm"); return true; },
  };
  return { options: { packages, nextVersion: "0.2.0", changelog: "## [0.2.0]\nDocumented changes and checks.\n" }, commands, events, versions };
}

test("release dry run validates current artifacts without changing manifests or publishing", async (t) => {
  const harness = await releaseHarness(t);
  const result = await performRelease({ ...harness.options, dryRun: true }, harness.commands);
  assert.equal(result.status, "dry-run");
  assert.deepEqual(harness.events, ["registry", "build", "typecheck", "lint", "test"]);
  assert.deepEqual(await harness.versions(), ["0.1.10", "0.1.10"]);
});

test("release cancellation follows validation and leaves manifests intact", async (t) => {
  const harness = await releaseHarness(t);
  harness.commands.confirm = async () => false;
  assert.equal((await performRelease(harness.options, harness.commands)).status, "cancelled");
  assert.deepEqual(await harness.versions(), ["0.1.10", "0.1.10"]);
  assert.ok(!harness.events.includes("publish"));
});

test("invalid policy, version, changelog and mismatched package versions fail before external commands", async (t) => {
  for (const override of [{ nextVersion: "0.1.9" }, { nextVersion: "0.1.10" }, { tag: "latest" }, { changelog: "## [Unreleased]\nNotes\n" }]) {
    const harness = await releaseHarness(t);
    await assert.rejects(performRelease({ ...harness.options, ...override }, harness.commands));
    assert.deepEqual(harness.events, []);
    assert.deepEqual(await harness.versions(), ["0.1.10", "0.1.10"]);
  }
  const harness = await releaseHarness(t);
  harness.options.packages[1].version = "0.1.11";
  await assert.rejects(performRelease(harness.options, harness.commands), /unified version/);
  assert.deepEqual(harness.events, []);
});

test("failed validation cannot update manifests, request confirmation or publish", async (t) => {
  const harness = await releaseHarness(t);
  harness.commands.run = (_command, args) => { if (args[0] === "lint") throw new Error("lint failed"); };
  await assert.rejects(performRelease(harness.options, harness.commands), /lint failed/);
  assert.deepEqual(await harness.versions(), ["0.1.10", "0.1.10"]);
  assert.ok(!harness.events.includes("confirm"));
});

test("publication updates all package versions, supplies explicit tag, and preserves other fields", async (t) => {
  const harness = await releaseHarness(t);
  let publishArgs;
  harness.commands.run = (_command, args) => { if (args.includes("publish")) publishArgs = args; };
  assert.equal((await performRelease(harness.options, harness.commands)).status, "published");
  assert.deepEqual(await harness.versions(), ["0.2.0", "0.2.0"]);
  assert.equal(publishArgs[publishArgs.indexOf("--tag") + 1], "next");
  for (const pkg of harness.options.packages) assert.equal(JSON.parse(await readFile(pkg.packageJsonPath, "utf8")).description, "Retain unrelated fields");
});

test("partial publication keeps target versions and reports an explicit recovery path", async (t) => {
  const harness = await releaseHarness(t);
  harness.commands.run = (_command, args) => { if (args.includes("publish")) throw new Error("second package failed"); };
  await assert.rejects(performRelease(harness.options, harness.commands), /partial-publication recovery.*No external publication was rolled back/);
  assert.deepEqual(await harness.versions(), ["0.2.0", "0.2.0"]);
});

test("package discovery returns no packages for a missing workspace", () => {
  assert.deepEqual(getPublishablePackages("/tmp/pptkit-release-missing"), []);
});

test("repository release set contains the eight engine packages", () => {
  assert.deepEqual(
    getPublishablePackages().map((pkg) => pkg.name),
    [
      "@pptkit/cli",
      "@pptkit/core",
      "@pptkit/layout",
      "@pptkit/media-inspection",
      "@pptkit/pdf-renderer",
      "@pptkit/pptx-exporter",
      "@pptkit/pptx-parser",
      "@pptkit/svg-renderer",
    ],
  );
});
