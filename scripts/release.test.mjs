import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { bumpVersion, getPublishablePackages, validateVersion } from "./release.mjs";

test("bumpVersion supports patch, minor, and major releases", () => {
  assert.equal(bumpVersion("1.2.3", "patch"), "1.2.4");
  assert.equal(bumpVersion("1.2.3", "minor"), "1.3.0");
  assert.equal(bumpVersion("1.2.3", "major"), "2.0.0");
});

test("version validation accepts only x.y.z versions", () => {
  assert.equal(validateVersion("0.1.0"), true);
  assert.equal(validateVersion("1.2"), false);
  assert.equal(validateVersion("v1.2.3"), false);
});

test("package discovery selects only public packages", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "pptkit-release-"));
  await mkdir(path.join(root, "public"));
  await mkdir(path.join(root, "private"));
  await writeFile(path.join(root, "public", "package.json"), JSON.stringify({ name: "@pptkit/public", version: "0.1.0", private: false }));
  await writeFile(path.join(root, "private", "package.json"), JSON.stringify({ name: "@pptkit/private", version: "0.1.0", private: true }));

  const packages = getPublishablePackages(root);
  assert.deepEqual(packages.map((pkg) => pkg.name), ["@pptkit/public"]);
});

test("package discovery returns no packages for a missing workspace", () => {
  assert.deepEqual(getPublishablePackages("/tmp/pptkit-release-missing"), []);
});

test("repository release set contains only the five engine packages", () => {
  assert.deepEqual(
    getPublishablePackages().map((pkg) => pkg.name),
    [
      "@pptkit/cli",
      "@pptkit/core",
      "@pptkit/layout",
      "@pptkit/pptx-exporter",
      "@pptkit/svg-renderer",
    ],
  );
});
