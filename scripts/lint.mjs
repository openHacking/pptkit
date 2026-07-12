import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { listPackageDirs, root } from "./package-utils.mjs";

const packageDirs = listPackageDirs();

const requiredRootFiles = [
  "package.json",
  "pnpm-workspace.yaml",
  "tsconfig.json",
  "tsconfig.base.json",
  "README.md",
];

const issues = [];

for (const file of requiredRootFiles) {
  if (!existsSync(path.join(root, file))) {
    issues.push(`Missing required root file: ${file}`);
  }
}

for (const dirName of packageDirs) {
  const expectedPackageName = `@pptkit/${dirName}`;
  const packageRoot = path.join(root, "packages", dirName);
  const manifestPath = path.join(packageRoot, "package.json");
  const readmePath = path.join(packageRoot, "README.md");
  const sourcePath = path.join(packageRoot, "src", "index.ts");
  const testPath = path.join(packageRoot, "test", "smoke.test.mjs");

  for (const file of [manifestPath, readmePath, sourcePath, testPath]) {
    if (!existsSync(file)) {
      issues.push(`Missing package file: ${path.relative(root, file)}`);
    }
  }

  if (!existsSync(manifestPath)) {
    continue;
  }

  const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));

  if (manifest.name !== expectedPackageName) {
    issues.push(
      `Package ${dirName} should be named ${expectedPackageName}, found ${String(manifest.name)}`,
    );
  }

  if (manifest.private !== false) {
    issues.push(
      `Package ${expectedPackageName} should set "private": false for future publishing.`,
    );
  }
}

if (issues.length > 0) {
  console.error("Lint failed:");
  for (const issue of issues) {
    console.error(`- ${issue}`);
  }
  process.exitCode = 1;
} else {
  console.log("Lint passed.");
}
