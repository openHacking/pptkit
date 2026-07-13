import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import { root } from "./package-utils.mjs";

const issues = [];
const packagesRoot = path.join(root, "packages");

function sourceFiles(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const entryPath = path.join(directory, entry.name);
    if (entry.isDirectory()) return sourceFiles(entryPath);
    return entry.isFile() && entry.name.endsWith(".ts") ? [entryPath] : [];
  });
}

function relativeToPackages(file) {
  return path.relative(packagesRoot, file);
}

if (!existsSync(path.join(root, "AGENTS.md"))) issues.push("Missing project AGENTS.md");

for (const packageName of ["layout", "pptx-exporter"]) {
  const sourceRoot = path.join(packagesRoot, packageName, "src");
  const unexpectedRootModules = readdirSync(sourceRoot, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.endsWith(".ts") && entry.name !== "index.ts");
  for (const entry of unexpectedRootModules) {
    issues.push(`${packageName} source root contains non-entry module: ${entry.name}`);
  }
}

for (const file of sourceFiles(path.join(packagesRoot, "core", "src"))) {
  const value = readFileSync(file, "utf8");
  if (value.includes("@pptkit/layout") || value.includes("@pptkit/pptx-exporter")) {
    issues.push(`core imports downstream package: ${relativeToPackages(file)}`);
  }
}

for (const file of sourceFiles(path.join(packagesRoot, "layout", "src"))) {
  const value = readFileSync(file, "utf8");
  if (/from ["']node:|from ["'](?:node:)?(?:fs|http|https|zlib)/.test(value) || /\.xml|\.rels|createZip/.test(value)) {
    issues.push(`layout contains output side effects or OOXML concerns: ${relativeToPackages(file)}`);
  }
}

for (const packageEntry of readdirSync(packagesRoot, { withFileTypes: true }).filter((entry) => entry.isDirectory())) {
  const sourceRoot = path.join(packagesRoot, packageEntry.name, "src");
  if (!existsSync(sourceRoot)) continue;
  for (const file of sourceFiles(sourceRoot)) {
    const value = readFileSync(file, "utf8");
    if (/from ["']@pptkit\/[a-z0-9-]+\//i.test(value)) {
      issues.push(`workspace package deep import: ${relativeToPackages(file)}`);
    }
  }
}

if (issues.length) {
  console.error("Boundary check failed:");
  for (const issue of issues) console.error(`- ${issue}`);
  process.exitCode = 1;
} else {
  console.log("Boundary check passed.");
}
