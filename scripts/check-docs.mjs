import { existsSync, readFileSync, readdirSync, statSync } from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { root } from "./package-utils.mjs";

const ignoredDirectories = new Set([".git", "dist", "node_modules"]);
const issues = [];

function collectMarkdownFiles(directory) {
  const files = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    if (entry.isDirectory()) {
      if (!ignoredDirectories.has(entry.name)) files.push(...collectMarkdownFiles(path.join(directory, entry.name)));
    } else if (entry.name.endsWith(".md")) {
      files.push(path.join(directory, entry.name));
    }
  }
  return files;
}

function githubSlug(value) {
  return value
    .trim()
    .toLowerCase()
    .replace(/<[^>]+>/g, "")
    .replace(/[`*_~]/g, "")
    .replace(/[^\p{L}\p{N}\s-]/gu, "")
    .replace(/\s+/g, "-");
}

function headingAnchors(markdown) {
  const counts = new Map();
  const anchors = new Set();
  for (const line of markdown.split(/\r?\n/)) {
    const match = /^(?: {0,3})#{1,6}\s+(.+?)\s*#*\s*$/.exec(line);
    if (match === null) continue;
    const base = githubSlug(match[1]);
    const count = counts.get(base) ?? 0;
    counts.set(base, count + 1);
    anchors.add(count === 0 ? base : `${base}-${count}`);
  }
  return anchors;
}

function checkLinks(file, markdown, anchorCache) {
  const linkPattern = /!?\[[^\]]*\]\(([^)]+)\)/g;
  for (const match of markdown.matchAll(linkPattern)) {
    let target = match[1].trim();
    if (target.startsWith("<") && target.endsWith(">")) target = target.slice(1, -1);
    if (/^(?:https?:|mailto:|data:)/i.test(target)) continue;

    const [rawPath, rawAnchor] = target.split("#", 2);
    const targetPath = rawPath === "" ? file : path.resolve(path.dirname(file), decodeURIComponent(rawPath));
    const resolvedPath = existsSync(targetPath) && statSync(targetPath).isDirectory()
      ? path.join(targetPath, "README.md")
      : targetPath;

    if (!existsSync(resolvedPath)) {
      issues.push(`${path.relative(root, file)}: broken local link ${target}`);
      continue;
    }

    if (rawAnchor !== undefined && rawAnchor !== "") {
      let anchors = anchorCache.get(resolvedPath);
      if (anchors === undefined) {
        anchors = headingAnchors(readFileSync(resolvedPath, "utf8"));
        anchorCache.set(resolvedPath, anchors);
      }
      const anchor = decodeURIComponent(rawAnchor).toLowerCase();
      if (!anchors.has(anchor)) issues.push(`${path.relative(root, file)}: missing heading anchor ${target}`);
    }
  }
}

function checkDocTests(file, markdown) {
  const markerCount = [...markdown.matchAll(/<!--\s*doc-test:/g)].length;
  const pattern = /<!--\s*doc-test:\s*([^\s]+)\s*-->\s*\n```(?:ts|typescript)\n([\s\S]*?)\n```/g;
  let matched = 0;
  for (const match of markdown.matchAll(pattern)) {
    matched += 1;
    const sourcePath = path.resolve(root, match[1]);
    if (!existsSync(sourcePath)) {
      issues.push(`${path.relative(root, file)}: missing doc-test source ${match[1]}`);
      continue;
    }
    const source = readFileSync(sourcePath, "utf8").trimEnd();
    const snippet = match[2].trimEnd();
    if (source !== snippet) issues.push(`${path.relative(root, file)}: snippet differs from ${match[1]}`);
  }
  if (matched !== markerCount) issues.push(`${path.relative(root, file)}: doc-test marker is not followed by a TypeScript fence`);
}

const stalePatterns = [
  [/canonical-ir-v0\.1/gi, "removed Canonical IR v0.1 reference"],
  [/fontWeight\s*:/g, "removed fontWeight style field"],
  [/altText\s*:/g, "removed altText field"],
  [/shape\s*:\s*["']line["']/g, "removed line shape; use connector"],
  [/background\s*:\s*["']#?[0-9a-f]{6}["']/gi, "legacy string background; use PaintInput"],
  [/createPresentation\s*\(\s*\{\s*title\s*:/gs, "legacy top-level presentation title"],
  [/addSlide\s*\(\s*\{\s*elements\s*:/gs, "legacy direct slide element array example"],
];

const anchorCache = new Map();
for (const file of collectMarkdownFiles(root)) {
  const markdown = readFileSync(file, "utf8");
  checkLinks(file, markdown, anchorCache);
  checkDocTests(file, markdown);
  for (const [pattern, description] of stalePatterns) {
    pattern.lastIndex = 0;
    if (pattern.test(markdown)) issues.push(`${path.relative(root, file)}: ${description}`);
  }
}

const tscPath = path.join(root, "node_modules", "typescript", "bin", "tsc");
if (!existsSync(tscPath)) {
  issues.push("TypeScript is not installed; run pnpm install before checking documentation examples.");
} else {
  const result = spawnSync(process.execPath, [tscPath, "-p", path.join(root, "docs", "examples", "tsconfig.json")], {
    cwd: root,
    encoding: "utf8",
  });
  if (result.status !== 0) {
    issues.push(`documentation example typecheck failed:\n${result.stdout}${result.stderr}`.trimEnd());
  }
}

if (issues.length > 0) {
  console.error("Documentation check failed:");
  for (const issue of issues) console.error(`- ${issue}`);
  process.exitCode = 1;
} else {
  console.log("Documentation links, contracts, and typed examples passed.");
}
