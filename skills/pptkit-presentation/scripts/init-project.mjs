#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { cpSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const skillRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const starterRoot = path.join(skillRoot, "assets", "starter");
const themeIds = new Set(["clean-business", "swiss-grid", "editorial-story"]);

function usage(message) {
  if (message) process.stderr.write(`${message}\n\n`);
  process.stderr.write("Usage: init-project.mjs --output <directory> [--title <slug>] [--theme <id>] [--no-install]\n");
  process.exit(2);
}

function parseArgs(args) {
  const options = { title: "pptkit-deck", theme: "clean-business", install: true };
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === "--no-install") options.install = false;
    else if (arg === "--output" || arg === "--title" || arg === "--theme") {
      const value = args[index + 1];
      if (!value) usage(`Missing value for ${arg}`);
      options[arg.slice(2)] = value;
      index += 1;
    } else usage(`Unknown argument: ${arg}`);
  }
  if (!options.output) usage("--output is required");
  if (!themeIds.has(options.theme)) usage(`Unknown theme: ${options.theme}`);
  if (!/^[a-z0-9][a-z0-9._-]*$/i.test(options.title)) usage("--title must be a filesystem-safe slug");
  return options;
}

const options = parseArgs(process.argv.slice(2));
const output = path.resolve(options.output);
if (existsSync(output)) usage(`Output already exists: ${output}`);

mkdirSync(path.dirname(output), { recursive: true });
cpSync(starterRoot, output, { recursive: true, errorOnExist: true });

const manifestPath = path.join(output, "package.json");
const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
manifest.name = options.title.toLowerCase();
writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);

const specPath = path.join(output, "src", "deck-spec.ts");
const spec = readFileSync(specPath, "utf8")
  .replace('themeId: "clean-business"', `themeId: "${options.theme}"`)
  .replace('title: "Untitled PPTKit Deck"', `title: ${JSON.stringify(options.title.replace(/[-_]+/g, " "))}`);
writeFileSync(specPath, spec);

if (options.install) {
  const result = spawnSync("npm", ["install", "--no-audit", "--no-fund"], {
    cwd: output,
    stdio: "inherit",
  });
  if (result.error || result.status !== 0) {
    process.stderr.write("Project files were created, but npm install failed. Fix registry/network access and run npm install in the project.\n");
    process.exitCode = result.status ?? 1;
  }
}

process.stdout.write(`${JSON.stringify({ output, theme: options.theme, installed: options.install }, null, 2)}\n`);
