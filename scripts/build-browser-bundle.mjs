import { mkdir, readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

import { nodeResolve } from "@rollup/plugin-node-resolve";
import terser from "@rollup/plugin-terser";
import { rollup } from "rollup";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const packages = {
  "@pptkit/core": { directory: "core", external: [] },
  "@pptkit/layout": { directory: "layout", external: ["@pptkit/core"] },
  "@pptkit/pptx-exporter": {
    directory: "pptx-exporter",
    external: ["@pptkit/core", "@pptkit/layout"],
  },
  "@pptkit/svg-renderer": {
    directory: "svg-renderer",
    external: ["@pptkit/core", "@pptkit/layout"],
  },
};

function assertPackageName(value) {
  const config = packages[value];
  if (config === undefined) {
    throw new Error(`Unknown browser bundle package "${value}".`);
  }
  return config;
}

async function assertBrowserSafe(file) {
  const source = await readFile(file, "utf8");
  const forbidden = [
    [/(?:^|["'])node:/, "Node built-in import"],
    [/\brequire\s*\(/, "CommonJS require()"],
    [/\bprocess(?:\.|\[)/, "Node process global"],
  ];
  for (const [pattern, description] of forbidden) {
    if (pattern.test(source)) {
      throw new Error(`${path.relative(root, file)} contains a forbidden ${description}.`);
    }
  }
}

async function build(packageName) {
  const config = assertPackageName(packageName);
  const packageRoot = path.join(root, "packages", config.directory);
  const input = path.join(packageRoot, "dist", "index.js");
  const outputDirectory = path.join(packageRoot, "dist", "browser");
  await mkdir(outputDirectory, { recursive: true });

  const bundle = await rollup({
    input,
    external: config.external,
    plugins: [
      nodeResolve({
        browser: true,
        exportConditions: ["browser", "import", "default"],
      }),
    ],
    onwarn(warning, warn) {
      if (warning.code === "CIRCULAR_DEPENDENCY") {
        throw new Error(warning.message);
      }
      warn(warning);
    },
  });

  const esmFile = path.join(outputDirectory, "index.js");
  const globalFile = path.join(outputDirectory, "global.js");
  try {
    await bundle.write({
      file: esmFile,
      format: "es",
      plugins: [terser()],
    });
    await bundle.write({
      file: globalFile,
      format: "iife",
      name: "PPTKit",
      extend: true,
      globals: Object.fromEntries(config.external.map((dependency) => [dependency, "PPTKit"])),
      plugins: [terser()],
    });
  } finally {
    await bundle.close();
  }

  await Promise.all([assertBrowserSafe(esmFile), assertBrowserSafe(globalFile)]);
  console.log(`Built browser bundles for ${packageName}.`);
}

const packageName = process.argv[2];
if (packageName === undefined || process.argv.length !== 3) {
  throw new Error("Usage: node scripts/build-browser-bundle.mjs <package-name>");
}

await build(packageName);
