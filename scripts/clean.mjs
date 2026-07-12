import { rmSync } from "node:fs";
import path from "node:path";
import { listPackageDirs, root } from "./package-utils.mjs";

const packageDirs = listPackageDirs();

for (const relativePath of packageDirs.map((dirName) =>
  path.join("packages", dirName, "dist"),
)) {
  rmSync(path.join(root, relativePath), {
    force: true,
    recursive: true,
  });
}

for (const relativePath of packageDirs.map((dirName) =>
  path.join("packages", dirName, "tsconfig.tsbuildinfo"),
)) {
  rmSync(path.join(root, relativePath), {
    force: true,
  });
}

for (const relativePath of [
  path.join("examples", "dev-app", "dist"),
  path.join("examples", "dev-app", "tsconfig.app.tsbuildinfo"),
  path.join("examples", "dev-app", "tsconfig.server.tsbuildinfo"),
]) {
  rmSync(path.join(root, relativePath), {
    force: true,
    recursive: true,
  });
}

console.log("Cleaned build artifacts.");
