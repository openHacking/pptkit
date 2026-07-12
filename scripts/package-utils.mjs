import { existsSync, readdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));

export const root = path.resolve(here, "..");

export function listPackageDirs() {
  const packagesRoot = path.join(root, "packages");

  return readdirSync(packagesRoot, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .filter((dirName) =>
      existsSync(path.join(packagesRoot, dirName, "package.json")),
    )
    .sort((left, right) => left.localeCompare(right));
}
