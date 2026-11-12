import { execFileSync, spawnSync } from "node:child_process";
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createInterface } from "node:readline/promises";
import { assertReleaseNotes, bumpVersion, compareVersions, releaseTag, validateVersion } from "./release-policy.mjs";

export { bumpVersion, validateVersion } from "./release-policy.mjs";

export const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
export const packagesDir = path.join(repoRoot, "packages");

function readJson(filePath) {
  return JSON.parse(readFileSync(filePath, "utf8"));
}

function writeJson(filePath, value) {
  writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`);
}

export function getPublishablePackages(packagesRoot = packagesDir) {
  if (!existsSync(packagesRoot)) return [];

  return readdirSync(packagesRoot, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => {
      const dir = path.join(packagesRoot, entry.name);
      const packageJsonPath = path.join(dir, "package.json");
      if (!existsSync(packageJsonPath)) return null;
      const packageJson = readJson(packageJsonPath);
      if (packageJson.private !== false) return null;
      return {
        dir,
        packageJsonPath,
        packageJson,
        name: packageJson.name,
        version: packageJson.version,
      };
    })
    .filter(Boolean)
    .sort((left, right) => left.name.localeCompare(right.name));
}

function run(command, args, options = {}) {
  execFileSync(command, args, {
    cwd: repoRoot,
    stdio: "inherit",
    ...options,
  });
}

function assertVersionsMatch(packages) {
  const versions = [...new Set(packages.map((pkg) => pkg.version))];
  if (versions.length !== 1) {
    throw new Error(`Expected a unified version across packages, found: ${versions.join(", ")}`);
  }
  return versions[0];
}

export function assertTargetVersionsAvailable(packages, version, spawn = spawnSync) {
  for (const pkg of packages) {
    const result = spawn(
      "npm",
      [
        "view",
        `${pkg.name}@${version}`,
        "version",
        "--json",
        "--fetch-timeout=10000",
        "--fetch-retries=0",
      ],
      { cwd: repoRoot, encoding: "utf8", timeout: 15000 },
    );
    if (result.status === 0) {
      throw new Error(`Version already exists on npm: ${pkg.name}@${version}`);
    }
    let code;
    try { code = JSON.parse(result.stdout)?.error?.code; } catch { /* Unstructured failures are not proof of absence. */ }
    if (result.error || result.signal || code !== "E404") {
      throw new Error(
        `Could not verify ${pkg.name}@${version} on npm. Check npm login, registry connectivity, and try again.`,
      );
    }
  }
}

function updatePackageVersions(packages, nextVersion) {
  for (const pkg of packages) {
    writeJson(pkg.packageJsonPath, { ...pkg.packageJson, version: nextVersion });
  }
}

export function parseArgs(args) {
  const options = { dryRun: false };
  const seen = new Set();
  for (let index = 0; index < args.length; index++) {
    const arg = args[index];
    if (seen.has(arg)) throw new Error(`Duplicate option: ${arg}`);
    seen.add(arg);
    if (arg === "--dry-run") options.dryRun = true;
    else if (arg === "--version" || arg === "--tag") {
      const value = args[++index];
      if (!value || value.startsWith("--")) throw new Error(`Missing value for ${arg}`);
      options[arg.slice(2)] = value;
    } else throw new Error("Usage: pnpm release:npm [--dry-run] [--version x.y.z[-preview.N]] [--tag next|latest]");
  }
  if (options.version !== undefined && !validateVersion(options.version)) throw new Error(`Invalid version: ${options.version}`);
  if (options.tag !== undefined && !["next", "latest"].includes(options.tag)) throw new Error("Release tag must be next or latest");
  return options;
}

async function chooseVersion(currentVersion, rl) {
  console.log(`Current version: ${currentVersion}`);
  console.log("Select release type:");
  console.log("1. patch");
  console.log("2. minor");
  console.log("3. major");
  console.log("4. custom");
  console.log("5. preview");

  const choice = (await rl.question("Enter choice [1-5]: ")).trim();
  if (choice === "1") return bumpVersion(currentVersion, "patch");
  if (choice === "2") return bumpVersion(currentVersion, "minor");
  if (choice === "3") return bumpVersion(currentVersion, "major");
  if (choice === "4") {
    const customVersion = (await rl.question("Enter custom version (x.y.z or x.y.z-preview.N): ")).trim();
    if (!validateVersion(customVersion)) throw new Error(`Invalid custom version: ${customVersion}`);
    return customVersion;
  }
  if (choice === "5") return bumpVersion(currentVersion, "preview");
  throw new Error(`Invalid choice: ${choice}`);
}

// Explicit process dependencies let tests exercise failure ordering with actual
// temporary manifests, without publishing or requiring registry credentials.
export async function performRelease({ packages, nextVersion, tag, dryRun = false, changelog }, commands) {
  if (packages.length === 0) throw new Error("No publishable packages found in ./packages");
  const currentVersion = assertVersionsMatch(packages);
  if (compareVersions(nextVersion, currentVersion) <= 0) throw new Error("Target version must be newer than the current workspace version");
  const selectedTag = releaseTag(nextVersion, tag);
  assertReleaseNotes(changelog, nextVersion, dryRun);
  commands.log(`Packages to ${dryRun ? "inspect" : "publish"} at ${nextVersion} (tag: ${selectedTag}):`);
  for (const pkg of packages) commands.log(`- ${pkg.name}: ${pkg.version} -> ${nextVersion}`);
  commands.checkRegistry(packages, nextVersion);
  // Build first: declaration and documentation checks need built package exports
  // on a clean checkout. The test suite includes actual installed tarball tests.
  for (const check of ["build", "typecheck", "lint", "test"]) commands.run("pnpm", [check]);
  if (dryRun) {
    commands.log("Dry run complete. Current-version tarballs passed consumer tests; no versions or registry tags were changed.");
    return { status: "dry-run", version: nextVersion, tag: selectedTag };
  }
  if (!await commands.confirm(`Publish ${packages.length} packages at ${nextVersion} with tag ${selectedTag}?`)) return { status: "cancelled", version: nextVersion, tag: selectedTag };
  updatePackageVersions(packages, nextVersion);
  try {
    commands.run("pnpm", ["-r", "--filter", "./packages/*", "publish", "--access", "public", "--tag", selectedTag, "--no-git-checks"]);
  } catch (error) {
    throw new Error(`npm publish failed after version files were updated to ${nextVersion}. Keep these versions, inspect registry state for every package, and follow docs/migration/release-strategy.md partial-publication recovery. No external publication was rolled back. ${error.message}`);
  }
  return { status: "published", version: nextVersion, tag: selectedTag };
}

async function main(args = process.argv.slice(2)) {
  const options = parseArgs(args);
  const packages = getPublishablePackages();
  if (packages.length === 0) throw new Error("No publishable packages found in ./packages");

  const currentVersion = assertVersionsMatch(packages);
  const rl = createInterface({ input: process.stdin, output: process.stdout });

  try {
    const nextVersion = options.version ?? await chooseVersion(currentVersion, rl);
    const result = await performRelease({ packages, nextVersion, tag: options.tag, dryRun: options.dryRun, changelog: readFileSync(path.join(repoRoot, "CHANGELOG.md"), "utf8") }, {
      run, checkRegistry: assertTargetVersionsAvailable, log: console.log,
      confirm: async (question) => ["y", "yes"].includes((await rl.question(`\n${question} [y/N]: `)).trim().toLowerCase()),
    });
    console.log(`Release ${result.status}: ${result.version} (${result.tag}).`);
  } finally {
    rl.close();
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(`\nRelease failed: ${error.message}`);
    process.exitCode = 1;
  });
}
