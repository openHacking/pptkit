import { execFileSync, spawnSync } from "node:child_process";
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createInterface } from "node:readline/promises";

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

export function bumpVersion(version, releaseType) {
  const match = /^(\d+)\.(\d+)\.(\d+)$/.exec(version);
  if (!match) throw new Error(`Unsupported version format: ${version}`);

  const [major, minor, patch] = match.slice(1).map(Number);
  if (releaseType === "patch") return `${major}.${minor}.${patch + 1}`;
  if (releaseType === "minor") return `${major}.${minor + 1}.0`;
  if (releaseType === "major") return `${major + 1}.0.0`;
  throw new Error(`Unsupported release type: ${releaseType}`);
}

export function validateVersion(version) {
  return /^\d+\.\d+\.\d+$/.test(version);
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

function assertTargetVersionsAvailable(packages, version) {
  for (const pkg of packages) {
    const result = spawnSync(
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
    const output = `${result.stdout ?? ""}\n${result.stderr ?? ""}`;

    if (result.status === 0) {
      throw new Error(`Version already exists on npm: ${pkg.name}@${version}`);
    }
    if (result.error || result.signal || !/E404|notarget|404 Not Found/i.test(output)) {
      throw new Error(
        `Could not verify ${pkg.name}@${version} on npm. Check npm login, registry connectivity, and try again.`,
      );
    }
  }
}

function validatePackageContents(packages) {
  for (const pkg of packages) {
    run("npm", ["pack", "--dry-run", "--json"], { cwd: pkg.dir });
  }
}

function updatePackageVersions(packages, nextVersion) {
  for (const pkg of packages) {
    writeJson(pkg.packageJsonPath, { ...pkg.packageJson, version: nextVersion });
  }
}

function parseArgs(args) {
  if (args.length === 0) return { dryRun: false };
  if (args.length === 1 && args[0] === "--dry-run") return { dryRun: true };
  throw new Error("Usage: pnpm release:npm [--dry-run]");
}

async function chooseVersion(currentVersion, rl) {
  console.log(`Current version: ${currentVersion}`);
  console.log("Select release type:");
  console.log("1. patch");
  console.log("2. minor");
  console.log("3. major");
  console.log("4. custom");

  const choice = (await rl.question("Enter choice [1-4]: ")).trim();
  if (choice === "1") return bumpVersion(currentVersion, "patch");
  if (choice === "2") return bumpVersion(currentVersion, "minor");
  if (choice === "3") return bumpVersion(currentVersion, "major");
  if (choice === "4") {
    const customVersion = (await rl.question("Enter custom version (x.y.z): ")).trim();
    if (!validateVersion(customVersion)) throw new Error(`Invalid custom version: ${customVersion}`);
    return customVersion;
  }
  throw new Error(`Invalid choice: ${choice}`);
}

async function main(args = process.argv.slice(2)) {
  const { dryRun } = parseArgs(args);
  const packages = getPublishablePackages();
  if (packages.length === 0) throw new Error("No publishable packages found in ./packages");

  const currentVersion = assertVersionsMatch(packages);
  const rl = createInterface({ input: process.stdin, output: process.stdout });

  try {
    const nextVersion = await chooseVersion(currentVersion, rl);
    console.log(`\nPackages to ${dryRun ? "inspect" : "publish"} (${packages.length}):`);
    for (const pkg of packages) console.log(`- ${pkg.name}: ${pkg.version} -> ${nextVersion}`);

    assertTargetVersionsAvailable(packages, nextVersion);
    console.log("\nRunning validation...");
    run("pnpm", ["lint"]);
    run("pnpm", ["typecheck"]);
    run("pnpm", ["test"]);
    run("pnpm", ["build"]);
    validatePackageContents(packages);

    if (dryRun) {
      console.log("\nDry run complete. No package versions were changed and nothing was published.");
      return;
    }

    const shouldContinue = (await rl.question("\nContinue with version update and publish? [y/N]: "))
      .trim()
      .toLowerCase();
    if (shouldContinue !== "y" && shouldContinue !== "yes") {
      console.log("Cancelled.");
      return;
    }

    updatePackageVersions(packages, nextVersion);
    console.log("\nPublishing packages to npm...");
    try {
      run("pnpm", ["-r", "--filter", "./packages/*", "publish", "--access", "public", "--no-git-checks"]);
    } catch (error) {
      throw new Error(
        `npm publish failed after version files were updated. Review the published packages and git diff before retrying. ${error.message}`,
      );
    }
    console.log(`\nPublished ${packages.length} packages at version ${nextVersion}.`);
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
