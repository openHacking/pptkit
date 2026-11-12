// Repository release policy, deliberately limited to numeric and preview.N releases.
export function parseVersion(value) {
  if (typeof value !== "string") return undefined;
  const match = /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-preview\.(0|[1-9]\d*))?$/.exec(value);
  if (!match) return undefined;
  const numbers = match.slice(1, 4).map(Number);
  const preview = match[4] === undefined ? undefined : Number(match[4]);
  if (!numbers.every(Number.isSafeInteger) || (preview !== undefined && !Number.isSafeInteger(preview))) return undefined;
  return { major: numbers[0], minor: numbers[1], patch: numbers[2], preview };
}

export const validateVersion = (value) => parseVersion(value) !== undefined;

export function compareVersions(left, right) {
  const a = parseVersion(left);
  const b = parseVersion(right);
  if (!a || !b) throw new Error("Cannot compare unsupported versions");
  for (const key of ["major", "minor", "patch"]) if (a[key] !== b[key]) return a[key] < b[key] ? -1 : 1;
  if (a.preview === b.preview) return 0;
  if (a.preview === undefined) return 1;
  if (b.preview === undefined) return -1;
  return a.preview < b.preview ? -1 : 1;
}

export function bumpVersion(version, releaseType) {
  const current = parseVersion(version);
  if (!current) throw new Error(`Unsupported version format: ${version}`);
  const { major, minor, patch, preview } = current;
  let next;
  if (releaseType === "patch") next = `${major}.${minor}.${preview === undefined ? patch + 1 : patch}`;
  else if (releaseType === "minor") next = `${major}.${minor + 1}.0`;
  else if (releaseType === "major") next = `${major + 1}.0.0`;
  else if (releaseType === "preview") next = preview === undefined ? `${major}.${minor}.${patch + 1}-preview.0` : `${major}.${minor}.${patch}-preview.${preview + 1}`;
  else throw new Error(`Unsupported release type: ${releaseType}`);
  if (!validateVersion(next)) throw new Error("Version increment exceeds supported integer range");
  return next;
}

export function releaseTag(version, requested) {
  const parsed = parseVersion(version);
  if (!parsed) throw new Error(`Invalid release version: ${version}`);
  const preview = parsed.major === 0 || parsed.preview !== undefined;
  const tag = requested ?? (preview ? "next" : "latest");
  if (tag !== "next" && tag !== "latest") throw new Error("Release tag must be next or latest");
  if (preview && tag === "latest") throw new Error("Provisional 0.x and preview.N releases must use the next tag");
  return tag;
}

export function assertReleaseNotes(markdown, version, dryRun) {
  const sections = [...markdown.matchAll(/^## \[([^\]]+)\][^\n]*\n([\s\S]*?)(?=^## |$(?![\s\S]))/gm)];
  const section = sections.find((match) => match[1] === version) ?? (dryRun ? sections.find((match) => match[1] === "Unreleased") : undefined);
  if (!section || !section[2].trim()) throw new Error(`CHANGELOG.md requires a nonempty [${version}] entry${dryRun ? " or [Unreleased] for dry runs" : ""}`);
}
