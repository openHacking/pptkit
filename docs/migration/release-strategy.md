# Release and Compatibility Policy

The eight publishable packages use one synchronized version; `@pptkit/svg-parser` is private research code and is not published by this workflow. APIs remain provisional and Canonical IR is v2. This policy governs future publications; editing this repository does not change any existing npm version or dist-tag.

## Version and tag rules

- Supported versions are `x.y.z` and `x.y.z-preview.N`, with nonnegative safe integer components and no leading zeros. Other prerelease labels/build metadata are intentionally outside this release workflow.
- Before 1.0, patch releases preserve public signatures, documented defaults and valid existing document behavior. New compatible APIs can ship in patches; a breaking API/default/IR change requires a minor version and migration notes. Bug fixes must describe any observable output correction.
- At 1.0 and later, patches fix compatible behavior, minors add compatible capabilities, and majors carry breaking changes. The v1 release gate remains STB-01–STB-04.
- Every `0.x` or `-preview.N` publication uses `next`; it cannot advance `latest`. Stable numeric versions at 1.0 or later default to `latest` and may explicitly use `next` for staged evaluation.
- The menu offers patch/minor/major/custom/preview. From a numeric version, preview advances the patch and starts `preview.0`; from a preview it increments N. Patch from a preview promotes its existing numeric base. Minor/major increments their respective component and resets lower components.
- The target must be newer than the current unified workspace version. Mixing versions or republishing an existing target fails before mutation.

Do not imply that a dist-tag establishes native Office compatibility. Supported behavior is the documented capability set and its recorded evidence. Pin exact versions for reproducible builds and evaluate a preview in a separate consumer before updating production use.

## Public API and IR compatibility

Public entry points and exported types are listed in the [capability audit](../planning/capability-matrix.md). Private source/dist paths are not supported contracts. All published manifests declare Node.js >=20 for installation/tooling; the default Core/Layout/Renderer/exporter byte APIs and `/browser` bundles remain browser-neutral. Only `/node` adds local path/file operations.

A breaking normalized shape or changed field meaning increments `irVersion`. Package version and IR version are independent; a package patch does not imply an IR bump. Consumers persisting IR must store its version and reject unsupported versions explicitly. PPTKit does not currently supply automatic migration from older IR or parse arbitrary stored JSON through `resolveNormalizedLayout`; that entry requires validated current IR.

During 0.x, document breaking removals and migration examples in the release changelog and API docs before publishing a minor. Where practical, retain a deprecated compatibility API through at least one minor; if a correctness issue makes retention impractical, state the exception and replacement. After 1.0, deprecations remain available until the next major. No compatibility promise covers undocumented fields or vendor-private OOXML.

## Changelog and release gate

[CHANGELOG.md](../../CHANGELOG.md) is the release-note source of truth. A real publication requires a nonempty heading matching the exact target version, including the preview suffix. Each entry records changes, public API/default/IR compatibility, migration guidance when applicable, validation evidence and remaining limitations. A dry run may use Unreleased notes. Historical entries must not be invented from the current workspace state.

Before publishing, review the diff and candidate capability scope, update the changelog, then pass build, typecheck, lint and the complete test suite. The suite includes actual pnpm tarballs installed by npm into an isolated consumer, Node/browser/type/CLI checks, package tests, XML fixtures and visual regression. Native application verification and known limitations must be accurately recorded in the [application matrix](../testing/application-compatibility.md).

## Local release commands

Authenticate to the intended registry/scope, then run the interactive workflow:

```sh
npm whoami
pnpm release:npm
```

Inspect a proposed target without updating versions or publishing:

```sh
pnpm release:npm --dry-run --version 0.1.11-preview.0 --tag next
```

This target is an example, not a reservation. Choose a newer unpublished version if it already exists. `--version` bypasses version selection; a real release still requests final publication confirmation after validation. There is no implicit `--yes` mode.

The script verifies the target is absent for every engine package. Only a structured registry E404 establishes absence; authentication, network, timeout and ambiguous errors stop the release. It runs `build`, `typecheck`, `lint`, then `test`, so declaration/example checks work on a clean checkout. The dry run tests **current-version** packed artifacts and leaves manifests/tags untouched; it does not claim to produce exact target-version tarballs.

After confirmation the script updates all engine manifest versions and invokes recursive pnpm publish with explicit public access and the selected tag. pnpm rewrites workspace dependency specifiers when packing/publishing. Publishing advances the selected dist-tag to the newly published version; other tags are not deliberately modified. The command does not create commits or Git tags and does not bypass npm authentication or OTP requirements. `--no-git-checks` permits the intentional manifest edits; maintainers must review and commit the final source and version changes and create their release Git tag afterwards.

## Partial-publication recovery

Publication across multiple packages is not atomic. On failure, preserve the target version files and inspect **every** package at that version before taking further action. Do not reset versions and rerun the entire release: some packages may already be immutable on the registry.

1. For each package, inspect `npm view @pptkit/<name>@<version> version dist.integrity dist-tags --json`. Distinguish confirmed E404 from connectivity/auth errors.
2. Compare any existing published artifact with the reviewed source/artifact for that release. If unexpected contents exist, stop recovery and prepare a corrected newer release; never claim the existing version was overwritten or rolled back.
3. Re-run validation with the unchanged target-version manifests. For packages confirmed absent, run `pnpm publish --access public --tag <selected-tag> --no-git-checks` from that package directory in dependency order (Core and media inspection, Layout/Parser, exporter/SVG renderer, then PDF renderer and CLI after their dependencies). Publish only the missing packages.
4. Verify every publishable package at the exact version and its selected tag, install them together in a consumer, then commit/tag the completed release. If abandoning the release, document the partial set and publish a corrected unified version; tag changes/deprecation are separate explicit registry operations.

Recovery is deliberately manual; the normal command refuses an existing target and never silently skips packages or rolls back external publication. [Workflow tests](../../scripts/release.test.mjs) exercise mutation ordering, cancellation, failure behavior, tag safety and version retention using temporary manifests and simulated subprocesses. They never publish.

## Installed consumer verification

Run `pnpm test:packages` to build and pack publishable packages through pnpm, install only those archives and their declared dependencies with npm in an OS temporary directory, then exercise public entry points. See [installed package tests](../../scripts/installed-packages.test.mjs) and [consumer sources](../../tests/consumers/node.mjs).

The consumer has no workspace symlinks, aliases or Node type search path. TypeScript resolves declarations from the installed exports. Browser ESM/global bundles run from installed files with network requests blocked. CLI execution uses npm's installed binary shim. Temporary consumers/tarballs are removed after the test. Installation may need registry access or a populated npm cache for external dependencies such as fflate; browser execution itself is offline.

A changeset system is not required while packages share a release version and changelog. Revisit it if packages gain independent release ownership; do not add a dependency solely to mirror this policy.

The unreleased [slide-bound timing extension](../api/slide-event-timeline.md) adds optional versioned eventTimeline metadata to experimental IR v2 and new managed-slide methods. Untimed shapes/output remain unchanged. Structural implementations of PresentationSlide must provide the added methods; consumers encountering the extension must retain it or reject/explicitly omit it. Native mapping is unavailable, so the current exporter refuses by default and requires an explicit loss mode for static output.
