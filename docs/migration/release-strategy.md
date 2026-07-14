# Release Strategy

This document records the local npm release workflow for the bootstrap phase.

## Current Phase

PPTKit is in a pre-release workspace bootstrap stage. Public preview packages can be
published locally when the publishing gate below has been satisfied.

## Versioning

- initial package manifests use `0.1.0`
- breaking changes are expected while APIs remain provisional
- documentation must call out provisional APIs explicitly

## Publishing Gate

Do not publish packages until all of the following are true:

- the canonical IR has a reviewed draft
- the `@pptkit/core` authoring surface has survived at least one API review
- the exporter writes real `.pptx` output
- local build, test, and CI workflows are stable

## Local npm release

From the repository root, log in to npm with an account that can publish the
`@pptkit` scope, then run:

```bash
npm whoami
pnpm release:npm
```

The command discovers all public packages under `packages/`, requires a unified
version, offers patch/minor/major/custom version selection, checks that the target
versions do not already exist on npm, runs the full validation suite, previews the
package contents, and asks for confirmation before publishing.

To validate the workflow without changing versions or publishing:

```bash
pnpm release:npm --dry-run
```

The release command does not create commits or Git tags. Review the resulting
version changes, commit them, and create a tag manually after a successful release.
If npm publishing fails after version files are updated, inspect the published
package state and the working tree before retrying; the command does not roll back
an external npm publication.

## Changesets

Changesets are intentionally deferred during workspace bootstrap.

Adopt them when one of these becomes true:

- preview releases are needed for external testing
- multiple packages start changing independently
- release notes need per-package tracking

## Pre-Release Recommendation

When the project is ready for external feedback:

- introduce changesets
- publish under pre-release tags
- keep documentation tightly aligned with the actually published package set
