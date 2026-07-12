# Release Strategy

This document records the release posture for the bootstrap phase.

## Current Phase

PPTKit is in a pre-release workspace bootstrap stage.

The repository is structured like a publishable monorepo, but public package publication is not enabled yet.

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

