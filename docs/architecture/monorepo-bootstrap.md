# Monorepo Bootstrap

This document records the repository structure and tooling decisions for the PPTKit workspace bootstrap phase.

It complements the higher-level architecture docs and focuses on implementation-facing repository decisions.

## Summary

PPTKit uses a `pnpm`-managed TypeScript monorepo with a small initial package set and repository-wide validation commands.

The goal is to make package boundaries executable early without pretending the real layout and export implementations already exist.

## Goals

- establish a real workspace contributors can install and validate locally
- preserve documented package boundaries while implementation is still early
- keep the initial toolchain small enough to evolve without heavy migration cost
- make room for future publishing and example projects without premature package splits

## Initial Workspace Layout

Package roots live under `packages/*`.

Initial packages:

- `@pptkit/core`
- `@pptkit/layout`
- `@pptkit/pptx-exporter`
- `@pptkit/cli`

Supporting directories:

- `docs/` for user and contributor documentation
- `examples/` for future runnable samples
- `scripts/` for small repository maintenance utilities

## Dependency Direction

Dependencies flow inward toward more format-aware packages:

- `@pptkit/core` depends on no PPTX-specific package
- `@pptkit/layout` may depend on `@pptkit/core`
- `@pptkit/pptx-exporter` may depend on `@pptkit/core` and `@pptkit/layout`
- `@pptkit/cli` should stay thin and orchestrate package APIs rather than own core logic

Reverse dependencies across those boundaries are not allowed.

## Build and TypeScript Strategy

- use TypeScript project references from the workspace root
- compile each package from `src/` to `dist/`
- publish only `dist/` output when package publishing begins
- use `NodeNext` module settings to stay aligned with modern Node.js ESM usage

## Validation Strategy

Repository-level validation commands:

- `pnpm lint`
- `pnpm typecheck`
- `pnpm test`
- `pnpm build`

Bootstrap validation intentionally stays lightweight:

- Node's built-in test runner provides package smoke tests
- a small repository lint script checks required workspace files and package shells
- CI runs the same commands contributors run locally

## Release Posture

- packages currently use `0.1.0`
- publishing is not enabled yet, but package manifests stay publish-shaped
- release automation is deferred until preview publishing becomes useful

## Examples and Fixtures

- runnable examples belong under `examples/`
- package-specific fixtures should live close to the owning package
- golden files should wait until export output is stable enough to snapshot

## Documentation Rule

Design decisions should be written directly into the long-lived docs that own the topic.

For this repository that means:

- architecture decisions go into `docs/architecture/`
- API decisions go into `docs/api/`
- contributor workflow decisions go into `docs/guides/` or `CONTRIBUTING.md`

Avoid creating a parallel decision-record process for changes that are better expressed as maintained product and engineering documentation.
