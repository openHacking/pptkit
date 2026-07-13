# Monorepo Bootstrap

This document records the repository structure and tooling decisions made during the PPTKit workspace bootstrap. The workspace now contains implemented Core, Layout, and PPTX export pipelines; this page remains the historical tooling rationale.

It complements the higher-level architecture docs and focuses on implementation-facing repository decisions.

## Summary

PPTKit uses a `pnpm`-managed TypeScript monorepo with a small initial package set and repository-wide validation commands.

The bootstrap made package boundaries executable before the main feature implementations landed; those boundaries remain enforced today.

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
- `examples/` for runnable development samples
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

Workspace validation remains intentionally understandable:

- Node's built-in test runner provides package smoke tests
- repository lint scripts check required files, dependency boundaries, documentation links/contracts, and typed documentation examples
- CI runs the same commands contributors run locally

## Release Posture

- packages currently use `0.1.0`
- publishing is not enabled yet, but package manifests stay publish-shaped
- release automation is deferred until preview publishing becomes useful

## Examples and Fixtures

- runnable examples belong under `examples/`
- package-specific fixtures should live close to the owning package
- golden files are introduced only for stable, named output behaviors

## Documentation Rule

Design decisions should be written directly into the long-lived docs that own the topic.

For this repository that means:

- architecture decisions go into `docs/architecture/`
- API decisions go into `docs/api/`
- contributor workflow decisions go into `docs/guides/` or `CONTRIBUTING.md`

Avoid creating a parallel decision-record process for changes that are better expressed as maintained product and engineering documentation.
