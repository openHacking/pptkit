# Source Organization

This document defines how source files should be organized inside PPTKit packages.

## Organize Around Stable Responsibilities

Package directories should describe cohesive responsibilities, not mirror a generic template. A package should add a directory when the directory owns a stable concept or stage of the package pipeline.

Examples include:

- `engine/` for layout resolution
- `assets/` for asset loading and preparation
- `ooxml/` for pure OOXML serialization
- `packaging/` for collecting package parts
- `archive/` for ZIP encoding
- `output/` for filesystem output

Do not create empty directories for planned features. Add `constraints/`, `measurement/`, `pagination/`, or similar directories only when concrete implementations exist.

## Shared Categories

The following directories are shared categories, not default buckets:

- `types/` contains contracts shared across multiple responsibilities or intentionally exposed through the package entry point.
- `constants/` contains stable domain or protocol constants used across multiple modules.
- `utils/` contains side-effect-free helpers reused by multiple responsibilities.

A helper used by only one module should remain private to that module. Avoid catch-all files such as `helpers.ts`, `common.ts`, or a single type file that mixes public API contracts with unrelated internal state.

## Entry Points and Dependency Direction

For packages with multiple responsibilities, `src/index.ts` should be the only root source module. It owns public exports and top-level orchestration, while implementation remains behind responsibility directories.

Dependencies should flow in one direction:

```text
public entry point / orchestration
              |
              v
      responsibility modules
              |
              v
 shared types, constants, and pure utilities
```

Responsibility modules must not import the package entry point, and sibling responsibilities must not form cycles. Imports from another workspace package must use that package's public export rather than a private source or generated path.

## Pure Logic and Side Effects

Pure transformations should remain separate from filesystem, network, process, compression, and other environment-dependent work. This keeps domain behavior independently testable and makes failure boundaries explicit.

For example, an exporter may orchestrate asset I/O, OOXML serialization, package collection, ZIP encoding, and file output, but those stages should remain separate modules rather than a single export function with hidden side effects.

## Review Guidance

Before adding or moving a source file, verify that:

- its directory owns the file's primary responsibility
- shared categories contain genuinely shared code
- the change does not introduce a reverse dependency or cycle
- public exports remain intentional
- side effects remain isolated from pure transformations
