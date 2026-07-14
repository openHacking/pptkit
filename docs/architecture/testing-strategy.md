# Testing Strategy

PPTKit tests each responsibility at its narrowest useful boundary and then verifies the complete authoring-to-PPTX path.

## Required workspace checks

Contributors run the same commands expected in CI:

```bash
pnpm build
pnpm typecheck
pnpm lint
pnpm test
```

`pnpm lint` includes repository structure, dependency-boundary, documentation-link, stale-contract, and typed documentation-example checks.

## Package tests

- **Core:** operation semantics, immutable snapshots, IDs, duplication/removal, full diagnostics, normalized defaults, inheritance, and detached IR.
- **Layout:** detached results, connector anchors/bounds, image contain/cover, and nested group traversal.
- **PPTX exporter:** deterministic ZIP structure, OOXML parts/relationships, editable feature serialization, runtime asset behavior, warnings, and Node output.
- **SVG renderer:** deterministic serialization, all current element families, input detachment, asset/fidelity warnings, escaping, and real-Chromium visual fixtures.
- **CLI:** stable command entry behavior appropriate to its current minimal surface.

Tests prefer focused documents that isolate one contract. Cross-package tests cover boundaries where a feature would otherwise be “modeled but not exported.”

## Export verification layers

1. Validate authoring and normalized structures.
2. Open generated ZIPs and parse every XML/relationship part.
3. Assert key XML semantics and package relationships.
4. Render named fixture decks and compare reviewable snapshots.
5. Open/save representative decks in PowerPoint and LibreOffice.

The first three layers are automated today. Visual fixture coverage and a formal cross-application matrix are public-preview gates in the [Roadmap](../../ROADMAP.md).

SVG renderer fixtures run in pinned Playwright Chromium. Install the local browser once
with `pnpm exec playwright install chromium`; CI installs Chromium and its Linux
dependencies before `pnpm test`. Fixed viewport, device scale, color scheme, and a small
cross-platform pixel tolerance keep the preview snapshots reviewable.

## Fixtures and regressions

- Package fixtures live under `packages/<name>/test/fixtures/` when needed.
- Cross-package demonstrations live under `examples/` or typed documentation examples.
- Golden artifacts must be small, deterministic, and tied to a named behavior/regression.
- Large “everything” decks supplement rather than replace focused tests.
- Every fixed export regression adds the narrowest assertion that would have detected it.

## Review rule

A behavior change is incomplete without tests at the owning package and, when package boundaries are affected, an integration/export assertion. Public contract changes also update examples, API reference, and architecture documentation in the same change.
