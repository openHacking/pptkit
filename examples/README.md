# Examples

This directory contains runnable development examples for PPTKit.

## Report dashboard

The [report dashboard](report-dashboard/dashboard.mjs) creates a fresh, editable one-slide PPTX with a chart and table, plus a static SVG preview. From the repository root, run `pnpm install`, `pnpm build`, then `pnpm --filter @pptkit-example/report-dashboard check` and `pnpm --filter @pptkit-example/report-dashboard start`. The last command writes `report-dashboard.pptx` and `report-dashboard.svg` in the example directory; existing files are not overwritten. If the second write fails, the first file may already exist. The sample prints validation and exporter/renderer warnings; inspect these before using the output. Preview appearance can differ from PowerPoint, and this example does not perform native application verification. Replace the illustrative data with actual results.

Run `node --test scripts/report-dashboard-example.test.mjs` after the build to check the generated chart, embedded workbook, table and preview through public APIs. The example's `package.json` declares the engine packages it uses; consumers can install their published equivalents outside this workspace.

## Diagram and reusable layout

The [diagram template](report-dashboard/diagram-template.mjs) defines one shared layout, uses its title placeholder on two pages, and connects two editable shapes using the experimental straight native connector profile. After the install and build above, run `pnpm --filter @pptkit-example/report-dashboard diagram` to create `diagram-template.pptx` in the example directory. Existing output is not overwritten. The command prints validation, export and SVG-preview warnings; the SVG previews are available from `renderDiagramTemplate()` but are not written by the command. The connector profile is limited to supported targets and sites; check warnings for fallback, and consult [native connectors](../docs/api/native-connectors.md) for its application evidence and remaining limits. The example tests inspect both connector attachments and the shared layout relationship, then reimport the slides. They do not verify interactive connector movement or native appearance.

## React preview and download

The [React example](react-preview/src/main.tsx) rebuilds a two-slide document from component state, displays each SVG in a sandboxed iframe and downloads fresh PPTX bytes. It uses the [incremental renderer](../docs/api/svg-invalidation.md), retaining the unchanged supporting slide when the title changes and committing a new cache only while the request is current. After `pnpm install` and `pnpm build:browser`, run `pnpm --filter @pptkit-example/react-preview dev` and open the displayed local address. `pnpm --filter @pptkit-example/react-preview build` typechecks and creates a production bundle. Its package declares the React, Vite and PPTKit dependencies. The example cancels obsolete asynchronous preview requests, displays warnings, and releases the download URL; it is a small integration example rather than an editor. Browser preview and native application rendering may differ. Avoid treating a sandboxed preview as a general-purpose sanitizer for arbitrary untrusted assets.

After the production build, `node --test scripts/react-preview-example.test.mjs` serves the built files locally and tests Chromium preview updates, rapid title changes, unchanged supporting-slide SVG, the downloaded two-slide PPTX and absence of external requests. Other browsers and native application behavior are not covered by that test.

## Developer Workbench

Run the internal workbench with:

```bash
pnpm dev
```

The workbench starts at `http://localhost:3210` and is designed for fast UI-level verification of whichever capabilities are most relevant to the current project phase.

The main navigation is feature-first so contributors can quickly isolate one capability at a time.
Each example may also include `scenarioTags` so broader end-to-end flows like `basic-deck` or `pitch-deck` remain reusable as the library grows.

## Fixture Contract

Examples in the workbench should register a single `ExampleDefinition` with:

- `id`
- `title`
- `feature`
- `description`
- `inputKind`
- `source`
- `scenarioTags`
- `expectedCapabilities`
- `status`

Use feature grouping for primary navigation and scenario tags for cross-feature regression coverage.
Do not treat any specific feature list as fixed: the first batch of cases should follow the actual implementation priorities in the repo.

## Current Layout

- `dev-app/`: internal preview server, registry, and workbench UI
- `browser-esm/`: zero-build CDN example for SVG preview and browser PPTX download
- `report-dashboard/`: Node report generation and SVG preview
- `react-preview/`: React browser preview and PPTX download
