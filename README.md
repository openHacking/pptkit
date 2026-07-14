# PPTKit

PPTKit is a developer-first TypeScript toolkit for building structured, editable presentations. It provides a format-independent authoring model, a stable Canonical Presentation IR, an independent layout stage, and editable PPTX output without exposing OOXML to application code.

> **Project status:** PPTKit is under active pre-release development. Preview packages are available on npm, and the public API may still change before a stable release.

## Features

- Method-managed presentation and slide authoring with stable document-wide IDs.
- Canonical Presentation IR v1 with complete diagnostics and materialized defaults.
- Rich text, themes, paints, strokes, layouts, placeholders, notes, and accessibility metadata.
- Images, preset shapes, anchored connectors, nested groups, and native editable tables.
- Detached layout resolution for connector anchors and image contain/cover behavior.
- Editable PPTX generation with native layouts, theme parts, notes, tables, media, and relationships.
- Browser-neutral byte generation plus a Node.js adapter for local assets and file output.

## Installation

PPTKit requires Node.js 20 or newer. Install the current preview packages in your
application with pnpm:

```bash
pnpm add @pptkit/core @pptkit/pptx-exporter
```

or:

```bash
npm install @pptkit/core @pptkit/pptx-exporter
```

Use `@pptkit/core` for authoring and validation. Use the default exporter entry for
browser generation, or the Node entry when writing a local `.pptx` file:

PPTKit is currently pre-release. See [Install PPTKit](docs/getting-started/installation.md)
for package entry points, runtime requirements, and asset behavior. Then continue
with the [Quick Start](docs/getting-started/quick-start.md) to create your first deck.

### Presentation Skill

PPTKit also includes the cross-agent `pptkit-presentation` skill for turning text,
documents, spreadsheets, and images into an editable PPTX plus its TypeScript source
project. Install the skill globally with:

```bash
npx skills add openHacking/pptkit --skill pptkit-presentation -g
```

Then ask your agent to create a presentation, for example: “Use PPTKit to turn this
quarterly report into an editable 10-slide PPTX.” The skill guides the brief and
outline, offers three curated visual themes, generates through the public PPTKit APIs,
and runs structural/package checks before delivery.

The skill pins matching PPTKit preview packages inside each generated project. See the
[Presentation Skill guide](docs/guides/presentation-skill.md) for the workflow and
repository-local evaluation steps.

## Quick Start

The following example assumes PPTKit is installed in your project and writes an
editable `hello-pptkit.pptx` file.

<!-- doc-test: docs/examples/quick-start.ts -->
```ts
import { createPresentation, validatePresentation } from "@pptkit/core";
import { writePptx } from "@pptkit/pptx-exporter/node";

const presentation = createPresentation({
  metadata: { title: "Hello PPTKit", author: "Example Team" },
  theme: { colors: { accent1: "2457D6" } },
});

const slide = presentation.addSlide();
slide.addElement({
  type: "text",
  content: [{
    runs: [
      { text: "Hello ", style: { fontSize: 36 } },
      {
        text: "PPTKit",
        style: { fontSize: 36, bold: true, color: { theme: "accent1" } },
      },
    ],
  }],
  box: { x: 64, y: 64, width: 520, height: 72 },
});

const diagnostics = validatePresentation(presentation);
if (diagnostics.some((diagnostic) => diagnostic.severity === "error")) {
  throw new Error(JSON.stringify(diagnostics, null, 2));
}

const result = await writePptx(presentation, {
  output: "./hello-pptkit.pptx",
});

console.log(result.status, result.output, result.warnings);
```

See the full [Quick Start](docs/getting-started/quick-start.md) for validation,
browser generation, assets, and warnings. Contributors can run the checked example
from the [Developer Workflow](docs/guides/developer-workflow.md).

## Packages

| Package | Responsibility |
| --- | --- |
| `@pptkit/core` | Authoring model, elements, assets, themes/layouts, validation, and Canonical IR. |
| `@pptkit/layout` | Detached connector and image geometry resolution. |
| `@pptkit/pptx-exporter` | Browser-neutral editable PPTX generation and Node.js file output. |
| `@pptkit/cli` | Thin command workflow shell; its authoring command surface is not stable yet. |

Parser, SVG, preview, and rendering packages remain roadmap work. See the [package overview](docs/api/package-overview.md).

## Architecture

The implemented forward pipeline is:

```text
Authoring Model → validation/normalization → Canonical IR v1
                → Layout Result → PPTX Package Model → .pptx
```

The separation keeps mutable authoring state, layout behavior, and OOXML packaging independently testable. Start with the [Architecture Overview](docs/architecture/overview.md) and [Core Authoring Model](docs/architecture/core-authoring-model.md).

## Documentation

- [Documentation index](docs/README.md)
- [Install PPTKit](docs/getting-started/installation.md)
- [Quick Start](docs/getting-started/quick-start.md)
- [Create Your First Deck](docs/guides/create-your-first-deck.md)
- [API reference](docs/api/README.md)
- [Core API](docs/api/core.md)
- [Layout API](docs/api/layout.md)
- [PPTX exporter API](docs/api/pptx-exporter.md)
- [Architecture](docs/architecture/README.md)

## Development

This section is for contributors working from the PPTKit monorepo. The complete
[Developer Workflow](docs/guides/developer-workflow.md) covers setup, validation,
documentation checks, and the examples workbench.

```bash
pnpm install
pnpm dev
pnpm build
pnpm typecheck
pnpm lint
pnpm test
```

`pnpm dev` starts the local examples workbench at `http://localhost:3210`.
See [CONTRIBUTING.md](CONTRIBUTING.md) for contribution expectations.

## Roadmap

The [project roadmap](ROADMAP.md) separates implemented foundations, public-preview gates, layout fidelity, import/round-trip work, ecosystem expansion, and deferred Office features.

## Contributing

Issues and contributions should preserve package ownership, add focused tests, update public/architecture documentation, and pass build, typecheck, lint, and test checks. See [CONTRIBUTING.md](CONTRIBUTING.md) for the complete workflow.

## License

PPTKit is licensed under the [MIT License](LICENSE).
