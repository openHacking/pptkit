# PPTKit

Modern presentation generation toolkit for JavaScript.

## Features

- Developer-first toolkit for generating editable presentation output
- Package-based architecture for core, layout, parsing, and export workflows
- Designed for modern structured content and automation scenarios
- Documentation-first repository structure for long-term open-source growth

## Quick Start

The repository now includes a formal `@pptkit/core` package for presentation construction, asset registration, and document normalization.

## Installation

For local development:

```bash
pnpm install
pnpm dev
pnpm lint
pnpm typecheck
pnpm test
```

Published package installation instructions will be added once preview releases exist.

`pnpm dev` starts the internal examples workbench at `http://localhost:3210` so contributors can browse capability-focused cases based on the current project priorities and implementation state.

## Hello World

```ts
import { createPresentation, normalizePresentation } from "@pptkit/core";

const presentation = createPresentation({
  title: "Hello PPTKit",
});

presentation.addSlide({
  elements: [
    {
      type: "text",
      text: "Hello PPTKit",
      box: { x: 64, y: 64, width: 320, height: 36 },
    },
  ],
});

const normalized = normalizePresentation(presentation);

console.log(normalized.slides.length);
```

`@pptkit/pptx-exporter` still exposes a placeholder export boundary and does not write a real `.pptx` file yet.

## Architecture

PPTKit is being designed as a multi-package toolkit with clear boundaries between content modeling, layout, parsing, and export.

See the contributor-facing architecture docs in [docs/architecture/README.md](docs/architecture/README.md).
Start with [docs/architecture/overview.md](docs/architecture/overview.md), [docs/architecture/core-authoring-model.md](docs/architecture/core-authoring-model.md), and [docs/architecture/document-models.md](docs/architecture/document-models.md).

## Packages

Bootstrap packages:

- `@pptkit/core`
- `@pptkit/layout`
- `@pptkit/pptx-exporter`
- `@pptkit/cli`

Planned follow-up packages:

- `@pptkit/pptx-parser`
- `@pptkit/svg-parser`
- `@pptkit/svg-renderer`

## Documentation

- Docs index: [docs/README.md](docs/README.md)
- Getting started: [docs/getting-started/README.md](docs/getting-started/README.md)
- Architecture: [docs/architecture/README.md](docs/architecture/README.md)
- API: [docs/api/README.md](docs/api/README.md)
- Core API: [docs/api/core.md](docs/api/core.md)
- Guides: [docs/guides/README.md](docs/guides/README.md)
- Examples: [examples/README.md](examples/README.md)
- Monorepo bootstrap: [docs/architecture/monorepo-bootstrap.md](docs/architecture/monorepo-bootstrap.md)

## Roadmap

The open-source roadmap is available in [ROADMAP.md](ROADMAP.md).

## Contributing

Contribution guidelines are available in [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT
