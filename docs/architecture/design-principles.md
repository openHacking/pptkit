# Design Principles

This document records the architecture guardrails for PPTKit.

## PPTKit Is a Document Engine First

PPTKit should start as a document engine for structured presentation generation, not as an attempt to recreate the full PowerPoint application surface on day one.

That means the early architecture should prioritize:

- document modeling
- normalization
- layout resolution
- export reliability
- parsing boundaries

It should not assume the project must immediately solve:

- full visual editing
- app-scale workflow orchestration
- complete office-suite parity
- every import and export path at once

## Package Boundaries Follow Responsibility

Packages should be split by responsibility, not by object category.

Good examples:

- `@pptkit/core`
- `@pptkit/layout`
- `@pptkit/pptx-exporter`
- `@pptkit/pptx-parser`

Bad examples:

- `@pptkit/text`
- `@pptkit/image`
- `@pptkit/shape`
- `@pptkit/table`

Text, images, shapes, and tables are all parts of one presentation document model. Splitting them into separate packages too early creates versioning churn, circular type pressure, fragmented docs, and unnecessary integration complexity.

## Layout Must Be a First-Class Layer

One lesson from older command-style PPT libraries is that authoring and file writing often become too tightly coupled. That usually pushes layout back onto the user in the form of manual coordinates and one-off sizing logic.

PPTKit should treat layout as a first-class layer so that:

- authoring APIs stay higher level
- layout behavior can be tested independently
- file-format code stays isolated
- future outputs do not require rewriting the same placement logic

## Normalize Before Export

PPTKit should not export directly from ad hoc authoring structures.

Instead, the system should normalize authoring input into a stable `Canonical Presentation IR` before preview, export, or parse-related transformations rely on it.

This gives the project:

- a stable internal contract
- clearer diagnostics
- better future support for parsing and round-tripping
- a cleaner place for validation and fallback logic

## Document Tradeoffs Honestly

Presentation systems have hard tradeoffs:

- editable output versus visual fidelity
- authoring ergonomics versus normalized rigor
- preview convenience versus package accuracy
- parser preservation versus strict normalization

The architecture should document these tradeoffs directly instead of hiding them behind vague feature language.

## Avoid Premature Product Assumptions

The architecture should leave room for:

- preview tooling
- browser-based editing
- additional importers and exporters
- future plugin hooks

But it should not require those systems to exist before the core document engine is useful.
