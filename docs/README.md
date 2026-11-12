# PPTKit Documentation

The documentation is organized by audience and task. Current behavior is written in the present tense; parser and other unimplemented systems are labeled as future architecture.

## Start using PPTKit

- [Install PPTKit](getting-started/installation.md)
- [Quick Start](getting-started/quick-start.md)
- [Use PPTKit Directly in a Browser](guides/browser-cdn.md)
- [Create Your First Deck](guides/create-your-first-deck.md)
- [PPTKit Presentation product](https://github.com/openHacking/pptkit-presentation)

## Use the API

- [API index](api/README.md)
- [`@pptkit/core`](api/core.md)
- [`@pptkit/layout`](api/layout.md)
- [`@pptkit/pptx-exporter`](api/pptx-exporter.md)
- [`@pptkit/svg-renderer`](api/svg-renderer.md)
- [Package overview](api/package-overview.md)
- [Placeholder matching and dimension import](api/placeholder-inheritance.md)
- [Text style inheritance in content import](api/text-style-inheritance.md)
- [Text color transforms and alpha](api/text-color-transforms.md)
- [OPC source comparison](api/package-comparison.md)
- [OPC source relationship closure](api/preservation-closure.md)
- [Direct text-run preservation save](api/text-edit-save.md)
- [Shape-text and speaker-notes preservation save](api/text-owner-save.md)
- [Table-anchor character preservation save](api/table-text-save.md)
- [Node.js atomic file import and output](api/file-save.md)

## Understand the system

- [Concepts](concepts/README.md)
- [Architecture overview](architecture/overview.md)
- [Core authoring model](architecture/core-authoring-model.md)
- [Canonical Presentation IR v2](architecture/canonical-ir-v2.md)
- [Package boundaries](architecture/package-boundaries.md)
- [Rendering pipeline](architecture/rendering-pipeline.md)

## Contribute

- [Developer workflow](guides/developer-workflow.md)
- [Example workbench](guides/example-workbench.md)
- [Testing strategy](architecture/testing-strategy.md)
- [Capability and public contract audit](planning/capability-matrix.md)
- [Named regression corpus](testing/fixture-corpus.md)
- [Diagram cache consistency protocol](testing/smartart-cache-consistency.md)
- [Cycle layout research protocol](testing/smartart-cycle-layout.md)
- [SmartArt source-binding protocol](testing/smartart-source-bindings.md)
- [SmartArt identity-state research protocol](testing/smartart-identity-state.md)
- [Generated SmartArt native-edit protocol](testing/smartart-generated-native-editing.md)
- [Native SmartArt geometry baseline](testing/smartart-native-geometry.md)
- [Application compatibility evidence](testing/application-compatibility.md)
- [Performance baselines](performance/README.md)
- [Source organization](architecture/source-organization.md)
- [Release strategy](migration/release-strategy.md)

## Future areas

The [Roadmap](../ROADMAP.md) defines capability milestones. The [Implementation Backlog](planning/implementation-backlog.md) tracks concrete tasks, dependencies, ownership, and acceptance criteria for layout, authoring, import/round-trip, SmartArt, media, motion, tooling, and stable releases. Planned capabilities are not presented as current APIs.

The [Advanced Capability TODOs](planning/advanced-capability-todos.md) break these larger tasks into individually verifiable slices and a recommended execution queue.

The [Next Delivery Plan](planning/next-delivery-plan.md) maps the whole slide capability surface to task IDs, orders the next bounded work packages, and defines the samples and evidence needed to accept them.

Its [capability delivery order](planning/next-delivery-plan.md#规划入口完善幻灯片能力的领取顺序) provides the current overview. The [SmartArt output/layout cards](planning/advanced-capability-todos.md#下一步执行卡smartart-最小原生输出与布局准入) separate differential experiments and native display/editing gates. [Minimum generation research](architecture/decisions/0023-native-smartart-generation-research.md) records actual producer observations and the correction of fragmented-word detection.

Start with its [current batch entry](planning/next-delivery-plan.md#本次排期先领取哪一批): accept the workspace text-save candidate, close common-object semantics, expand source-preserving edits, then structural operations and the individually gated SmartArt/media/motion tracks. The [object-save TODOs](planning/advanced-capability-todos.md#编辑保存的对象覆盖补齐) explicitly separate tables, charts, notes/properties and actions from ordinary text-box saves. These are planned work and pending acceptance, not completion claims.

For the next executable tasks, start with its evidence-closure and task-card tables. The TODO checklist includes separate acceptance slices for multi-master/theme work, package comparison and save transactions, SmartArt feasibility/native editing, and animation triggers/time evaluation.

Input policies, background/master artwork, template rebinding and fields, notes-page layout, native section editing, and custom properties also have individual task cards in the same checklist. Their planned fixtures and entry points are in the delivery plan.

The checklist also has separate tasks for text layout/line breaking, geometry/effects, image format policies, table editing, SmartArt cache consistency, slideshow interruption/re-entry, mixed-content regression and inspection reports. Use the delivery plan's starter cards to choose a bounded next contribution.

- [Experimental slide Fade transitions](api/slide-transitions.md)

[Direct SmartArt text observations](api/diagram-text-comparison.md) add explicit character diagnostics; [controlled cache evidence](testing/smartart-stale-cache.md) keeps native freshness/admission separate.

[Read-only CLI inspect/compare](api/cli-inspection.md) provides local diagnostics and versioned JSON over public Parser APIs.

[Versioned import capability observations](api/import-capabilities.md) expose complete operation vectors with explicit missing/conflicting evidence and task/artifact hints.

- [Chart numeric source-save/native evidence](testing/chart-value-save.md): independent cache/XLSX and unchanged-payload checks over three rounds, named native resaves and explicit interactive/eligibility limitations.

- [Direct table borders](api/table-borders.md) — authoring/import, shared Layout edges and native evidence boundaries.

- [Table region styles](api/table-style.md) — authoring defaults, bands/corners, effective values and native materialization boundaries.

- [Source table styles and themes](api/table-style-import.md) — effective fields and source provenance, separate from authoring/native reuse or source style writes.
