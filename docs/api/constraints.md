# Frame constraints

`ConstraintsComposition` adds bounded size and position relationships to `LayoutOptions.composition`. A rule specifies a container and an `elements` array, each with a unique `elementId` and one or more `x`, `y`, `width`, `height` constraints. Input state remains detached. Only direct children share a coordinate system; cross-container references are diagnosed rather than converted implicitly.

## Expressions and bounds

Each property accepts `{ value?, min?, max? }`, with at least one field. An expression is either a finite number or `{ constant?, terms }`. Each term is `{ source: "container" | { elementId }, metric, factor? }`; factor defaults to 1 and constant to 0. Metrics are `x`, `y`, `width`, `height`, `right`, `bottom`, `centerX`, `centerY`. Expressions sum the scaled metrics and constant; negative and zero factors are valid.

- With `value`, the equality must satisfy min/max. Conflicts reject the rule; they are not silently clamped.
- Without `value`, the incoming property is clamped to min/max. Omitted bounds are unbounded; dimensions must remain nonnegative.
- Bound expressions participate in dependency resolution. Inverted bounds, negative dimensions, nonfinite arithmetic and overflowing frame extents reject the rule. Only machine-rounding tolerance applies to equality comparisons.
- Omitted properties retain their incoming values. There is no soft-priority objective or general simultaneous linear-equation solver.

References read the constrained values from the **same rule**, even when the referenced item appears later. Unconstrained properties and unselected/hidden siblings use their incoming frames. Hidden targets are skipped unless `includeHidden` is true; references to them remain explicit inputs. Unknown targets and references produce diagnostics. A zero factor creates no dependency, but its referenced ID must still exist.

The container source is the padded `area` (default slide dimensions or group-local coordinate size). Constants and element references use the same container-local coordinates. Rotation/flips apply later; only `bounds: "frame"` is accepted. For visual alignment, use a separate alignment/grid rule.

## Dependency and precedence contract

Within a rule, each constrained property is a graph node. Derived metrics such as right depend on both x and width. The solver topologically evaluates the graph without recursion. Any reference cycle, including a mathematically redundant equality, is rejected with an actual dependency chain. Cycles are not solved by iteration or hidden tie-breaking. Item order does not change results or cycle reporting.

Across rules, composition order is explicit: later rules consume and may override earlier geometry. Intrinsic/table sizing precedes composition; connector/image/chart resolution and diagnostics follow it. To size objects before placing them in a grid, put constraints before the grid. To anchor a grid item afterward, put constraints after it. No rules are automatically re-run.

Changing text width after intrinsic measurement may require new text measurement; the constraint stage does not recompute an intrinsic height. Enable text overflow diagnostics. Constraints can also change image allocations before contain/cover and chart frames before chart geometry. Native Office can still reflow text; geometry agreement is not font/layout parity.

## Element behavior

Text, shapes, images, charts and groups support constrained dimensions. A group resize preserves `coordinateSize` and intentionally scales its children; it does not fit or rebase content. Use [intrinsic sizing](intrinsic-sizing.md) to fit group contents while preserving child world geometry.

Tables can translate, but dimension changes are rejected: change rows/columns or use table sizing to retain their native grid contract. Free connectors can translate all points/routes; resizing them is rejected. Anchored connectors follow constrained target shapes after composition and cannot themselves be selected or referenced before final connector resolution. Explicit connector route points remain authored; this is not automatic route solving.

## Atomic failure and diagnostics

Malformed configuration throws before composition starts. Every rule computes candidate frames before mutation; a semantic failure discards all of that rule's changes, retains earlier successful rules, and permits later rules to run. The error path identifies `options.composition.<index>`; messages identify the property/reference or dependency cycle.

In addition to common composition warnings, codes are:

| Code | Meaning |
| --- | --- |
| `composition-constraint-reference-missing` | Referenced ID is not a direct child. |
| `composition-constraint-unresolved-reference` | A reference requires unresolved anchored connector geometry. |
| `composition-constraint-cycle` | Property dependencies form a cycle. |
| `composition-constraint-unsatisfiable` | Equality and bounds conflict, bounds are inverted or a size is negative. |
| `composition-constraint-invalid` | Arithmetic or frame extents are nonfinite. |
| `composition-constraint-resize-unsupported` | Table/connector dimensions would change independently of their structural geometry. |

Diagnostics are forwarded by exporter and renderer. The caller decides whether warnings block delivery.

## Checked example

<!-- doc-test: docs/examples/constraints.ts -->
```ts
import { createPresentation } from "@pptkit/core";
import { resolveLayout, type ConstraintsComposition } from "@pptkit/layout";

export function createConstrainedCards() {
  const document = createPresentation({ size: { width: 600, height: 300 } });
  const slide = document.addSlide({ id: "cards" });
  for (const id of ["left", "right"]) slide.addElement({ type: "shape", id, shape: "rect", box: { x: 0, y: 30, width: 100, height: 80 } });
  const rule: ConstraintsComposition = {
    type: "constraints", container: { kind: "slide", id: slide.id }, padding: 20,
    elements: [
      { elementId: "left", x: { value: 20 }, width: { value: 180, min: 100, max: 200 } },
      { elementId: "right", x: { value: { terms: [{ source: { elementId: "left" }, metric: "right" }], constant: 20 } }, width: { value: { terms: [{ source: { elementId: "left" }, metric: "width" }] } } },
    ],
  };
  return resolveLayout(document, { composition: [rule] });
}
```

## Verification

[Constraint tests](../../packages/layout/test/constraints.test.mjs) cover affine/bounded relationships, reversed declaration order, cycle chains, atomic failures, hidden input, nested groups, connector/table limits, grid precedence, a 12,000-node chain and 100 seeded branching DAGs compared with independent forward evaluation. [Cross-package parity](../../scripts/constraints-parity.test.mjs) compares full PPTX bytes and SVG with manually authored frames, and verifies unchanged outputs plus forwarded warnings on cycles.
