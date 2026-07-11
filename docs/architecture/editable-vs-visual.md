# Editable vs Visual

`editable` and `visual` should be treated as architecture terms, not marketing labels.

They describe a real system tradeoff: presentation content often cannot be optimized for maximum editability and maximum visual fidelity at the same time.

## The Constraint

Some content maps well to native presentation objects:

- text blocks
- simple shapes
- standard images
- basic grouped structures

Other content is much harder to preserve as fully editable native objects without losing appearance:

- complex SVG effects
- browser-like layout behavior
- transforms that do not map cleanly to DrawingML
- advanced styling or unsupported visual effects

## Strategy Modes

PPTKit should use `editable` and `visual` as strategy terminology that future import, render, and export layers may rely on.

### `editable` Strategy

Favor native presentation objects wherever possible.

This strategy should prefer:

- editable text
- editable shapes
- native grouping where feasible
- structure that survives post-export authoring

### `visual` Strategy

Favor closer visual fidelity when native object mapping would degrade the result too much.

This strategy may rely more heavily on:

- SVG-backed output
- rasterized fallbacks
- grouped visual approximations
- explicit loss of some object-level editability

## Mixed Fallback Behavior

Some workflows will need mixed behavior rather than a single pure mode.

For example:

- preserve editable text but flatten one unsupported effect region
- export most elements as native objects but keep one visual cluster as SVG-backed output
- parse a deck into normalized editable structures while preserving unknown content for later export

The architecture should allow these mixed outcomes instead of pretending one global mode always solves everything.

## Documentation Rule

PPTKit should prefer explicit degradation over implicit false promises.

That means the system should be able to say:

- this part stays editable
- this part stays visually closer
- this part was preserved but not fully understood

That is more trustworthy than claiming both properties can always reach 100%.
