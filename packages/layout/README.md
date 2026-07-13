# `@pptkit/layout`

Layout resolution primitives for PPTKit.

## Current Status

`resolveNormalizedLayout` consumes normalized core IR and returns an export-ready layout IR. `resolveLayout` remains a compatibility adapter for authoring documents and performs normalization once before delegating to the normalized entry point. The result preserves slide and element ordering, geometry, and normalized styles while keeping file-system and PPTX/XML concerns out of this package.

The same layout result is intended to be reusable by preview and export paths.

Internally, layout conversion is kept separate from future constraint, measurement, and pagination engines. This package never reads assets or writes output files.
