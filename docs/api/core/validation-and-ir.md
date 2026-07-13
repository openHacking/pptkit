# Validation and Canonical IR

Validation is non-mutating and exhaustive. Normalization is the gate from ergonomic authoring state to detached, fully explicit Canonical Presentation IR v1.

## `validatePresentation(document)`

```ts
declare function validatePresentation(
  document: PresentationDocument,
): PresentationDiagnostic[];
```

The function always returns a diagnostic array and does not throw for validation findings.

```ts
interface PresentationDiagnostic {
  code: string;
  severity: "error" | "warning";
  message: string;
  path: string;
  slideId?: string;
  elementId?: string;
  assetId?: string;
  layoutId?: string;
}
```

`path` identifies the authoring field; identity fields make diagnostics stable when surrounding array positions change.

```ts
const diagnostics = validatePresentation(presentation);

for (const diagnostic of diagnostics) {
  console.error(
    diagnostic.severity,
    diagnostic.code,
    diagnostic.path,
    diagnostic.message,
  );
}
```

Validation accumulates independent findings, including invalid geometry/style ranges, missing assets/layouts/placeholders, missing action slides, connector references, duplicate identities, malformed text, invalid crops, invalid group coordinate sizes, and table span errors.

## `normalizePresentation(document)`

```ts
declare function normalizePresentation(
  document: PresentationDocument,
): NormalizedPresentation;
```

Normalization first validates the complete document. If any error diagnostics exist, it throws one `PresentationValidationError` and produces no partial IR.

```ts
import {
  normalizePresentation,
  PresentationValidationError,
} from "@pptkit/core";

try {
  const normalized = normalizePresentation(presentation);
  console.log(normalized.irVersion); // 1
} catch (error) {
  if (error instanceof PresentationValidationError) {
    for (const diagnostic of error.diagnostics) {
      console.error(diagnostic.code, diagnostic.path);
    }
  } else {
    throw error;
  }
}
```

## Normalization guarantees

- `irVersion` is exactly `1`.
- Every slide, layout, asset, element, and nested group child has a stable ID.
- All output arrays and objects are detached from authoring state.
- Metadata, size, theme, backgrounds, transforms, opacity, accessibility, text styles, paints, strokes, crops, table styles, and placeholder values are explicit.
- String text becomes structured paragraphs and runs plus `plainText` on text elements.
- Layout selection and background/style inheritance are resolved.
- Element order remains the drawing order.
- Exporters do not supply Core business defaults.

The complete schema and invariants are documented in [Canonical Presentation IR v1](../../architecture/canonical-ir-v1.md).

## Error ownership

Authoring operations immediately throw for local ownership errors such as duplicate IDs, unknown IDs, and invalid collection indices. Validation owns cross-document semantic errors. Exporters own recoverable environment failures such as an unreadable image source and report those as warnings.

## Mutability boundary

Normalized output does not share mutable references with the presentation. Changing a normalized object cannot update authoring state, and later authoring operations cannot retroactively alter an earlier normalized result.
