# Document metadata projection contract

This table describes Parser's current structure and content imports. Inspection retains XML-normalized source values and identities; it is not a byte-exact lexical record or an assertion that every value is represented in Core. Invalid structures fail import, while valid but unsupported values remain in inspection and receive an unsupported finding. Unchanged no-op saves retain the original package bytes. Source-preserving edits require their separate opt-in profile and candidate validation; a Core projection alone does not grant edit-save permission.

| Source | Core projection | Boundary |
| --- | --- | --- |
| Core `title`, `creator`, `subject`, `description`, `language` | `title`, `author`, `subject`, `description`, `language` | Decoded strings, including empty strings. Missing fields remain absent. |
| Core `keywords` | One `keywords` entry containing aggregate text | Segment boundaries and language annotations are not represented; mixed or annotated content additionally receives an unsupported finding. An empty element becomes `[""]`. |
| Core `revision` | `revision` | Canonical nonnegative integer text within JavaScript's safe-integer range; other values stay in inspection. |
| Core `created`, `modified` | `createdAt`, `modifiedAt` | Valid explicitly zoned timestamps exactly representable in milliseconds and supported four-digit UTC year; converted to canonical UTC. Unzoned, unknown-offset, invalid or sub-millisecond values remain source-only. |
| Other core properties | None | Retained in inspection and individually reported unsupported. |
| Extended `Company` | `company` | Decoded scalar string, including empty. Other extended properties remain in inspection and are individually reported unsupported. |
| Named, unlinked custom `lpstr`, `lpwstr`, `bstr` | String custom property | Literal escape decoding must succeed; empty string values are allowed. |
| Named, unlinked custom `bool` | Boolean custom property | Only exact source text `true`, `false`, `1` or `0` maps, even if inspection accepts surrounding whitespace. |
| Named, unlinked custom integer variants | Integer custom property | Canonical integer text in both the original variant range and Core's int32 range; other valid source integers remain source-only. |
| Named, unlinked custom `r8`, `r4`, `decimal` | Real custom property | Finite canonical number text; `r4` additionally requires exact binary32 representation and `decimal` exact mathematical binary64 representation. |
| Named, unlinked custom `date`, `filetime` | Date custom property | The same explicitly zoned, exact-millisecond and four-digit UTC boundary as core dates. |
| Other custom values, including `empty`, `null` and complex variants | None | Retained in inspection and reported unsupported. Names absent, empty, linked, undecodable or ambiguous after decoding also prevent projection. |
| Custom `cy` | None by default | An explicit read-only opt-in can project an exact binary64 subset; see [opaque custom-property inventory](opaque-custom-properties.md). |

Custom source identity is format GUID plus PID; it is not a Core identifier. Duplicate source names or identities are rejected during inspection. A decoded-name collision with any named source property prevents projection, including when the other value is opaque or linked. Names whose escapes cannot be decoded remain source-only and are not assumed to match a decoded name. Package, XML and report limits apply, and Strict and Transitional namespaces must match the presentation dialect.

The [projection contract tests](../../packages/pptx-parser/test/metadata-projection-contract.test.mjs) exercise missing versus empty fields, mixed-language keywords, out-of-range revision, date precision, source-only values and byte-identical no-op saves in both dialects. Detailed type boundaries and failure cases are covered by the [metadata inspection tests](../../packages/pptx-parser/test/metadata.test.mjs), [decimal tests](../../packages/pptx-parser/test/custom-property-decimal.test.mjs) and [currency projection tests](../../packages/pptx-parser/test/custom-property-currency-projection.test.mjs).
