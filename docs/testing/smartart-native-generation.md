# Minimum native SmartArt generation experiments

This protocol advances ADV-02.1.3 and ADV-02.2.1. It does not admit a public
SmartArt writer. Contract/ownership: [ADR0023](../architecture/decisions/0023-native-smartart-generation-research.md).

## Automated structural coverage

After `pnpm build`, run:

```sh
rtk proxy node --test scripts/native-diagram-research.test.mjs scripts/diagram-content-research.test.mjs
```

The 15 native research groups and seven content-model groups use owned synthetic
XML/component roots. No downloaded component is required. The native groups are
also wired into `pnpm test:import` and the full `pnpm test` command. Browser DOM
checks exercise all three families, both cache policies and both presentation
modes; unrelated text edits preserve component bytes over three rounds.

## Optional offline producer experiment

Keep source files and transformed artifacts local until licence admission.
Use the existing hash-pinned corpus described by
[source-editing evidence](smartart-native-editing.md); the runner requires its
`provenance.json` and refuses mismatched source hashes. No script downloads inputs.

```sh
rtk proxy node scripts/generate-smartart-native-candidates.mjs --presentation
rtk proxy node scripts/generate-smartart-native-candidates.mjs
```

Each prints a **new** artifact directory and creates its own `manifest.json`.
The first writes [presentation evidence](smartart-native-presentation-candidate-evidence.json),
the second [content-only evidence](smartart-native-candidate-evidence.json).
They cover empty/single/standard/long cases and absent/empty cache, with independent
ZIP/XML checks, exact no-op and placeholder-preview observations. Empty cache
does not provide a preview image or demonstrate geometry generation.

Use each printed manifest path and an explicit local `soffice` executable:

```sh
rtk proxy node scripts/inspect-smartart-native-candidates.mjs /absolute/run/manifest.json /absolute/soffice
rtk proxy node scripts/inspect-smartart-native-differential.mjs /absolute/presentation-run/manifest.json /absolute/source-controls/manifest.json /absolute/soffice
```

The producer runner saves and reopens the **standard** cases only; do not extend
its conclusion to empty/single/long. The optional source-controls manifest has
the same records schema, original file hashes and separately projected authored
nodes. The Cycle differential requires one original source control and creates
five local variants, each directly rendered and then saved/reopened/rendered.
The resulting [differential evidence](smartart-native-differential-evidence.json)
records input/output hashes, extracted text, logs and picture paths/hashes.
These are controlled research transformations, not source-preserving API saves.

## Results and interpretation

| Observation | Named runtime result | What remains unproven |
| --- | --- | --- |
| Six standard presentation-mode outputs | Data inventory retained; all authored word-spacing markers extracted after save/reopen | PowerPoint display, editing, geometry fidelity, cache freshness |
| Six content-only controls | Ordinary control text extracted; authored diagram markers unmatched | A native layout path from content alone |
| Cycle custom-text flag only | Words fragment in extraction both directly and after reopen | Universal behavior of customization flags in other apps/layouts |
| Cycle unlocked/GUID/crossed-shell variants | Word-spacing markers match directly and after reopen | All layout/identity policies or visual equivalence |

`allTextMarkersInPdf` uses whitespace-normalized whole marker strings.
`allWhitespaceInsensitiveMarkersInPdf` additionally ignores whitespace inside
words; it can expose extraction fragmentation. Both are substring observations,
not per-node occurrence proofs. Empty/whitespace-only markers are not valid
evidence. Preserve raw text, inspect pictures, and do not infer missing glyphs
from a word-spacing mismatch. Neither PDF matching nor an ordinary control label
establishes distinct visible nodes, correct connectors or native editability.

Review each generated picture for node/edge count, clipped/overlapping labels,
direction, text alignment and unexpected shape conversion. Record the exact
application/OS/fonts; missing environments stay untested. Final family admission
requires actual SmartArt recognition, insert/delete/reorder, text/theme edits,
re-layout and three native rounds on generated candidates. No full task is closed
by these experiments.

## Explicit GUID specimen controls

```sh
rtk proxy node scripts/generate-smartart-native-candidates.mjs --presentation --guid
rtk proxy node scripts/inspect-smartart-native-candidates.mjs /absolute/guid-run/manifest.json /absolute/soffice
```

`--guid` requires `--presentation`; default integer controls remain unchanged. GUID manifests produce separate [candidate](smartart-native-guid-candidate-evidence.json) and [producer](smartart-native-guid-producer-evidence.json) evidence. All 24 structural cases and six standard headless producer cases passed their scoped checks. Portable tests cover both identity policies, including independent DOM and unrelated text saves. [PowerPoint generated-edit evidence](smartart-generated-native-editing.md) now covers the standard absent-cache subset; the earlier headless results table does not describe this separate UI experiment. Traversal-derived GUIDs are specimen-local and are not managed author identities.


## Stateful identity experiment

The [explicit identity-state protocol](smartart-identity-state.md) adds stable semantic-key reservations across generated structural edits. It requires explicit GUID policy and caller-managed state commits. Stateless defaults remain unchanged. Stateful output is tested independently across structural rounds; native application admission for this option remains untested.
