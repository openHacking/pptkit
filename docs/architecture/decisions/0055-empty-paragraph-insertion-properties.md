# ADR0055: Empty-paragraph insertion properties

Status: accepted for fresh output, 2026-10-26. Source style writes and general trailing insertion styles remain outside this contract.

## Decision

The shared Exporter paragraph serializer now copies the final empty run's character style onto `a:endParaRPr` when every run in that paragraph has zero characters. It also retains every original empty run. PowerPoint removes zero-character runs on save; the paragraph mark preserves the insertion style when no real run remains. For multiple differently styled empty runs, the final run supplies the single insertion style; native preservation of the individual zero-width identities is not promised.

Reuse the existing character-property serializer for fonts, language, size, flags, color and alpha. Script maps use the same empty-text/default-family materialization as the corresponding run and retain the existing warning. Actions remain on authored runs and are omitted from the newly written end properties, so this does not duplicate relationships or give subsequent inserted text an implicit link.

Nonempty paragraphs retain their existing ending serialization. There is no new public type, Core/IR field, Parser permission, dependency or host-font lookup. Text boxes, shape text, table cells and notes all use the same format-owned rule.

## Evidence and limits

[Focused regression](../../../packages/pptx-exporter/test/notes.test.mjs) first failed on the old default-language-only end mark, then passed after the fix. It checks multiple empty runs, all current style flags, separate font slots, alpha/language, action uniqueness, unchanged nonempty endings and the three drawing/cell owners. [Node/ESM/global regression](../../../scripts/empty-insertion-parity.test.mjs) removes zero-character source run nodes explicitly as an independent source control, then checks full notes/bytes through three fresh cycles and exact no-ops; it is not labeled a native operation.

[Native controls](../../testing/notes-insertion-styles.md) retain the complete end-property trees across three no-edit reopen rounds, followed by actual typing without formatting changes and three further rounds. Typed runs inherit those properties. The native app then omits the redundant end mark, which is recorded separately from the earlier empty stage. The previous real failed file remains unchanged and continues to demonstrate missing size/bold.

An empty native paragraph with only end properties has no bound character node. Existing source scalar saves therefore reject inserting new text there; no source-writer scope was widened. Empty-run actions and distinct zero-width run identities may still be removed by native applications. General trailing-empty runs after real text, authored insertion-position models, all scripts/fonts, inherited style resets and full notes-page semantics require separate acceptance.
