# Supplemental source font evidence

The [runner](../../scripts/inspect-supplemental-font-import.mjs) creates a named generated source fixture with explicit native theme references and seven supplemental records. It imports through `script-families-v2` and saves three rounds of admitted scalar edits across text boxes, shape text, table anchors and direct notes. Source construction is separate from author-map fresh output: the fixture replaces native slot references in generated XML before importing.

The [recorded manifest](supplemental-font-import-evidence.json) identifies hashes and local artifacts. An independent [Python stdlib ZIP/ElementTree oracle](../../scripts/research/oracles/supplemental-font-import.py) reads the actual four packages, without importing PPTKit. Each package has **29** selected map records and **six** scalar text leaves checked against approved snapshots. Actual map script/typeface attributes, effective selector/collection/dependency/target paths are verified. A deliberate wrong Hant typeface is rejected both by actual source-map checks and exact preservation checks.

Only slide1 and notesSlide1 payloads change after edits. The fixture-specific lexical mask admits scalar `a:t` character contents only; every other byte in those XML parts, every other part payload and the part-name set remain unchanged. The mask is not a general XML edit algorithm. Original plus all three stages retain hashes in the manifest; artifacts remain local under its recorded directory.

Run after building:

```sh
node scripts/inspect-supplemental-font-import.mjs
```

[Browser tests](../../scripts/supplemental-font-import-parity.test.mjs) separately check both dialects, Node/ESM/global agreement, actual XPath-selected source keys/values/dependencies/targets and another wrong-map negative control. Mixed-owner saves and warned fresh materialization are distinct checks. Existing source font/map/style edits remain blocked.

This source-save fixture itself has no native typography certification; its manifest's `native` and `fontCoverageAndShaping` remain `untested`. Later [separate installed-font and native controls](real-fonts.md) record PowerPoint/LibreOffice selection differences and three native save/reopen rounds. Neither those named observations nor earlier synthetic-font results certify editable theme behavior for every source fixture. The full gates remain open under [the leaf queue](../planning/advanced-capability-todos.md#脚本字体映射的接续任务).
