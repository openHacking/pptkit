# Table text appearance observations

Recorded 2026-10-26. Application: LibreOfficeDev 26.8.0.0.alpha0, commit `2c87e51eeaa2b413ff4ae097b2705eea1995d8e5`, macOS arm64. No GUI edits or PowerPoint playback/editing occurred in this batch.

[Runner](../../scripts/inspect-table-text-appearance.mjs), [Python oracle](../../scripts/research/oracles/table-text-appearance.py), [manifest](table-text-appearance-evidence.json). Reproduce after building with `rtk proxy node scripts/inspect-table-text-appearance.mjs`; optional research dependencies are LibreOffice and Python/PyMuPDF. Local artifacts are `.tmp/table-text-native/run-wmF1hx`; the manifest records source/resave hashes and per-stage logs/failures.

Four valid generated source packages each contain a 2×2 table with source table-style definitions and stripped direct run font/color declarations (local control reintroduces explicit color). The independent Python ZIP/ElementTree oracle resolves selected source paths, theme colors and font dependencies for eight properties per scene. It mutates the selected source color in the actual inspected tree and requires rejection. All four source checks pass.

| Control | Expected color | Source PDF | Resaved PDF |
| --- | --- | --- | --- |
| Reference-only minor font/accent2 | CA6420 | FFFFFF: partial | 000000: partial |
| Explicit region color after reference | 102030 | Pass, four labels/colors | Pass, four labels/colors |
| Local run color after reference | 445566 | Pass, four labels/colors | Pass, four labels/colors |
| Equal lt/ea/cs collection + explicit color | CA6420 | Pass, four labels/colors | Pass, four labels/colors |

Native conversion stages completed without process failures. PDF span extraction records font `LiberationSans` and approximately 15.987pt for requested Arial/16pt. Color pass only checks opaque RGB and unique text labels; fonts and font sizes are observations, not parity passes. Alpha, mixed-script glyphs/shaping, line breaking and whole-slide geometry remain unverified here.

The reference source-control PNG was actually viewed: white text on a white background makes the page appear blank despite extractable labels. The explicit-color saved PNG visibly contains all four labels. PDF text existence alone would therefore have concealed an actual appearance failure. Keep both source and resaved observations; changing the canonical expected accent2 to match either producer result would hide the discrepancy.

Full default appearance and PowerPoint controls, script segmentation/selection, glyph coverage, font substitution and interactive table style/theme changes remain open. This experiment does not establish source-style write permission.
