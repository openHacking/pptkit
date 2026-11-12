# Source script-font projection and scalar preservation observations

Recorded 2026-10-26 with LibreOfficeDev 26.8.0.0.alpha0 (`2c87e51eeaa2b413ff4ae097b2705eea1995d8e5`), macOS arm64. No PowerPoint/GUI interaction occurred in this batch.

[Runner](../../scripts/inspect-script-font-import.mjs), [independent Python oracle](../../scripts/research/oracles/script-font-import.py), [manifest](script-font-import-evidence.json). Reproduce after building with `rtk proxy node scripts/inspect-script-font-import.mjs`; optional research dependencies are LibreOffice and Python/PyMuPDF. Artifacts are `.tmp/script-font-import-native/run-CYKJmg`; the manifest records four source hashes, native hash, options and conversion logs.

The generated control contains text, linked text, shape text, two cells and notes, using explicit synthetic font names plus a theme-role/reset case. Parser imports with `script-families-v1` and V3, then changes all six admitted scalar text leaves over three source-save rounds. The final saved file is opened/resaved and exported to PDF by the named native producer.

| Independent source check | Result |
| --- | --- |
| Original + three saved packages | Four packages verified |
| Actual source-selected font declarations | 18 properties per package; all agree with source typeface/theme-slot values |
| Approved scalar text | Six leaves per package agree with approved snapshots |
| Changed parts after each edit round | Only slide1.xml and notesSlide1.xml |
| Bytes outside scalar text in those parts | Exact match with original |
| All other part payloads and part set | Exact match with original |
| Deliberately wrong selected ea family | Rejected by both font-value and non-character-byte checks in every round |

The font oracle uses ZIP/ElementTree and source paths, without the Parser resolver or Core selector. For this named generated fixture, a byte mask replaces only `a:t` character contents; everything outside those spans must match exactly. XML text checks also verify the approved leaf sequence. The deliberately wrong font changes only one lexical typeface value, so either check independently detects it. This fixture-specific mask does not establish general arbitrary-XML preservation or new write permission.

All native conversion processes exit successfully. Native resave splits five slide runs into fifteen; requested font names remain somewhere in the resaved source, which does not prove each native selection. Source-control and resaved PDFs retain owner1–owner5 labels and extractable Chinese/Arabic substrings. Notes owner6 is checked in source XML, not expected on slide PDF.

Both PNGs were actually viewed. Latin labels and Arabic are visible, while Chinese remains visually absent in this synthetic-font control. Actual PDF fonts are LinuxLibertineG and DejaVuSans, substituting the requested names. Native/PDF status remains `observed`, not a typography, glyph-coverage or source-identity pass.

This evidence supports explicit slot projection and the named scalar source-save boundary. Source font/style writes stay blocked. Actual installed-font coverage, complete language/supplemental/notes inheritance, shaping/bidi/vertical behavior and PowerPoint editing remain open.
