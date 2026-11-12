# Explicit script-font source and native observations

Recorded 2026-10-26 with LibreOfficeDev 26.8.0.0.alpha0 (`2c87e51eeaa2b413ff4ae097b2705eea1995d8e5`), macOS arm64. This is headless generation/open/resave/PDF evidence. Actual PowerPoint interaction, complete multilingual appearance and native font selection parity remain untested.

[Runner](../../scripts/inspect-script-fonts.mjs), [independent oracle](../../scripts/research/oracles/script-fonts.py), [manifest](script-fonts-evidence.json). Reproduce after building with `rtk proxy node scripts/inspect-script-fonts.mjs`; optional research dependencies are LibreOffice and Python/PyMuPDF. Artifacts are `.tmp/script-font-native/run-noO21w`; hashes and all conversion-stage logs are retained in the manifest.

The author fixture intentionally uses synthetic font names, with one theme-role/reset control. It contains text, linked text, shape text, two table cells and notes. Python ZIP/ElementTree independently compares all six source runs' text and Latin/ea/cs typefaces with the authored normalized fields. It checks theme-role suffixes and absent-slot fallback without calling the selector or writer. Deliberately changing one selected ea font to `Wrong Family` must fail. These source checks pass.

All native conversion stages exit successfully, but that does not certify typography:

- The slide's five source runs become fifteen native resaved runs, split around Latin/CJK/Arabic text. Some resaved font slots become absent. Native topology changes remain observations, not source identity preservation.
- Both PDFs contain extractable `Latin`, `中文` and `العربية` substrings. Actual spans use LinuxLibertineG and DejaVuSans at 18pt, substituting the synthetic names.
- Both PNGs were actually viewed. Latin and Arabic are visible; CJK is not visibly rendered despite extractable Chinese text. This control demonstrates why text extraction and successful conversion cannot substitute for glyph-coverage/appearance acceptance.

The runner labels native/PDF results `observed`, rather than `pass`. Source slot verification is a separate boolean. A caller-owned Layout font resolver receives exact selected segment text and can reject unavailable coverage before measurement; Core itself does not query installed fonts or package them.

Required next evidence includes actual installed font controls, glyph/ligature/mark/bidi and language variants, source/script-family reimport, font substitution decisions and native UI edits. No source-style write or complete script-font task is closed by this record.
