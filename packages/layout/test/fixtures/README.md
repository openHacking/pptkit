# Measurement font fixtures

`noto-sc-measurement.otf` is a test-only subset of [Noto Sans SC Regular, Sans2.004](https://github.com/notofonts/noto-cjk/blob/Sans2.004/Sans/SubsetOTF/SC/NotoSansSC-Regular.otf), under the included [SIL Open Font License](NOTO-SANS-LICENSE). Source SHA-256: `faa6c9df652116dde789d351359f3d7e5d2285a2b2a1f04a2d7244df706d5ea9`.

Generated using fontTools 4.61.1 with retained names/languages, no timestamp recalculation, ASCII U+0020–007E, U+2022, U+200B and characters in `中文测量换行混排字体测试`. Family/full/PostScript/typographic family names were changed to PPTKit Measurement CJK (PostScript: PPTKitMeasurementCJK). The actual cmap is recorded in `cjk-coverage.json`. This small subset is a pinned test asset, not a general-purpose CJK font.

Browser tests also reuse the existing licensed [Roboto Mono fixture](../../../svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2). Tests load bytes explicitly before measurement and block external requests. These fixtures prove browser font-backed behavior, not native Office parity.

`noto-kr-tab-measurement.otf` is a test-only subset of [Noto Sans KR Regular, Sans2.004](https://github.com/notofonts/noto-cjk/blob/Sans2.004/Sans/SubsetOTF/KR/NotoSansKR-Regular.otf), under the same included SIL Open Font License. The source SHA-256 is `69975a0ac8472717870aefeab0a4d52739308d90856b9955313b2ad5e0148d68`; the fixture SHA-256 is `6946b76e67a54e49c2edf8ff9d2b194e366e2dad5703f75202322884c9b79feb`. [The generator](../../../../scripts/research/build-tab-font.py) checks the source hash, uses fontTools 4.61.1, subsets the listed ASCII/Han/kana/Hangul samples and renames the modified family to PPTKit Tab Measurement. The exact test coverage is intentionally small and does not imply general script coverage.
