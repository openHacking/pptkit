# Theme-matrix source and native observations

Recorded 2026-10-26 with LibreOfficeDev 26.8.0.0.alpha0 (`2c87e51eeaa2b413ff4ae097b2705eea1995d8e5`), macOS arm64. This is headless import/resave/PDF evidence; no PowerPoint interaction was performed in this batch.

[Runner](../../scripts/inspect-theme-matrix.mjs), [independent Python oracle](../../scripts/research/oracles/theme-matrix.py), and [machine-readable manifest](theme-matrix-evidence.json). Local artifacts: `.tmp/theme-matrix-native/run-73TbTM`; the manifest retains input/output hashes and per-stage application failures. Reproduce with `rtk proxy node scripts/inspect-theme-matrix.mjs` after building; LibreOffice, Python and PyMuPDF are optional native-research dependencies, not runtime package requirements.

Each valid generated source contains a 4×4 table, real source style references and theme matrices. Python ZIP/ElementTree independently checks 80 selected properties per scene against actual source slots and continuous RGB/alpha programs. Deliberately changing a selected index must fail the source oracle. This source verification passed for all five scenes; `independent.verified` means source/oracle execution, not native parity.

| Scene | Native direct properties | Source-control / resaved PDF |
| --- | --- | --- |
| Placeholder fill1 / line1 | Fill/color agree; 1pt line becomes about .96378pt: partial | All 16 labels and fill pixels pass in both |
| Background fill1001 / line3 | Expected ABCDEF fill, observed 3579BD; line also differs: partial | Text passes; fill pixels fail in both |
| No-fill fill1000 / line1 | Unavailable: application exits 1 with `Unspecified Application Error` | Source PDF also fails; neither PDF available |
| No-line fill1 / line0 | NoFill retained with stored width1 instead of projected zero: partial | All 16 labels and fill pixels pass in both; this does not test line pixels |
| Ordered transforms | Expected 080808/alpha .2, observed 121212/alpha .5: partial | Text passes; expected white-composited 206 versus observed 137 in both |

PDF fill sampling uses 4× rasterization, one interior sample per cell and tolerance two per RGB channel; it is not a whole-slide visual or typography comparison. Placeholder/background source-control PNGs and transform resaved PNG were visually inspected. The blue background and gray transform controls visibly retain the recorded fill discrepancies. Border pixel fidelity is not established by interior fill samples.

Each scene runs with an isolated application profile. The no-fill failure includes Fontconfig warnings, but the cause has not been established; it is neither a successful compatibility result nor an automatic approval rejection. Other scenes completed despite the warnings. Index/transform expectations were not changed to obtain producer agreement.

Remaining gates: independent PowerPoint controls for sentinels/background indices and modifier order; actual style/theme/insert-row edits and save/reopen; complete line-width/alpha/transform/typography fidelity; merged/RTL/overlapping regions; source-style writer admission. Existing TODO parents remain open.
