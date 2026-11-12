# Native pagination row-style observation

PowerPoint for Mac 16.104 (16.104.25121423), macOS 26.5, opened one five-page PPTKit export and one separately patched research control on 2026-11-11. Each contains an unsplit five-row table followed by four two-row fragments, with a repeated header. All tables reference Medium 2 Accent 2 with first-row, last-row and banded-row switches enabled. PowerPoint's Export command produced the two PDFs on first open; no native edit, save/reopen or other application was tested.

The PPTKit output contains direct cell properties. At three text-free sample positions, the direct-output PDF was white on every sampled row; this sample does not reveal the native style. In the **research control only**, each table-cell `tcPr` was replaced by an empty `tcPr`, leaving all other ZIP parts and slide bytes unchanged. PowerPoint then rendered the source table's five sampled row colors as `ED7D31`, `F8D7CD`, `FCECE8`, `F8D7CD`, `ED7D31`. Each two-row continuation sampled `ED7D31` for both rows: the native last-row region was reapplied to every fragment. This control isolates a possible native style interaction; it is not fresh PPTKit output and does not prove all presets or consumers behave the same way.

Local artifacts are in `.tmp/pagination-native-observation/` and are not distributed fixtures:

| Artifact | SHA-256 |
| --- | --- |
| `source.pptx` | `15a9888aafcef0ae70c789b86de766777378e4d5ff3d0e181732dc7e78ff4f28` |
| `style-probe.pptx` | `f70403d4bb846ad37d7db639945b1fabdb2d6af45d005b252142e0c41221a9c3` |
| `pptkit-pagination-probe-20261111.pdf` | `6e397ae4593fdd8500a7ce5e25d15374e4c1c9a5e9d33a62c4e20300cdb75be6` |
| `style-probe.pdf` | `27d8095f9a3077d05f66347335e66bf05efd64570254818a2be01205640131c4` |

Recreate the source in a new empty directory with `node scripts/prepare-pagination-native-probe.mjs DIRECTORY` after building the workspace. The source was reproduced byte-for-byte. The local control was made by replacing each slide's `a:tcPr` with an empty element; the [independent verifier](../../scripts/research/inspect-pagination-native.py) checks the exact patch, every other ZIP part, native switches, page text and measured PDF pixels. Run it with the retained artifact directory as its argument. Its PyMuPDF dependency and the native PDFs are required; generating a new PDF is an actual PowerPoint operation, not an automated substitute.

The [verifier test](../../scripts/pagination-native-observation.test.mjs) confirms that changing a sampled PDF pixel or a native row switch is rejected. It skips if the retained local PDFs are absent.

## Opt-in source-boundary policy

A second first-open PowerPoint 16.104 observation uses the new `source-boundaries` policy, again with a separately patched control that empties all `tcPr` elements. Its first three continuation body samples changed from `ED7D31` to `F8D7CD`; the actual final row remains `ED7D31`. The source's second body row is `FCECE8`, so that continuation still differs: native bands restart. This confirms only this control's sampled first-open fill, not the appearance of the directly formatted PPTKit output or native editing and reopening.

Retained local artifacts under `.tmp/pagination-native-observation/source-boundaries/` have SHA-256 values: source PPTX `b8aa0e1ba2ee7c6cb71f8b03f253ae90fba0319326e277524885b144c5f6332d`, patched PPTX `caa3fabbc972329e30d7394cbfab0a28eb1d97f4311af446fe93e400dc269375`, and PowerPoint PDF `5f1d45fa244db3a69760cb28310763c7577fa9df8a682f5feec6a18d5372f86a`. Recreate the generated PPTX in a new directory by passing `source-boundaries` as the second argument to the generator. The [independent verifier](../../scripts/research/inspect-pagination-boundaries.py) checks every non-slide ZIP part, the exact cell-property-only patch, native flags, page text and three sampled pixels per row. Its [negative controls](../../scripts/pagination-native-observation.test.mjs) reject a modified switch and sampled color.

The actual, unpatched PPTKit output was also opened in PowerPoint 16.104, saved as `native-saved.pptx`, closed and reopened once. The saved package retains the five page-specific text lists, table frame/row geometry and first-row, banding and last-row switches. PDFs exported **after the save** and after reopening contain the same page text and identical five-page rasters. This is one save and one reopen, with no table edit; it is not a comparison to a first-open PDF. Saved PPTX SHA-256 is `bc57010c92492206117aff62c9a221b84cd22607cf1494f52ebb11ab59df0ad5`, post-save PDF is `a5a3b0102f8d5a4a54b4d7fc12d3f6ef8da00f66742b28b5b74839951a69b9c7`, and reopened PDF is `620184a605ba2506c6ea884e41a80a3baace420506dc305670dc3136174504d7`. The [independent package/PDF verifier](../../scripts/research/inspect-pagination-reopen.py) and its negative controls check the retained local artifacts; altered text, row flags and reopened pixels fail.

Native banding remains a warning. Default native switches, other row counts and spans, edited styles, interactive edits, screenshots, multiple native save/reopen rounds and cross-platform behavior remain unverified. The AUT-05 and pagination-native compatibility gates stay open.
