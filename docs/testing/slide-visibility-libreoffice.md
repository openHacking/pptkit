# Source-preserved slide visibility: LibreOffice headless observation

The [experiment](../../scripts/research/slide-visibility-libreoffice.mjs) generates a three-slide PPTKit presentation with distinct labels, then uses the explicit slide-visibility preservation profile to hide only the second slide. The unmodified file is the control. Independent ZIP payload comparison confirms that the edit changes only `ppt/slides/slide2.xml`; the control and edited inputs have the same package entry set. Each input is exported to PDF, then independently resaved and reopened as PPTX three times in LibreOfficeDev, with another PDF export after each resave.

On macOS arm64 using LibreOfficeDev `26.8.0.0.alpha0 2c87e51eeaa2b413ff4ae097b2705eea1995d8e5`, the [retained record](slide-visibility-libreoffice-evidence.json) shows all three slides in the PPTX inventory for every round. The control has no explicit `show` declaration and its four PDFs each have three extracted pages in the order `Visible Alpha`, `Target Beta`, `Visible Gamma`. The edited file and all three producer outputs retain `show="0"` on the second slide; their four PDFs each have two extracted pages, `Visible Alpha` and `Visible Gamma`. This is one named sample and producer build. Extracted PDF text and page count do not prove slideshow playback, page legibility, interactive editing, absence of repair prompts or PowerPoint behavior.

The separate [Python oracle](../../scripts/research/slide-visibility-oracle.py) resolves presentation relationships in presentation order, reads slide XML and uses PyMuPDF to inspect PDF page counts, dimensions and text. Its tests cover relationship order, absent versus explicit `show`, external/wrong relationship types and PDF extraction. The [retained-file verifier](../../scripts/research/verify-slide-visibility-libreoffice.py) rehashes all 16 recorded artifacts and re-runs the oracle; negative tests reject modified hashes, inventories, paths and round order, including under `python -O`. Retained files remain under the manifest's `.tmp/slide-visibility-libreoffice/` paths. Recheck them with:

```sh
python3 -m unittest scripts/research/test_slide_visibility_oracle.py scripts/research/test_verify_slide_visibility_libreoffice.py
python3 scripts/research/verify-slide-visibility-libreoffice.py
```

With local `soffice`, Python/PyMuPDF and built workspace packages, `node scripts/research/slide-visibility-libreoffice.mjs` runs a new experiment and replaces the evidence manifest. Review the resulting observations; it does not check a universal producer contract. Broader native and parent backlog gates remain open.
