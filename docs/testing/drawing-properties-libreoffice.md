# Drawing property preservation: LibreOffice headless observation

This is a scoped observation for IMP-09.2.8 and the mixed source-preservation profile, not an application compatibility guarantee. The reproducible [runner](../../scripts/research/drawing-properties-libreoffice.mjs) generates a two-slide PPTKit deck and uses the Parser to change one text object's name, description, existing decorative value and text, plus the starting slide number. The resulting source-preserving candidate changes only `ppt/presentation.xml` and `ppt/slides/slide1.xml` payloads. An independent [Python ZIP/XML inventory](../../scripts/research/drawing-properties-oracle.py) inspects the candidate and each producer output.

LibreOfficeDev `26.8.0.0.alpha0 2c87e51eeaa2b413ff4ae097b2705eea1995d8e5`, macOS arm64, performed three successive headless PPTX conversions. In the edited source, the target has `name="Edited target 中文"`, `descr="Edited description & 中文"`, decorative `val="0"`, and `Edited text`; the presentation declares `firstSlideNum="17"`. Two other text objects on the first and second slides retain their names, descriptions, text and explicit decorative `val="1"`.

All three LibreOffice outputs retain the tested names, descriptions and text. The two unchanged objects retain decorative `val="1"`. Starting with the first conversion, the target's explicit `val="0"` extension and `firstSlideNum` declaration are absent. This records source syntax loss; it does not establish the displayed page number, accessible interpretation of absence, or behavior for other object types and producer versions. The conversions returned success, but headless output cannot establish whether an interactive application would show a repair dialog. No interactive editing, PowerPoint, visual comparison or assistive-technology behavior was tested.

[The retained evidence](drawing-properties-libreoffice-evidence.json) includes the exact application/version, commands' output, source and producer file paths and SHA-256 values, and per-round independent inventories. The files live under the recorded `.tmp/drawing-properties-libreoffice/` directory. Recheck available retained bytes and inventories with:

```sh
python3 -m unittest scripts/research/test_drawing_properties_oracle.py scripts/research/test_verify_drawing_properties_libreoffice.py
python3 scripts/research/verify-drawing-properties-libreoffice.py
```

To repeat the experiment with local `soffice`, `python3` and built workspace packages, run `node scripts/research/drawing-properties-libreoffice.mjs`. It creates a new retained directory and replaces the evidence manifest with the new run; review observed results before making compatibility claims. Parent backlog tasks remain open.
