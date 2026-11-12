# Explicit company save: LibreOfficeDev headless observation

This experiment creates a PPTKit presentation with company `Old Company`, then compares an unchanged control with an `extended-company-v4` source save explicitly replacing it with `New & 中文`. An independent ZIP comparison found identical entry sets and only `docProps/app.xml` changed; its bytes differ solely in the `Company` text.

LibreOfficeDev **26.8.0.0.alpha0**, on macOS 26.5 arm64, headlessly resaved each branch three times. The control retained `Old Company` and the edited branch retained `New & 中文` in all three outputs. The observed title and revision also remained unchanged. In both branches the application replaced `Application` with its own identifier and removed the extended `Slides` property. This experiment establishes neither complete extended-property fidelity nor interactive editing behavior.

The [record](extended-company-libreoffice-evidence.json) contains hashes, exact inventories and process logs. Check retained files with `python3 scripts/research/verify-extended-company-libreoffice.py` and reproduce in the recorded environment with `python3 scripts/research/extended-company-libreoffice.py`. Artifacts under `.tmp` are local, not a portable fixture corpus. PowerPoint, repair prompts, interactive editing and other company values remain untested.
