# Explicit keyword save: LibreOfficeDev headless observation

A SHA-256-verified PPTKit-produced input from `core-metadata-libreoffice-evidence.json` contains the aggregate keyword string `alpha, beta`. This experiment compares an unchanged control with a `core-keywords-v3` source save explicitly replacing it with `New, value & 中文`. An independent ZIP comparison found identical entry sets and only `docProps/core.xml` changed; its bytes differ solely in the keyword text.

LibreOfficeDev **26.8.0.0.alpha0**, on macOS 26.5 arm64, headlessly resaved each branch three times. On the first save the control became `alpha beta` and the edited branch became `New value & 中文`; both then remained unchanged through the following two resaves. In both branches the observed title, creator and revision remained unchanged. The comma removal means this sample does **not** establish exact keyword retention by LibreOffice.

The [record](core-keywords-libreoffice-evidence.json) contains hashes, exact inventories and process logs. Check retained files with `python3 scripts/research/verify-core-keywords-libreoffice.py` and reproduce in the recorded environment with `python3 scripts/research/core-keywords-libreoffice.py`. Artifacts under `.tmp` are local, not a portable fixture corpus. PowerPoint, interactive property editing, repair prompts and other keyword values remain untested.
