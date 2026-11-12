# Explicit revision save: LibreOfficeDev headless observation

A SHA-256-verified PPTKit-produced input from `core-metadata-libreoffice-evidence.json` has revision 7. This experiment compares an unchanged control with a `core-values-v2` source save explicitly changing revision to 42. An independent ZIP comparison found identical entry sets and only `docProps/core.xml` changed; its bytes differ solely in the revision text.

LibreOfficeDev **26.8.0.0.alpha0**, on macOS 26.5 arm64, headlessly resaved each branch three times. In this sample the control retained revision 7 and the edited branch retained 42 in all three outputs; the observed title and creator also remained unchanged. The application rewrote the keywords from `alpha, beta` to `alpha beta` in both branches. This observation does not establish general metadata fidelity.

The [record](core-revision-libreoffice-evidence.json) contains file hashes, exact inventories and process logs. Check retained files with `python3 scripts/research/verify-core-revision-libreoffice.py` and reproduce in the recorded environment with `python3 scripts/research/core-revision-libreoffice.py`. Artifacts under `.tmp` are local, not a portable fixture corpus. PowerPoint, interactive property editing, repair prompts and other revision values remain untested.
