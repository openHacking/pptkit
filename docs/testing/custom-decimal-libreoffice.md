# Custom decimal values: LibreOfficeDev headless observation

A synthetic source replaces the custom-property part of a hash-checked PPTKit package with four `vt:decimal` values: `1.25`, `0.125`, `0.1` and `9007199254740992`. The edited branch changes the three exactly projected values to `2.5`, `0.5` and `4503599627370496`; `0.1` stays source-only and unchanged. The control is not edited. Independent byte comparison confirms that only the three selected text values in the custom-property part changed.

LibreOfficeDev 26.8.0.0.alpha0 headlessly resaved each branch three times. On the first resave all four decimal properties disappeared in both branches; the application instead added `PresentationFormat` and `Slides`. Their absence remained through round three. This is an observation of one synthetic source and environment, not proof about every Office file or decimal value. No interactive editing or repair prompt was tested.

The [record](custom-decimal-libreoffice-evidence.json) includes hashes, inventories and process logs. Run `python3 scripts/research/verify-custom-decimal-libreoffice.py` to check retained artifacts, or `python3 scripts/research/custom-decimal-libreoffice.py` to reproduce in the recorded environment. Artifacts under `.tmp` are local.
