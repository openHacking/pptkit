# Wrapped mixed SVG text observation

A 75-point-wide text box contains two runs, red `Alpha ` and blue bold `Beta Gamma`,
with wrapping enabled. MuPDF 1.29.0 through PyMuPDF 1.28.0 displayed both colors
with `mixedTextMode: "native"`, and extracted three lines: `Alpha`, `Beta`,
`Gamma`. The authored text is one line with spaces, so this extraction changes
whitespace. The default XHTML output produced no visible colored pixels or extracted
text in this engine.

The [manifest](native-wrapped-svg-evidence.json) records the text, pixel counts,
versions and output hashes. The [observer](../../scripts/research/oracles/native-wrapped-svg-probe.py)
checks both colors, multiple extracted lines and the non-whitespace character
sequence. This is one sample, not proof of font identity, shaping, line-break
equivalence, exact selection semantics or arbitrary script support.

```sh
rtk proxy node scripts/research/native-wrapped-svg-probe.mjs
rtk proxy python3 scripts/research/oracles/native-wrapped-svg-probe.py
```

The artifacts are retained in `.tmp/native-wrapped-svg-probe`.
Later font-dependency reporting adds structured details to warning reports, so
rerunning with current code changes the historical warning-report hashes.
