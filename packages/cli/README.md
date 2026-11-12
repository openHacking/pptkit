# `@pptkit/cli`

Local JSON generation, PPTX inspection, OPC comparison and create-only preservation using public engine APIs.

```sh
pptkit generate deck.json deck.pptx --json
pptkit generate deck.json deck.pptx --asset-base ./assets --json
pptkit validate deck.json --json
pptkit validate deck.json --export --json
pptkit generate deck.json deck.pptx --fail-on-export-warning --json
pptkit preview deck.json --slide 1 --json
pptkit preview deck.json --slide 1 --output slide.svg --json
pptkit preview deck.json --slide all --output deck.html --json
pptkit preview deck.pptx --input-format pptx --slide 1 --json
pptkit preview deck.pptx --input-format pptx --slide all --json
pptkit inspect deck.pptx --projection content --json
pptkit inspect deck.pptx --projection content --edit-save-profile text-runs-v3 --text-values --json
pptkit compare before.pptx after.pptx --json
pptkit compare before.pptx after.pptx --profile opc-xml-v1 --max-differences 100
pptkit save deck.pptx copy.pptx --mode no-op --json
pptkit save deck.pptx numbered.pptx --mode edit --first-slide-number 17 --json
pptkit save deck.pptx edited.pptx --mode edit --text-edits edits.json --json
pptkit save deck.pptx chart-edited.pptx --mode edit --chart-edits charts.json --json
pptkit save deck.pptx picture-edited.pptx --mode edit --picture-edit picture.json --picture-bytes image.gif --json
pptkit save deck.pptx edited.pptx --mode edit --text-edits edits.json --dry-run --json
```

Use `pptkit --help` for flags and exit codes. JSON reports have a versioned envelope, relative input paths, captured SHA-256 and explicit native-verification status. Malformed inputs, I/O, cancellation and resource limits return bounded errors; comparison differences use exit 1 and stdout. Existing source files remain unchanged. See the [contract](../../docs/api/cli-inspection.md) for profiles, bounds and limitations.

Programmatic file commands use `runCliAsync`; synchronous `runCli` remains available for informational help. Save requires an explicit mode and creates a new destination without replacing existing files; edit mode supports starting-number, source-bound text, existing chart values and a separate single-picture replacement profile. See the [save contract](../../docs/api/cli-save.md).

`inspect --capabilities` selects the versioned seven-operation view, with exact feature/part filters and complete-subject truncation. See [capability observations](../../docs/api/import-capabilities.md).

The [generation contract](../../docs/api/cli-generation.md) describes the versioned JSON input, embedded assets and warnings.
