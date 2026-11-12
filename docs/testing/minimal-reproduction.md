# Preservation-first reproduction example

This example establishes a reproducible baseline for ECO-09.1.3; it does not extract slides, reduce a package, or sanitize it. Run it in a fresh directory with a built checkout and an installed `pptkit` executable, as in the [CLI workflow](debugging-workflow.md). The content is synthetic. Set `PPTKIT_REPO` to the absolute checkout path. Each command must succeed before proceeding; use fresh filenames because redirection can truncate existing reports and publication is create-only.

```sh
cat > deck.json <<'JSON'
{"version":"pptkit-authoring-v1","presentation":{"size":{"width":960,"height":540}},"slides":[{"id":"page","elements":[{"type":"text","id":"title","box":{"x":40,"y":40,"width":400,"height":80},"content":"Diagnostic example"}]}]}
JSON
pptkit generate deck.json source.pptx --json > generation.json
pptkit inspect source.pptx --projection content --capabilities --json > capabilities.json
node "$PPTKIT_REPO/scripts/inspect-preservation-closure.mjs" source.pptx /ppt/slides/slide1.xml > closure.json
pptkit save source.pptx candidate.pptx --mode no-op --json > save.json
pptkit compare source.pptx candidate.pptx --json > comparison.json
node "$PPTKIT_REPO/scripts/inspect-preservation-closure.mjs" candidate.pptx /ppt/slides/slide1.xml > candidate-closure.json
node --version > node-version.txt
```

The inspection, closure and save reports carry the captured source hash. Confirm that these hashes agree and that the comparison reports `equal: true` without truncation; compare the two closure inspections and hashes as well. `opc-payload-v1` compares package payloads, not ZIP container bytes, so also compare whole-file SHA-256 hashes when exact bytes matter. The no-op save is expected to preserve those bytes. Record the PPTKit revision and package versions, OS and relevant application version separately, together with the exact command or operation, expected versus observed result, reproduction outcome, and task IDs such as ECO-09.1.3 and IMP-07.1. A successful run of this synthetic example is not a report of a failure in an external presentation.

For an actual issue, preserve the source and repeat the original operation on the no-op candidate. Accept it as a reproduction only if the specified failure recurs. Inspection may fail or exceed a bound; retain that result, stop the affected candidate and keep the original. Source-relationship completeness does not cover arbitrary implicit XML/binary references, and the closure's selected parts cannot simply be copied into a valid or equivalent package. A future reduced candidate needs its own dependency analysis, part-difference review and failure-predicate check; see [ADR 0064](../architecture/decisions/0064-reproduction-and-redaction.md).

Reports and manifests can themselves contain document content, paths, external targets and original values. Removing visible text or metadata does not sanitize all package content. Share only artifacts that their owner has independently reviewed, or reconstruct the issue with synthetic data.

The [CLI workflow test](../../packages/cli/test/commands.test.mjs) exercises the generated fixture, inspection, no-op copy, exact hashes, closure equivalence, failed-root handling and source preservation. The [closure integration tests](../../scripts/preservation-closure-parity.test.mjs) separately cover shared resources, inactive relationships and unreachable content. Neither establishes native behavior or automated minimization.
