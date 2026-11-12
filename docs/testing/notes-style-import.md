# Notes-master text source evidence

The [runner](../../scripts/inspect-notes-style-import.mjs) creates two notes parts with different actual notes masters/themes, and deliberately different local/master placeholder indices (`1` and `99`). Only type-body matching selects placeholder defaults. Four stages retain the original and three mixed notes/drawing scalar saves, using explicit notes and supplemental-font profiles.

**Conformance clarification (2026-10-26):** this particular fixture contains a second unlisted actual notes master and is deliberately non-conforming. [Microsoft's specification](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document#notes-master-part) permits at most one Notes Master part per package. The manifest now records that classification and the actual `unlisted-master` finding. These checks establish diagnosed source transport and byte preservation, not valid multi-notes-master authoring or native compatibility. The separate [single-master override/isolation evidence](notes-master-isolation.md) covers the cardinality-correct alternative and rejects two listed IDs.

The [manifest](notes-style-import-evidence.json) records artifact hashes and local directory. A separate [Python stdlib OPC/XML oracle](../../scripts/research/oracles/notes-style-import.py) follows actual notesMaster/theme relationships from ZIP payloads. It verifies **22** selected note style properties per stage: generic/supplemental fonts, language, inherited placeholder font size/italic and mapped color/alpha. It checks actual type attributes and unequal indices, source/target/dependency existence and both scalar note strings. Parent prompt strings never become Core notes.

After each edit, only the two notes parts and slide1 change. A named-fixture lexical mask admits scalar `a:t` character contents; all other bytes in those XML payloads, all master/theme/relationship/other payloads and the part set remain identical. This mask is an evidence check, not a general XML writer. Deliberately changing the actual notes Hant family or the master placeholder's type is rejected by both actual source-style checks and byte-preservation checks in every stage.

Run after building:

```sh
node scripts/inspect-notes-style-import.mjs
```

[Browser tests](../../scripts/notes-style-import-parity.test.mjs) separately cover both dialects, Node/ESM/global full report/document/bytes parity, actual XPath and color-map checks, fresh notes font slots and explicit map-rule loss, profile capture and failure/retry. Public packed runtime/types exercise the Node file option and one real scalar prepare/save.

No new native application/GUI typography or notes-page rendering check was run. The manifest explicitly retains native appearance/editing as `untested`, full notes-page layout as `unsupported`, and notes style writes as `blocked`. Named source evidence does not close complete notes/font/native acceptance gates.
