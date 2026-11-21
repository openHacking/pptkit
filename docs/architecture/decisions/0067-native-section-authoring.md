# 0067: Explicit native section authoring

## Decision

Core owns an ordered collection of sections with independent string IDs, optional names, and explicit slide IDs. Names may be absent, empty, or repeated; empty sections are retained. `setSections` atomically replaces the collection, so creation, renaming, merging, deletion, and moving membership use the same checked operation. Removing a section does not remove slides. IDs identify sections only within the document and are independent of source PPTX GUIDs.

The existing per-slide `section` string remains an independent legacy label. It does not create or change native section membership. Inserting or copying slides from another document leaves them unassigned; duplicating a slide within a document retains its section. Removing a slide removes its membership, and reordering slides retains membership. A section's members are returned in presentation order.

Core permits unassigned or interleaved membership while editing. Fresh PPTX export requires nonempty section collections to cover every slide exactly once, in section order, with contiguous members; otherwise it fails. Empty collections produce no native section extension. Export derives stable section GUIDs from Core IDs, checks collisions, and maps slide identities to the numeric IDs in the generated presentation. Generated GUIDs are not source GUID preservation.

## Boundaries

Native inspection still reports source GUIDs and unknown extension data separately. Existing import projections only project member names as legacy slide labels; they do not reconstruct Core sections. A no-op source save preserves the original bytes, but changing the new collection is not a supported source-preserving edit. Native application interoperability and full source-to-Core section projection remain separate work.
