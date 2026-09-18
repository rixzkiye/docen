# Preserve-only elements and complex structures

The Tiptap layer edits WordprocessingML content; several OOXML branches have no
editable node by design. This document is the written contract for those
branches — what is carried, what is rendered, what is warned about, and what is
explicitly out of scope. The machine-checkable half lives in
[`src/extensions/coverage.ts`](../src/extensions/coverage.ts) (`Disposition`
tables + `PRESERVE_ONLY_ELEMENTS`) and is proven by
[`coverage.spec.ts`](../src/extensions/coverage.spec.ts): every audited tag has
a note, every note agrees with the disposition tables, and each passthrough
branch survives `resolve → compile` and — where office-open can stringify a
synthetic options object — a real `generateDocument → parseDocument` byte trip.

## Preserve-only elements

| Element                                     | Where          | Carried as                                       | Canvas / editor behavior                                                                                                                             | Authoring                                                          |
| ------------------------------------------- | -------------- | ------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------ |
| `smartArt`                                  | inline         | `inlinePassthrough` atom (part bytes + rels)     | inert atom; no diagram replay/preview yet                                                                                                            | **excluded** (SmartArt authoring is an explicit program exclusion) |
| `object` (OLE)                              | inline         | `inlinePassthrough` atom (embed bytes + preview) | the OLE vector preview paints as a framed picture; no editing inside                                                                                 | **excluded** (OLE payload editing)                                 |
| `symbolRun`                                 | inline         | `inlinePassthrough` atom (`w:char`/`w:font`)     | round-trip faithful; glyphs are not painted (font-relative PUA mapping is out of model scope)                                                        | no                                                                 |
| `commentRangeStart/End`, `commentReference` | inline         | `inlinePassthrough` atom                         | zero-width markers; the comment range is tinted and its balloon anchored. The comment body IS editable through the comments pane                     | comment body: yes; markers: no                                     |
| `customXml`                                 | inline + block | inline atom / block `passthrough` atom           | block-level customXml paints a labeled placeholder box; inline wrappers are zero-width                                                               | no                                                                 |
| `subDoc`                                    | inline         | `inlinePassthrough` atom (referenced part bytes) | atom is inert; the editor shows the unsupported-content warning for the document                                                                     | no                                                                 |
| `proofErr`                                  | inline         | `inlinePassthrough` atom                         | zero-width proofing metadata (the editor recomputes its own spelling state); invisible                                                               | no                                                                 |
| `rawXml`                                    | inline + block | inline atom / block `passthrough` atom           | block-level rawXml paints a labeled placeholder box; inline atoms are zero-width unless they are a 3D/ink drawing (those parse into `model3d`/`ink`) | no                                                                 |
| `altChunk`                                  | block          | `passthrough` atom (part bytes)                  | labeled placeholder box; the editor shows the unsupported-content warning                                                                            | no                                                                 |

"Byte-faithful" means the branch's complete office-open options object rides in
the atom's JSON (`encodePassthroughData` base64s binary leaves such as OLE/VML
bytes) and compiles back deep-equal — the coverage spec asserts exactly that for
every listed element, plus real-XML trips for `symbolRun`, `customXml`,
`rawXml`, `proofErr`, comment anchors and `altChunk`.

## Complex structures (item 14)

| Structure              | Parse                                                                     | Edit                                                                               | Round-trip                                                                                                                                                          | Editor treatment                               |
| ---------------------- | ------------------------------------------------------------------------- | ---------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------- |
| `altChunk`             | preserved (block passthrough)                                             | not editable                                                                       | byte-faithful (part names normalized deterministically)                                                                                                             | placeholder box + warning bar                  |
| `subDoc`               | preserved (inline passthrough; office-open loads the part)                | not editable                                                                       | byte-faithful                                                                                                                                                       | warning bar                                    |
| glossary (Quick Parts) | parsed into `documentExtras.glossary` (office-open)                       | Quick Parts dialog inserts parts; the glossary table itself is not directly edited | byte-faithful (`glossary.spec.ts`)                                                                                                                                  | Quick Parts dialog / Building Blocks organizer |
| encrypted container    | **rejected with `EncryptedDocumentError`** (`code: "ENCRYPTED_DOCUMENT"`) | n/a                                                                                | office-open can pass the container through at the `DocumentOptions` level; docen's `parseDOCX` refuses it so the editor never shows an empty page for a locked file | explicit open error, no editing                |

SmartArt authoring, OLE/ink/3D authoring and macro execution remain excluded
from R8 (see the program's §5); the preserve-only contract above is their
required close.

## Round-trip guarantees that are NOT preserve-only

Comments (`w:comment` bodies), tracked changes, bookmarks, fields, math, SDTs,
text boxes and shapes all have editable nodes or marks; they are documented in
the disposition tables as `editable` and covered by their own specs. The
`commentRange*` markers above are the only comment surface that stays opaque.
