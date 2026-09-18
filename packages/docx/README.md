# @docen/docx

![npm version](https://img.shields.io/npm/v/@docen/docx)
![npm downloads](https://img.shields.io/npm/dw/@docen/docx)
![npm license](https://img.shields.io/npm/l/@docen/docx)

> DOCX editor and converter powered by @office-open/docx with Tiptap editing layer, supporting bidirectional conversion between DOCX and Markdown (styled HTML is accepted as paste input).

> Need a ready-made visual editor? [`@docen/editor`](../editor/README.md) wraps this engine in a Fluent UI host with the turnkey `<docen-document>` web component.

## Features

- 📝 **Viewless Tiptap engine** — the DOCX editing model with DOCX-aware extensions; rendering belongs to the host (e.g. @docen/editor's canvas stage)
- 🔄 **DOCX Round-trip** — Near-lossless DOCX ↔ Editor conversion via @office-open/docx
- 📄 **Markdown Support** — Tiptap JSON ↔ Markdown conversion
- 🎨 **DOCX Properties** — Custom Tiptap extensions carry shading, borders, indent, spacing, floating, crop
- 🔗 **Template Patching** — Replace `{{placeholders}}` in DOCX templates with Tiptap-JSON content

## Installation

```bash
# Install with pnpm
$ pnpm add @docen/docx

# Install with npm
$ npm install @docen/docx
```

### Runtime dependencies

- `@docen/markdown` (the Markdown IR layer behind `parseMarkdown`/`generateMarkdown`) is **bundled into `dist`**: the published package is self-contained and never resolves a separate markdown package at install time.
- `linkedom` is a real **runtime dependency** — `parseHTML` lazily imports it on the server, where no native DOM exists. Browser bundles take the native `DOMParser` path and never load it.

## Quick Start

```typescript
import { docxExtensions, parseDOCX, generateDOCX } from "@docen/docx";
import { Editor } from "@docen/docx/core";

// Viewless: the editor is the editing model — rendering belongs to the host.
const editor = new Editor({
  element: null,
  extensions: docxExtensions,
  content: await parseDOCX(buffer),
});

// Save back to DOCX
const output = await generateDOCX(editor.getJSON());
```

## API

### Standalone Functions (Core)

These work without an editor instance — for headless/server/batch use.

```typescript
import {
  parseDOCX,
  generateDOCX,
  generateDOCXSync,
  generateDOCXStream,
  prepareImages,
  prepareImageSizes,
  parseMarkdown,
  generateMarkdown,
} from "@docen/docx";

// DOCX pipeline: DOCX binary ↔ Tiptap JSON
const json = parseDOCX(buffer); // → JSONContent (untrusted-input limits enforced)
const buffer = await generateDOCX(json); // → Buffer (local-only preparation; no network)
const blob = await generateDOCX(json, { packer: { type: "blob" } }); // → Blob
const sync = generateDOCXSync(json); // → Buffer (skips prepare)
const stream = await generateDOCXStream(json); // → ReadableStream<Uint8Array>

// External (http/https) images: opt in explicitly with a host allowlist —
// scheme + host are validated, the response is size-capped, redirects are
// counted and re-validated, and each request has a timeout.
const withImages = await generateDOCX(json, {
  prepare: [prepareImages({ allow: ["cdn.example.com"] }), prepareImageSizes()],
});

// Byte-reproducibility: every generated date/id sequence is fixed per
// generation. `date` overrides the fixed clock (default DOCX_EPOCH,
// "1980-01-01T00:00:00.000Z"; `null` omits generated dates).
const reproducible = await generateDOCX(json, { date: "2024-01-02T03:04:05.000Z" });

// Package variants: `variant` stamps the main document part's content type —
// "docx" (standard), "docm", "dotx", or "dotm" (ECMA-376 / MS-OFFMACRO main
// types). Office recognizes a template/macro package by that declaration, so
// saving a `.dotx`/`.docm` needs the matching variant. Requesting "docx" flips
// a macro-enabled/template source back to the standard document main type;
// omitting the variant keeps the source's own declaration verbatim
// (source-faithful round-trip). Macro and unknown parts carried in
// `documentExtras.rawParts` stay in every variant, byte-identical, and a
// source-less document keeps its model-derived parts (numbering, footnotes,
// endnotes) declared and emitted under every variant.
const template = await generateDOCX(json, { variant: "dotx" }); // → Buffer (Word Template)
const docm = await generateDOCX(json, { variant: "docm" }); // → Buffer (macro-enabled document)

// Markdown pipeline: Markdown string ↔ Tiptap JSON
const json = parseMarkdown("# Hello"); // → JSONContent
const md = generateMarkdown(json); // → string
```

### Editor

```typescript
import { createDocxEditor, docxExtensions } from "@docen/docx";

const editor = createDocxEditor({
  element: document.querySelector("#editor"),
  extensions: docxExtensions,
  spellcheck: true,
  editable: true,
});
```

Need a ready-made editor UI? [`@docen/editor`](../editor/README.md)'s `<docen-document>` bundles this engine with open/save, canvas rendering, and the Fluent host.

### Template Patching

Replace `{{placeholders}}` in a DOCX template with Tiptap-JSON content. Each
patch's `content` is prepared on a copy (default: local-only preparation; no
network) then compiled to DOCX.

```typescript
import { patchDOCX, parseMarkdown } from "@docen/docx";

const result = await patchDOCX({
  template: templateBuffer,
  patches: {
    title: { content: parseMarkdown("# Report") },
  },
  outputType: "nodebuffer",
});
```

### Advanced: Model Bridge

`generateDOCX` runs `prepareDocument → compileDocument → generateDocument`
internally. You rarely need these directly — reach for them only when working
with the intermediate `DocumentOptions` (the OOXML persistence model):

```typescript
import { resolveDocument, compileDocument, prepareDocument } from "@docen/docx";

const json = resolveDocument(docOpts); // DocumentOptions → JSONContent
const docOpts = compileDocument(json); // JSONContent → DocumentOptions
const prepared = await prepareDocument(json); // copy; local-only, `json` untouched
const withRemoteImages = await prepareDocument(json, [
  prepareImages({ allow: ["cdn.example.com"] }), // explicit network opt-in
  prepareImageSizes(),
]);
```

### Generated fields and TOC (Word-faithful cached results)

DOCX stores every field twice: the instruction and the cached result Word
paints until the next update. `generateDOCX` / `generateDOCXSync` /
`generateDOCXStream` re-derive the structural caches at generation time (the
pass is non-mutating — the input JSON is untouched):

- **SEQ** — per-label ordinals in document order, `\s <level>` heading
  restarts, `\*` number formats (Arabic/Roman/Alphabetic) and the
  `w:caption@w:sep` chapter separator.
- **REF / NOTEREF** — the target bookmark's inner text (forward references
  included). `\p`/`\n` keep the model's cache (they need reading-order terms
  and paragraph numbering the builder does not own).
- **TOC** — a `tocField` whose cached entries are missing (the from-scratch
  placeholder) gets real entry paragraphs built from the document's headings
  (`\o` window, `\t` custom styles, `\u` outline semantics, `\b` bookmark
  scope) or, for a `\c` field, from the matching captions. The parsed DOCX
  option names are honored (`stylesWithLevels`, `entriesFromBookmark`,
  `captionLabelIncludingNumbers`), as are the editor dialog's older
  `styles`/`customStyles`/`bookmark`/`captionLabel` keys. The field
  instruction stays intact, so Word/LibreOffice/the editor's Update Table still
  refresh it; a TOC that already carries entries is never recomputed. A TOC
  that cannot be filled keeps its field: the package carries the dirty
  placeholder field (`w:fldChar` + `w:dirty="1"`) Word updates on open instead
  of silently dropping the instruction.
- **PAGE / NUMPAGES / PAGEREF / SECTION / SECTIONPAGES** — resolved through
  the optional `fields` option: pass `pageOf` (and `pageCount`) from your
  pagination. The editor's save path feeds its live canvas pages; without a
  context these fields keep the model's cache rather than inventing a number:
  Word updates them on open from its own pagination.
- **TOC entry page numbers** — pass `tocPageOf` (1-based displayed page per
  collected heading/caption index) and a filled TOC carries the same cached
  numbers a reader shows without updating the field. The editor's save path
  derives them from the live canvas pagination; without the callback the
  entries keep Word's empty page-number slot. Each entry's right-aligned page
  number tab stops at its own section's text-column edge (page size minus
  margins, per-section sectPr honored; the docen A4/1" defaults when the model
  carries none), never at a fixed Word Letter-default position that overruns an
  A4/1" column and makes LibreOffice drop the number.

```typescript
// Page context for a headless save, e.g. after your own layout pass:
await generateDOCX(json, {
  fields: {
    pageCount: 12,
    pageOf: ({ index, bookmark }) => pageByFieldIndex.get(index), // 1-based
    tocPageOf: ({ index, kind }) =>
      kind === "heading" ? headingPage.get(index) : captionPage.get(index),
  },
});
```

### Field-update behavior (LibreOffice-verified subset)

The fixture matrix in [`tests/toc-fixtures/`](./tests/toc-fixtures/) drives
every generated document through LibreOffice headless (plain PDF, resave,
explicit "update all fields/indexes" over UNO) with python-docx/lxml as the
structural oracle (`tests/toc-oracles*.spec.ts`, `tests/toc-oracle.py`).
Verified behavior:

- **No drift** — a `\o`/`\t` heading TOC and a `\c` caption table keep their
  entries and page numbers; REF/NOTEREF text, PAGE/NUMPAGES/PAGEREF values and
  plain (no `\s`) SEQ ordinals stay stable. SEQ `\*` Roman/Alphabetic formats
  are stable; LibreOffice renders the cached TOC entries and page numbers
  without any field update.
- **LibreOffice limitations** (not docen defects): an explicit update ignores
  `\b` bookmark scopes and the `\o`/`\u` filters (rebuilding 1-3 plus outline
  paragraphs), drops `\s` chapter restarts/prefixes, recomputes `\p`/`\n`, and
  merges multiple TOC fields into one content index on resave; a TOC inside a
  table cell fails its UNO load (the same file converts to PDF). Header/footer
  story field caches are preserved (not re-derived) at generation;
  LibreOffice computes PAGE/NUMPAGES/PAGEREF live but never SECTION, so an
  author-supplied SECTION cache is what readers show.
- **Word interop deviation** — Word writes a chapter-numbered caption as a
  separate `STYLEREF` field plus `SEQ`; docen caches the chapter prefix inside
  the SEQ result, so a Word field update drops the prefix until the caption is
  re-inserted. The cached value itself is correct and shown by Word/LO until
  an explicit update.

### Determinism

`generateDOCX`, `generateDOCXSync` and `generateDOCXStream` are byte-reproducible:
the same input produces the same output in the same process and in a fresh one.
The fixed clock (`DOCX_EPOCH`, overridable with `date`) covers generated
core-property/comment dates, generated id sequences restart per generation
(drawings, VML shapes, SmartArt UUIDs, font keys, altChunk part names), and the
ZIP entries are written with a fixed timestamp and stable order. Ids/dates the
source document already carries are preserved. See
[`docs/determinism.md`](./docs/determinism.md) for the seam contract and how to
verify it.

```typescript
await generateDOCX(json, { date: "2024-01-02T03:04:05.000Z" }); // explicit clock
await generateDOCX(json, { date: null }); // omit generated dates
```

### Untrusted input

`parseDOCX`/`parseDOCXSync` validate the ZIP package before office-open sees
it, against the Akademi importer's admission limits (32 MiB archive, 512
entries, 16 MiB/entry, 64 MiB aggregate, ratio 100, 64 XML depth, 100k XML
nodes, 200k attributes, 8 MiB attribute/text budgets, 1024 relationships)
plus owned-media caps (8 MiB/image, 32 MiB total, 256 entries). Every DEFLATE
payload is streamed through a capped inflater, so a size-lying zip bomb is
stopped at its declared cap instead of being materialized; violations throw an
`ArchiveRejection` (`error.code` names the failed rule). Streamed entries that
park their CRC/sizes in a data descriptor (general-purpose bit 3 — what
LibreOffice's DOCX export writes) are admitted after the descriptor is
validated against the central directory, so LibreOffice-saved files re-import.

### Generation defaults (page geometry & styles)

Generation never lets the engine's MS Office **zh-CN** defaults leak. Every
compiled section carries explicit `sectionProperties` from docen's own
constants — A4, 1" (1440-twip) margins on all sides, 0.5" header/footer
distance (`DOCEN_DEFAULT_PAGE_SIZE`, `DOCEN_DEFAULT_PAGE_MARGIN`,
`docenDefaultSectionProperties()`), instead of office-open's
`sectionMarginDefaults` (1.25" side margins, 851/992 header/footer). A
model-supplied sectPr is preserved verbatim; only a section without one gets
the defaults. The document style table keeps office-open's locale-neutral
ECMA-376 defaults (theme font refs — Calibri/Calibri Light — and `en-US`
language), and theme-only East Asian font tokens resolve with the
caller-supplied document language (`DOCEN_DEFAULT_EAST_ASIA_LANGUAGE` is
`en-US`; pass `"zh-CN"` explicitly for a zh-CN document) — never an implicit
zh-CN face (`等线`/DengXian).

### Streaming benchmark

`generateDOCXStream` on 100–300 page documents is measured by
[`bench/streaming-bench.mjs`](./bench/streaming-bench.mjs) (time + peak RSS,
sync/async/stream); recorded numbers and thresholds live in
[`bench/README.md`](./bench/README.md).

### Preserve-only content

Elements the Tiptap model deliberately carries without authoring — SmartArt,
OLE objects, symbol runs, comment range markers, customXml, subDoc, proofErr,
rawXml, altChunk — plus the complex-structure matrix (altChunk/subDoc/glossary/
encrypted containers) are documented in
[`docs/passthrough.md`](./docs/passthrough.md). Round-trip is byte-faithful;
the editor warns when a loaded document contains uneditable container content
instead of silently dropping it.

## Architecture

```
Standalone Functions (core)
  parseDOCX / generateDOCX / generateDOCXSync / generateDOCXStream / patchDOCX
  parseMarkdown / generateMarkdown
  resolveDocument / compileDocument / prepareDocument  (model bridge, advanced)
        ↕ used by
createDocxEditor / docxExtensions — the viewless Tiptap editing model
        ↕ rendered by
@docen/editor's canvas stage (LeaferJS pages, caret/selection mapping)
```

- **Runtime model**: Tiptap JSON with DOCX-rich attributes via custom extensions
- **Persistence model**: DocumentOptions (complete OOXML expressiveness)
- **Standalone functions are core** — the editor bindings are thin wrappers

## License

- [MIT](../../LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
