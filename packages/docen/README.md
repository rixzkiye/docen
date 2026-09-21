# docen

![npm version](https://img.shields.io/npm/v/docen)
![npm downloads](https://img.shields.io/npm/dw/docen)
![npm license](https://img.shields.io/npm/l/docen)

> The full `<docen-document>` canvas web-component editor plus headless Markdown/DOCX conversion — one package (editor available via the `docen/editor` subpath).

## Features

- 🔄 **Markdown ⇄ DOCX Conversion** - Seamless bidirectional conversion (styled HTML is accepted as paste input)
- 🎯 **Unified API** - Consistent, intuitive interface across all format conversions
- 📦 **All-in-One Package** - Single dependency for both headless conversion AND the full `<docen-document>` editor (via `docen/editor`)
- 🔧 **Built on TipTap** - Powered by the robust TipTap/ProseMirror ecosystem
- 💪 **TypeScript-First** - Full type safety with comprehensive TypeScript support
- ⚡ **Zero Configuration** - Works out of the box with smart defaults
- 🌳 **Extensible** - Customize conversions with TipTap extensions and advanced options
- 🔄 **Bidirectional** - Convert in both directions (parse ↔ generate) for each format

## Installation

```bash
# Install with npm
$ npm install docen

# Install with yarn
$ yarn add docen

# Install with pnpm
$ pnpm add docen
```

## Quick Start

### Markdown ↔ TipTap JSON

```typescript
import { parseMarkdown, generateMarkdown } from "docen";

// Parse Markdown to TipTap JSON
const doc = parseMarkdown("# Hello World\n\nThis is **bold** text.");

// Generate Markdown from TipTap JSON
const markdown = generateMarkdown(doc);
```

### DOCX ↔ TipTap JSON

```typescript
import { parseDOCX, generateDOCX } from "docen";

// Parse DOCX to TipTap JSON
const doc = parseDOCX(buffer);

// Generate DOCX from TipTap JSON
const docxBuffer = await generateDOCX(doc); // defaults to a Node.js Buffer
```

### PDF Generation

Generate vector PDFs server-side in pure Node.js without browser or DOM:

```typescript
import { generatePDF } from "docen";

// Generate PDF from TipTap JSON
const pdfBytes = await generatePDF(doc, {
  title: "My Document",
  pdfa: "2b", // ISO 19005-2 PDF/A-2b compliance
  pdfUa: true, // ISO 14289-1 PDF/UA-1 accessibility
  backend: "node", // "node" (default) or "chromium"
});
```

### Cross-Format Conversion

Convert between formats by using TipTap JSON as the intermediate format:

```typescript
import { parseMarkdown, generateDOCX, generatePDF } from "docen";

// Markdown → DOCX
const md = "# Title\n\nContent...";
const doc = parseMarkdown(md);
const docx = await generateDOCX(doc, { packer: { type: "blob" } });

// Markdown → PDF
const pdf = await generatePDF(doc, { title: "Title" });
```

### Headless PDF Engine (via `docen/pdf`)

Import from `docen/pdf` to access low-level PDF primitives, scene serialization, and vector PDF compilation:

```typescript
import { renderPdf, pagesToPdf, serializeNodeScene } from "docen/pdf";
```

### Full Editor (via `docen/editor`)

The same `docen` package also re-exports the turnkey web-component editor. Import from the `docen/editor` subpath to register `<docen-document>` and apply a theme:

```html
<docen-document id="doc" filename="Welcome.docx"></docen-document>

<script type="module">
  import { registerComponents, applyTheme } from "docen/editor";
  registerComponents();
  applyTheme("light");
</script>
```

> The editor lives on a subpath so that pure-converter imports (`import { parseDOCX } from "docen"`) stay tree-shakable and never pull in the Fluent UI shell.

## API Reference

### Markdown Functions

#### `parseMarkdown(markdown)`

Parses a Markdown string into TipTap JSON content.

**Parameters:**

- `markdown: string` - Markdown string to parse

**Returns:** `JSONContent` - TipTap document object

```typescript
const doc = parseMarkdown("# Hello\n\nWorld");
```

#### `generateMarkdown(doc)`

Generates a Markdown string from TipTap JSON content.

**Parameters:**

- `doc: JSONContent` - TipTap document object

**Returns:** `string` - Markdown string

```typescript
const markdown = generateMarkdown({ type: 'doc', content: [...] });
```

### DOCX Functions

#### `parseDOCX(input)`

Parses a DOCX file into TipTap JSON content. The archive is validated against
untrusted-input admission limits (entry/size/ratio/media caps, bounded
inflation, XML depth/text budgets) before parsing; violations throw an
`ArchiveRejection` whose `code` names the failed rule.

**Parameters:**

- `input: Buffer | ArrayBuffer | Uint8Array | Blob | ReadableStream` - DOCX file data

**Returns:** `Promise<JSONContent>` - TipTap document object

```typescript
import { readFileSync } from "node:fs";
const buffer = readFileSync("document.docx");
const doc = await parseDOCX(buffer);
```

#### `generateDOCX(docJson, options?)`

Generates a DOCX file from TipTap JSON asynchronously. Styling is derived from the TipTap attrs. Byte-reproducible: the same input yields the same output in the same process and in a fresh one, with a fixed generation clock (`DOCX_EPOCH` by default; `options.date` overrides, `null` omits generated dates). The input JSON is never mutated.

**Parameters:**

- `docJson: JSONContent` - TipTap document object
- `options?: DocxGenerateOptions` - `{ prepare?, packer?, document?, variant?, date? }`:
  - `prepare` (default `true`): `true` runs the default **local-only** preparation (no network); `false` skips it; a `PrepareStep[]` runs custom steps. External `http(s)` images require an explicit `prepareImages({ allow: [...] })` policy — scheme/host allowlist, response size cap, redirect cap and timeout are enforced.
  - `packer`: `PackerOptions`; `type` controls the output format (`"nodebuffer"` default → Buffer, `"blob"`, `"arraybuffer"`, …).
  - `date`: fixed clock for generated dates (string/`Date`; default `DOCX_EPOCH`; `null` omits).

**Returns:** `Promise<Buffer | Blob | ArrayBuffer | Uint8Array | string>` - DOCX data in the requested format

```typescript
import { prepareImages } from "docen/docx";

// Default: local-only preparation, Node.js Buffer
const buffer = await generateDOCX(doc);

// Remote images: explicit host allowlist (never fetched by default)
const withImages = await generateDOCX(doc, {
  prepare: [prepareImages({ allow: ["cdn.example.com"] })],
});

// Skip preparation, Browser Blob
const blob = await generateDOCX(doc, { prepare: false, packer: { type: "blob" } });
```

#### `generateDOCXSync(docJson, options?)`

Synchronous variant — fastest throughput, blocks the event loop. Does **not**
run `prepareDocument` (it is async); call
`const prepared = await prepareDocument(doc)` first when images need
embedding. Byte-reproducible like `generateDOCX`, and never mutates its input.

**Returns:** `Buffer | Blob | ArrayBuffer | Uint8Array | string` - DOCX data in the requested format

```typescript
const buffer = generateDOCXSync(doc);
```

#### `generateDOCXStream(docJson, options?)`

Streams the DOCX as a `ReadableStream<Uint8Array>` — for large documents or
HTTP responses. Runs the default **local-only** `prepareDocument` (async) and
is byte-identical to the sync path for the same input. Time/peak-RSS numbers
for 100–300 page documents are recorded in
[`@docen/docx/bench/README.md`](../docx/bench/README.md).

**Returns:** `Promise<ReadableStream<Uint8Array>>`

```typescript
const stream = await generateDOCXStream(doc);
return new Response(stream);
```

## Advanced Usage

### Custom Extensions

Use custom TipTap extensions for Markdown conversions:

```typescript
import { CustomExtension } from "./custom-extension";
import { parseMarkdown, generateMarkdown } from "docen";

const doc = parseMarkdown(md, [CustomExtension]);
const mdContent = generateMarkdown(doc, [CustomExtension]);
```

### DOCX Template Patching

Replace `{{placeholders}}` in a DOCX template with TipTap-JSON content:

```typescript
import { patchDOCX, parseMarkdown } from "docen";

const result = await patchDOCX({
  template: templateBuffer,
  patches: {
    title: { content: parseMarkdown("# Report") },
    body: { content: parseMarkdown("## Section\n\nHello **world**.") },
  },
  outputType: "nodebuffer",
});
```

Each patch's `content` is compiled to DOCX (styling derived from attrs) and the first section's children replace the placeholder. `keepOriginalStyles`, `recursive`, and `placeholderDelimiters` mirror the underlying `@office-open/docx` `patchDocument`.

## Format Conversion Matrix

| From \ To    | Markdown | DOCX     |
| ------------ | -------- | -------- |
| **Markdown** | -        | via JSON |
| **DOCX**     | via JSON | -        |

All conversions go through TipTap JSON as the intermediate format, ensuring consistency and enabling cross-format transformations.

## Supported Content Types

### Text Formatting

- Bold, Italic, Underline, Strikethrough
- Superscript, Subscript
- Text highlights, colors, backgrounds
- Font families, sizes, line heights

### Block Elements

- Headings (H1-H6)
- Paragraphs with alignment
- Blockquotes
- Horizontal rules
- Code blocks with syntax highlighting

### Lists & Tables

- Bullet lists, ordered lists
- Task lists with checkboxes
- Tables with colspan/rowspan

### Media & Links

- Images with embedded base64
- Hyperlinks

## Use Cases

- **Content Management Systems** - Import/export documents in multiple formats
- **Documentation Tools** - Convert between Markdown and Word
- **Note-taking Apps** - Support various import/export formats
- **Report Generation** - Generate DOCX reports from Markdown templates
- **Content Migration** - Migrate content between different formats
- **Collaborative Editing** - Use TipTap editor with format support

## Under the Hood

`docen` ships three entry points: the **root** re-exports the high-level converters; **`docen/docx`** exposes the full engine (`createDocxEditor`, `docxExtensions`, resolve/compile/prepare, styles); **`docen/editor`** exposes the `<docen-document>` web component. It builds on:

- **@docen/docx** - DOCX / Markdown converters built on the DocxManager architecture (full surface via `docen/docx`)
- **@docen/editor** - Fluent UI shell + docx engine → `<docen-document>` (exposed via the `docen/editor` subpath)
- **@office-open/docx** - Native OOXML parse/generate (`parseDocument`, `generateDocument`, `patchDocument`)
- **@docen/markdown** - Format-agnostic Markdown syntax layer — the IR + renderer behind `parseMarkdown`/`generateMarkdown` (bundled into `@docen/docx`, no separate install)

## Comparison with Alternatives

| Feature     | docen | markdown-docx | mammoth | turndown |
| ----------- | ----- | ------------- | ------- | -------- |
| MD → DOCX   | ✅    | ✅            | ❌      | ❌       |
| DOCX → MD   | ✅    | ❌            | ❌      | ❌       |
| TypeScript  | ✅    | ✅            | ✅      | ✅       |
| Unified API | ✅    | ❌            | ❌      | ❌       |
| Extensible  | ✅    | ❌            | ❌      | ✅       |

## License

- [MIT](../../LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
