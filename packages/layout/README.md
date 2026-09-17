# @docen/layout

![npm version](https://img.shields.io/npm/v/@docen/layout)
![npm downloads](https://img.shields.io/npm/dw/@docen/layout)
![npm license](https://img.shields.io/npm/l/@docen/layout)

> Pure TypeScript layout engine for docen — zero DOM, zero rendering libraries. The inner ring (`text/`, `block/`) measures and breaks text and lays out paragraphs/tables; the outer ring builds on it (`flow/` — docx page boxing — is in; `fixed/` for pptx shape text and `grid/` for xlsx come with their editors).

Consumed by [`@docen/docx`](../docx/README.md)'s projection (DocumentOptions → LayoutDoc) and painted by [`@docen/core`](../core/README.md)'s canvas stage.

## Design

- **Input is a `LayoutDoc` projection, all-px, style-cascades already resolved.** Adapters (docx from Tiptap/ProseMirror, pptx from its shape tree, xlsx from its grid) convert their document units and resolve their style chains exactly once; the engine never sees a `styleId`, a twip, or a Tiptap node. Only layout semantics keep their OOXML shape (line rules, grid pitch, snap flags) — those are the rules this engine exists to implement.
- **FontMetrics and ShapedMeasurer are the measuring seams.** The engine works with deterministic OpenType text shaping backed by `@docen/shaping` (`createMeasurer` defaults to `ShapedMeasurer` using `rustybuzz` + `skrifa` WASM) and per-face Word-calibrated line ratios (`WORD_FONT_METRICS`). Browser `canvas.measureText` is retained as a transparent fallback and rollback option (`setShapingEnabled(false)` or `DOCEN_SHAPING_DISABLED=1`).
- **One packer for text, hard breaks, and inline pictures** — the unified breaker the DOM route never had. UAX #14 line breaking (linebreak.js) with CJK kinsoku, trailing-space hanging, first-line indent, float-zone width reduction, and per-line OOXML line-height semantics (exact / atLeast / multiple × docGrid pitch, CJK ceil snap).
- **Determinism is a contract.** Same input → same output every pass — the property the paginator's convergence depends on.

Known boundaries: rowspan cell content counts fully on its start row.

## Usage

```ts
import { browserFontMetrics, createMeasurer, layoutBlock } from "@docen/layout";

// Creates ShapedMeasurer by default for bit-exact OpenType shaping
const measurer = createMeasurer(browserFontMetrics);
const laid = layoutBlock(paragraph, 612, { linePitchPx: 25 }, measurer);
// laid.lines — y, height, positioned items, glyph runs, split points
```

## License

- [MIT](../../LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
