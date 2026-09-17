# @docen/shaping

![npm version](https://img.shields.io/npm/v/@docen/shaping)
![npm downloads](https://img.shields.io/npm/dw/@docen/shaping)
![npm license](https://img.shields.io/npm/l/@docen/shaping)

> Deterministic WebAssembly text shaping and font typography engine for docen. Backed by [`rustybuzz`](https://github.com/RazrFalcon/rustybuzz) (HarfBuzz-compatible OpenType shaping), [`read-fonts`](https://github.com/googlefonts/fontations) and [`skrifa`](https://github.com/googlefonts/fontations) (metrics, glyph vector outlines, variable axes), and zero-copy TrueType sfnt subsetting.

Consumed by [`@docen/layout`](../layout/README.md)'s `ShapedMeasurer` for bit-exact layout advances and [`@docen/core`](../core/README.md)'s canvas stage for vector glyph rendering.

---

## Features

- **Cross-Platform Determinism:** Replaces browser-divergent `CanvasRenderingContext2D.measureText` with bit-exact OpenType shaping identical across Chromium, Firefox, Safari, and Node.js SSR.
- **Complex Scripts & Direction:**
  - Multi-direction: LTR, RTL, TTB (vertical CJK with `vhea` / OpenType synthesis), BTT, and automatic Unicode script detection.
  - Full GSUB/GPOS shaping for Arabic cursive joining & mark positioning, Hebrew RTL Niqqud, Thai vowel reordering & mark stacking, Devanagari conjuncts & matras, Khmer coeng subscripts, and Latin ligature substitution (`fi`, `fl`, `ffi`, `ffl`).
  - Verified against HarfBuzz CLI (`hb-shape`) and `harfbuzzjs` differential oracle.
- **OpenType Features & Variable Fonts:**
  - Standard and discretionary ligatures (`liga`, `clig`, `calt`, `dlig`, `hlig`).
  - Small caps (`smcp`) and all-caps (`c2sc`).
  - Kerning control (`kern`).
  - Numerical forms & spacing (`onum`, `lnum`, `tnum`, `pnum`).
  - Stylistic sets (`ss01` through `ss20`) and custom feature overrides.
  - Variable font instancing across weight (`wght`), width (`wdth`), slant (`slnt`), and custom variation axes with dynamic metrics and outline interpolation.
- **Font Management & Fallback:**
  - Multi-tiered font sources: `BundledFontSource`, `MemoryFontSource`, and `LocalFontSource` (Web Local Font Access API).
  - Content-addressed SHA-256 caching with LRU memory eviction and optional browser OPFS persistence.
  - Script-aware fallback chains (`Latn`, `Arab`, `Hani`, `Hebr`, `Thai`, `Deva`, `Khmr`) with missing-glyph diagnostics.
- **Vector Outlines:** Extracts true Bézier glyph paths (`M`, `L`, `Q`, `C`, `Z`) via Skrifa for high-DPI vector canvas rendering.
- **Embedding & Subsetting:**
  - TrueType sfnt table subsetter with composite glyph closure, reducing font size by >66%.
  - Adobe ToUnicode CMap generation for searchable PDF exports.
  - ECMA-376 Part 4 §14.2.14 GUID obfuscation for DOCX `word/fontTable.xml`.
  - OpenType `fsType` licensing validation blocking restricted fonts (`0x0002`).
- **Worker Offloading:** Dedicated `ShapingWorkerHandler` and `ShapingWorkerClient` for non-blocking asynchronous shaping in background Web Workers.
- **Zero-Copy WASM ABI:** Direct shared WebAssembly linear memory layout with typed array views, achieving sub-millisecond startup and >3,000 paragraphs/second throughput.

---

## Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                    @docen/layout                             │
│       createMeasurer() ──► ShapedMeasurer (Pretext)          │
└───────────────────────────────┬──────────────────────────────┘
                                │ calls
┌───────────────────────────────▼──────────────────────────────┐
│                    @docen/shaping                            │
│  FontManager ──► FontRef ──► RustybuzzBackend                │
│    │               │              │                          │
│  FontCache       Subsetter        │ WebAssembly ABI          │
│  (LRU + OPFS)   + ToUnicode       │ (Zero-Copy Pointers)     │
└───────────────────────────────────┼──────────────────────────┘
                                    │
┌───────────────────────────────────▼──────────────────────────┐
│                   docen_shaping.wasm                         │
│  ┌────────────────────────┐    ┌──────────────────────────┐  │
│  │   rustybuzz 0.20.1     │    │ read-fonts / skrifa 0.47 │  │
│  │   (OpenType Shaping)   │    │ (Metrics, Outlines, Var) │  │
│  └────────────────────────┘    └──────────────────────────┘  │
└──────────────────────────────────────────────────────────────┘
```

---

## Quick Start

### 1. Basic Shaping

```ts
import { loadShapingWasm, createFontRef } from "@docen/shaping";

// Load and initialize WASM runtime
await loadShapingWasm();

// Wrap font bytes into a FontRef
const fontRef = createFontRef(fontBuffer);

// Shape text with OpenType features
const result = fontRef.shape("Office aesthetic fi 12345", {
  direction: "ltr",
  script: "Latn",
  features: [
    { tag: "kern", value: 1 },
    { tag: "liga", value: 1 },
    { tag: "onum", value: 1 },
  ],
});

console.log(`Total advance: ${result.totalAdvance} design units`);
for (const glyph of result.glyphs) {
  console.log(`Glyph ID: ${glyph.glyphId}, Advance: ${glyph.xAdvance}, Cluster: ${glyph.cluster}`);
}
```

### 2. Variable Fonts

```ts
// Inspect axes
const axes = fontRef.getAxes();
console.log(axes); // [{ tag: 'wght', min: 100, max: 900, default: 400 }]

// Shape with custom variations
const boldRun = fontRef.shape("Bold Run", {
  variations: [{ tag: "wght", value: 750 }],
});
```

### 3. Font Subsetting & CMap for PDF Export

```ts
import { subsetFont, generateToUnicodeCMap } from "@docen/shaping";

// Collect glyph IDs used in document
const usedGlyphs = [0, 15, 23, 42, 88];

// Produce compact subset font buffer
const subsetBuffer = subsetFont(fontBuffer, usedGlyphs);

// Generate Adobe ToUnicode CMap mapping glyph IDs to UTF-16
const cmapStream = generateToUnicodeCMap(usedGlyphs, glyphToUnicodeMap);
```

### 4. Background Web Worker Offload

```ts
import { ShapingWorkerClient } from "@docen/shaping";

const worker = new Worker(new URL("./shaping.worker.ts", import.meta.url), { type: "module" });
const client = new ShapingWorkerClient(worker);

await client.init();
await client.registerFont(1, fontBuffer);

const result = await client.shape(1, "Text shaped off-thread", {
  direction: "auto",
});
```

---

## Performance Budgets

| Metric                         | Budget         | Measured           | Status |
| ------------------------------ | -------------- | ------------------ | ------ |
| **WASM Binary Size (gzip)**    | ≤ 600 KB       | **~352 KB**        | PASSED |
| **WASM Instantiation Latency** | < 50 ms        | **~0.3 ms**        | PASSED |
| **Shaping Throughput**         | ≥ 2,000 para/s | **> 3,100 para/s** | PASSED |
| **Subsetting Size Reduction**  | > 50%          | **> 66%**          | PASSED |

---

## License

- `@docen/shaping` is licensed under [MIT](../../LICENSE) &copy; [Demo Macro](https://www.demomacro.com/).
- Embedded Rust dependencies: `rustybuzz` (MIT), `read-fonts` (MIT / Apache-2.0), `skrifa` (MIT / Apache-2.0). See [NOTICE](./NOTICE) for third-party notices.
