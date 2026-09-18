# @docen/shaping

![npm version](https://img.shields.io/npm/v/@docen/shaping)
![npm downloads](https://img.shields.io/npm/dw/@docen/shaping)
![npm license](https://img.shields.io/npm/l/@docen/shaping)

> Deterministic WebAssembly text shaping and font typography engine for docen. Backed by [`rustybuzz`](https://github.com/RazrFalcon/rustybuzz) (HarfBuzz-compatible OpenType shaping), [`read-fonts`](https://github.com/googlefonts/fontations) and [`skrifa`](https://github.com/googlefonts/fontations) (metrics, glyph vector outlines, variable axes), and a TrueType sfnt subsetter that rebuilds `cmap`/`post` for remapped glyph IDs.

Consumed by [`@docen/layout`](../layout/README.md)'s `ShapedMeasurer` for bit-exact layout advances and [`@docen/core`](../core/README.md)'s canvas stage for vector glyph rendering.

---

## Features

- **Cross-Platform Determinism:** Replaces browser-divergent `CanvasRenderingContext2D.measureText` with OpenType shaping. The shaped measurer is the default for registered fonts; `@docen/layout` bundles metric-compatible production faces (Carlito/Caladea/Liberation) and registers them through `registerDefaultFonts()`, so the Word default families never silently fall back to canvas — an unregistered family logs a one-time warning. `setShapingEnabled(false)` or `DOCEN_SHAPING_DISABLED=1` rolls back to the canvas measurer.
- **Complex Scripts & Direction:**
  - Multi-direction: LTR, RTL, TTB (vertical CJK with `vhea` / OpenType synthesis), BTT, and automatic Unicode script detection.
  - Full GSUB/GPOS shaping for Arabic cursive joining & mark positioning, Hebrew RTL Niqqud, Thai vowel reordering & mark stacking, Devanagari conjuncts & matras, Khmer coeng subscripts, and Latin ligature substitution (`fi`, `fl`, `ffi`, `ffl`).
  - Verified by committed golden records regenerated with `test/generate-goldens.mjs` (harfbuzzjs oracle) and a live rustybuzz-vs-harfbuzzjs differential suite.
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
  - TrueType sfnt subsetter with composite glyph closure that rewrites `glyf`/`loca`/`hmtx`/`maxp`/`hhea`/`head`, rebuilds `cmap` (format 4 + 12) over the retained code points and rewrites `post` as version 3.0. `subsetFontWithPlan` returns the old→new glyph map for PDF `CIDToGIDMap` consumers.
  - Adobe ToUnicode CMap generation for searchable PDF exports.
  - ECMA-376 Part 4 §14.2.14 GUID obfuscation for DOCX `word/fontTable.xml`.
  - OpenType `fsType` licensing validation blocking restricted fonts (`0x0002`).
- **Worker Offloading:** `ShapingWorkerHandler` runs the engine inside any worker; `ShapingWorkerClient` speaks the correlated request/response protocol over a host-supplied `Worker`, transfers font bytes, and falls back to in-thread shaping (reported by `client.offloaded`) when no worker is supplied or initialization fails.
- **Typed-Array WASM ABI:** Glyph, metrics, outline and string results are read through typed-array views over the WASM linear memory; font bytes are copied once at registration and deallocated immediately.

---

## Architecture

```
┌──────────────────────────────────────────────────────────────┐
│                    @docen/layout                             │
│       createMeasurer() ──► ShapedMeasurer (opt-in)           │
└───────────────────────────────┬──────────────────────────────┘
                                │ calls
┌───────────────────────────────▼──────────────────────────────┐
│                    @docen/shaping                            │
│  FontManager ──► FontRef ──► RustybuzzBackend                │
│    │               │              │                          │
│  FontCache       Subsetter        │ WebAssembly ABI          │
│  (LRU + OPFS)   + ToUnicode       │ (typed-array views)      │
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
import { initShapingWasm, createFontRefSync } from "@docen/shaping";

// Load and initialize WASM runtime
await initShapingWasm();

// Wrap font bytes into a FontRef (or `await createFontRef(bytes)`)
const fontRef = createFontRefSync(fontBuffer);

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

### 2. Font Collections (.ttc/.otc)

```ts
// `index` selects one face of a collection (Noto Sans CJK, Windows msyh.ttc):
const sc = createFontRefSync(collectionBytes, { index: 2 });
console.log(sc.familyName); // "Noto Sans CJK SC"

// The export-only subsetter loads on first use and refuses collections:
const subset = await sc.subset([...]); // throws for a collection face
```

### 3. Variable Fonts

```ts
// Inspect axes
const axes = fontRef.getAxes();
console.log(axes); // [{ tag: 'wght', min: 100, max: 900, default: 400 }]

// Shape with custom variations
const boldRun = fontRef.shape("Bold Run", {
  variations: [{ tag: "wght", value: 750 }],
});
```

### 4. Font Subsetting & CMap for PDF Export

```ts
import { readCmap, subsetFontWithPlan, generateToUnicodeCMap } from "@docen/shaping";

// Collect the old glyph IDs used in the document (cmap lookup or shaped runs)
const cmap = readCmap(fontBuffer);
const usedGlyphs = [0, ...[...text].map((ch) => cmap.get(ch.codePointAt(0)!))];
const { data: subsetBuffer, glyphMap } = subsetFontWithPlan(fontBuffer, usedGlyphs);

// Generate an Adobe ToUnicode CMap mapping CIDs to Unicode
const cmapStream = generateToUnicodeCMap(new Map([[1, 0x0041]]));
```

`subsetFont(buffer, glyphIds)` is the bytes-only form. CFF/bitmap fonts pass
through unchanged (valid bytes; glyph IDs are not remapped), so embedders must
use the returned `glyphMap` to address TrueType subsets.

### 5. Background Web Worker Offload

Write the worker entry (five lines) and hand the Worker to the client; without
one — or when it fails to initialize — the client shapes in-thread:

```ts
// shaping.worker.ts
import { ShapingWorkerHandler } from "@docen/shaping";
const handler = new ShapingWorkerHandler();
self.onmessage = (event) =>
  handler.handleMessage(event.data, (response, transfer) => self.postMessage(response, transfer));
```

```ts
import { ShapingWorkerClient } from "@docen/shaping";

const worker = new Worker(new URL("./shaping.worker.ts", import.meta.url), { type: "module" });
const client = new ShapingWorkerClient(worker);

await client.init(); // falls back in-thread on worker failure
await client.registerFont(fontBuffer);

const result = await client.shape(fontId, "Text shaped off-thread", {
  direction: "auto",
});
```

---

## Bundler / Next.js setup

`initShapingWasm()` loads the shipped `wasm/docen_shaping.wasm` from
`new URL("../wasm/docen_shaping.wasm", import.meta.url)` — in Node it reads the
file from the package, in the browser it `fetch`es the emitted asset URL. The
asset ships inside the npm tarball (`@docen/shaping/wasm/docen_shaping.wasm` is
an exported subpath), so no postinstall step is required.

Bundlers that understand `new URL(..., import.meta.url)` (webpack 5, Vite,
Turbopack) emit/copy the `.wasm` automatically. Next.js is the common case that
needs explicit guidance:

- **webpack (default in Next.js):** `new URL` asset references are supported
  out of the box. If a custom webpack config turns `.wasm` into a module, keep
  it an asset instead:

  ```js
  // next.config.mjs
  export default {
    webpack(config) {
      config.module.rules.push({ test: /docen_shaping\.wasm$/, type: "asset/resource" });
      return config;
    },
  };
  ```

- **Turbopack (`next dev --turbopack`):** Turbopack's built-in `.wasm` rule
  treats the file as a WebAssembly ES module, which is not what the loader
  fetches. Register it as a URL asset:

  ```js
  // next.config.mjs
  export default {
    turbopack: {
      rules: {
        "*.wasm": { type: "asset" },
      },
    },
  };
  ```

- **Any other environment / explicit control:** pass the bytes yourself — the
  loader accepts them and skips its default resolution entirely:

  ```ts
  import { readFileSync } from "node:fs";
  import { createRequire } from "node:module";
  import { initShapingWasm } from "@docen/shaping";

  // Node / server runtime
  const wasmPath = createRequire(import.meta.url).resolve("@docen/shaping/wasm/docen_shaping.wasm");
  await initShapingWasm(readFileSync(wasmPath));

  // Browser: fetch your own copy of the asset
  // await initShapingWasm(await (await fetch("/assets/docen_shaping.wasm")).arrayBuffer());
  ```

---

## Performance Budgets

Measured on the audited fix branch (`r6/fix-audit`) with Node 24, rustc 1.97.1;
throughput is a short-round peak/median pair because the suite runs
concurrently.

| Metric                         | Budget                         | Measured                                                            | Status                   |
| ------------------------------ | ------------------------------ | ------------------------------------------------------------------- | ------------------------ |
| **WASM Binary Size (gzip)**    | ≤ 600 KB                       | **~360 KB**                                                         | PASSED                   |
| **WASM Instantiation Latency** | < 50 ms                        | **~0.25 ms**                                                        | PASSED                   |
| **Shaping Throughput**         | ≥ 2,000 para/s                 | **~5,000–8,000 para/s isolated; ~2,400 peak under full-suite load** | PASSED                   |
| **Per-keystroke layout**       | ≤ 10% regression               | **best round 0.24–0.59×, under-load median 1.0–2.1×**               | PASSED (best-round gate) |
| **LayoutDoc determinism**      | identical hash                 | repeated + fresh-WASM runs hash equal                               | PASSED (in-process)      |
| **Subsetting validity**        | fontTools parse + cmap/advance | every fixture subset passes the fontTools gate                      | PASSED                   |
| **PDF text extraction**        | 100% (Latin + accents + CJK)   | `pdftotext` byte-exact on subset-embedded PDF                       | PASSED                   |

Known gaps, tracked for R6.T2: the pretext line breaker still measures through
canvas (`prepareRichInline`), so shaped advances only drive atom probes today.
The production font set is bundled by `@docen/layout` (`registerDefaultFonts()`)
with its metrics pinned in the cross-environment golden
(`packages/layout/test/font-metrics-golden.json`).

---

## Reproducible WASM build

The committed `wasm/docen_shaping.wasm` is built with the Rust toolchain pinned
in [`rust-toolchain.toml`](./rust-toolchain.toml) (currently `1.97.1`,
`wasm32-unknown-unknown`). The artifact hash is recorded in
`wasm/docen_shaping.wasm.sha256` and verified by
[`test/reproducibility.spec.ts`](./test/reproducibility.spec.ts).

```bash
# Rebuild + verify against the recorded hash
pnpm --filter @docen/shaping build:wasm

# CI check: hash the committed artifact only (no toolchain required)
pnpm --filter @docen/shaping check:wasm

# Intentional crate changes: rebuild and record the new hash
DOCEN_WASM_UPDATE=1 pnpm --filter @docen/shaping build:wasm
```

A build with a different toolchain/dependency resolution produces a different
hash; the check fails loudly instead of silently shipping an unreproducible
binary. The script refuses to overwrite the recorded hash unless
`DOCEN_WASM_UPDATE=1` is explicit.

## License

- `@docen/shaping` is licensed under [MIT](../../LICENSE) &copy; [Demo Macro](https://www.demomacro.com/).
- Embedded Rust dependencies: `rustybuzz` (MIT), `read-fonts` (MIT / Apache-2.0), `skrifa` (MIT / Apache-2.0). See [NOTICE](./NOTICE) for third-party notices.
