# RFC: Deterministic Text Stack for docen (rustybuzz + fontations)

- **Author:** docen R6 Team
- **Status:** Complete (R6.0 through R6.9 Implemented & Verified)
- **Target Branch:** `feat/r6-text-stack` (staging)
- **Date:** 2026-09-17

---

## 1. Summary

This RFC proposes replacing browser-dependent canvas `measureText` and DOM font measurement in `docen` with a fully deterministic, cross-platform text layout and font rendering engine implemented in WebAssembly via `@docen/shaping`, powered by:

1. **rustybuzz**: Complete OpenType text shaping (GSUB, GPOS, GDEF) ported from HarfBuzz.
2. **fontations (`read-fonts`, `skrifa`)**: Memory-safe font parsing, metrics extraction, and vector outline generation. (Subsetting is a hand-rolled TypeScript sfnt subsetter; see §6.)
3. **harfbuzzjs Differential Oracle**: Continuous differential testing against native HarfBuzz in CI.

---

## 2. Motivation & Current Problems

Currently, `docen` relies on:

- Canvas `measureText` via `@docen/pretext`, leading to OS- and browser-dependent layout discrepancies (Windows DirectWrite vs macOS CoreText vs Linux FreeType differ in rounding, hinting, and kerning).
- Missing OpenType features (ligatures, mark positioning, small caps, old-style numerals).
- Incomplete Arabic and complex script support (cursive joining and diacritics positioning are impossible with simple run-based `fillText`).
- PDF export without true font subsetting (Type1 and Identity-H CID references without embedded glyph subsets).

---

## 3. Spike Validation & Performance Benchmarks (R6.0)

A production WASM module (`docen_shaping.wasm`) was built using standard Rust `wasm32-unknown-unknown` without `wasm-bindgen` bloat, utilizing a typed-array linear-memory ABI.

### Empirical Spike Results

| Metric                          | Target Budget    | Spike Result                    | Status                       |
| ------------------------------- | ---------------- | ------------------------------- | ---------------------------- |
| WASM Bundle Size (Raw)          | —                | 986.3 KB                        | Healthy                      |
| WASM Bundle Size (Gzip)         | ≤ 600 KB         | **347.5 KB**                    | **Pass (42% under budget)**  |
| WASM Instantiation Latency      | < 50 ms          | **0.25 ms**                     | **Pass (200x faster)**       |
| Shaping Throughput              | ≥ 2,000 para/s   | **5,219 para/s** (835k chars/s) | **Pass (2.6x above target)** |
| Differential Parity vs HarfBuzz | 100% exact match | **Identical** (0 discrepancy)   | **Pass**                     |

### Differential Oracle Verification

- **Kerning Pairs**: Evaluated against `"AV"`, `"To"`, `"Wa"`, `"LT"`, `"Te"`, `"Vo"`, `"Yo"` — identical glyph IDs and advances to HarfBuzz.
- **Ligatures**: Evaluated against `"office"`, `"flight"`, `"waffle"`, `"fi fl ffi ffl"` — identical ligature substitution (`liga`).
- **Arabic Script**: Evaluated against `"مرحبا"`, `"العربية"`, `"كتاب"` with RTL direction — exact cursive join glyph IDs, cluster ordering, and mark positioning offsets (`xOffset`, `yOffset`).
- **Vector Outlines**: Extracted via `skrifa::outline::OutlinePen` into normalized path commands (`M`, `L`, `Q`, `C`, `Z`), ready for direct `Path2D` rendering.

---

## 4. Architecture

```
                    +---------------------------+
                    |       Font Manager        |
                    | (Bundled / Local / Cache) |
                    +-------------+-------------+
                                  |
                                  v
+---------------------------------+---------------------------------+
|                         @docen/shaping                            |
|                                                                   |
|   +-----------------------+             +---------------------+   |
|   |   RustybuzzBackend    |             |   HarfbuzzBackend   |   |
|   |  (Production WASM)    |             |  (Oracle / Testing) |   |
|   +-----------+-----------+             +----------+----------+   |
|               |                                    |              |
|               +-----------------+------------------+              |
|                                 v                                 |
|                     interface ShapingBackend                      |
+---------------------------------+---------------------------------+
                                  |
        +-------------------------+-------------------------+
        |                                                   |
        v                                                   v
+-----------------------+                           +-------------------+
|     @docen/layout     |                           |    @docen/core    |
|   (ShapedMeasurer)    |                           |  (Glyph Painter)  |
+-----------------------+                           +-------------------+
```

---

## 5. Phased Rollout Plan

- **R6.0 Spike (Complete)**: Toolchain, WASM build, `ShapingBackend`, differential oracle, benchmarks.
- **R6.1 Package `@docen/shaping` (Complete)**: Full TS API (`FontRef`, metrics, shape, outline, subset), typed-array WASM ABI.
- **R6.2 Font Manager (Complete)**: Sources, OPFS LRU cache, fallback chain per script, font identity model.
- **R6.3 Shaped Measurer (Complete)**: Integrate with `@docen/layout` as `ShapedMeasurer` (`TextMeasurer`), opt-in flag.
- **R6.4 Glyph Painting (Complete)**: Skrifa outline → `Path2D` + cache, glyph run painter, caret cluster mapping.
- **R6.7 Font Embedding (Complete)**: PDF subset (TS sfnt subsetter with rebuilt cmap/post), DOCX fontTable embed + fsType.
- **STOP-GATE T1 (Complete)**: Comprehensive verification gate before T2 (352KB gzip, 0 regressions).
- **R6.5 Complex Scripts & Bidi (Complete)**: Arabic/Hebrew/Thai/Indic/Khmer fixtures, vertical CJK metrics.
- **R6.6 OT Features & Variable Fonts (Complete)**: Real kerning, OpenType feature tags, fvar/HVAR/gvar instancing.
- **R6.8 Default Flip & Hardening (Complete)**: Worker offload, incremental cache, flip default with rollback flag.
- **R6.9 Cleanup & Documentation (Complete)**: Deduplication, font-metrics-data tidy, docs, dead code elimination.

---

## 6. Post-audit corrections (r6/fix-audit)

The R6 audit (see the R6 audit register) found several implementation claims
that did not survive independent verification. The branch now matches reality:

- **Subsetting is TypeScript, not `write-fonts`.** The subsetter rebuilds
  `cmap` (format 4 + 12) and `post` 3.0 over the remapped glyph order; CFF
  fonts pass through unmodified. It is validated by a fontTools gate.
- **The WASM ABI copies font bytes once at registration** and reads results
  through typed-array views — it is not zero-copy for font input.
- **The shaped measurer is opt-in, not the default.** The pretext line breaker
  still measures with canvas, no production font is bundled, and
  cross-environment LayoutDoc hashing is not yet covered, so the canvas path
  remains the default until those T1 gates are green.
- **Glyph runs have a real producer and consumer** (layout attaches
  `glyphRun`; the painter caches skrifa outlines and honours `unitsPerEm`),
  and the caret map consumes the same runs for cluster mapping.
- **The worker client is a real transport with fallback.** It posts correlated
  requests to a host-supplied Worker, transfers font bytes, and reports
  in-thread fallback via `client.offloaded`.
- **PDF embedding writes objects in object-number order** (the old writer
  interleaved an embedded font stream and produced a corrupt xref) and routes
  every span through the embedded font with a `/CIDToGIDMap`.
- **DOCX embedding is wired into the save path**: registered fonts become
  `word/fontTable.xml` + obfuscated `word/fonts/fontN.odttf` parts when
  `fsType` permits.
