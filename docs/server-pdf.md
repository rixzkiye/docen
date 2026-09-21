# Server-side PDF rendering (R11–R12)

Three pieces shipped, all merged:

| Piece                                  | What                                                                                    | Where                                                            |
| -------------------------------------- | --------------------------------------------------------------------------------------- | ---------------------------------------------------------------- |
| **S2 — pure-Node backend** _(default)_ | Headless renderer, no browser/DOM: layout → PaintKit → vector scene → PDF writer        | `@docen/pdf/node` — `renderPdf(json, options)`                   |
| **S3 — server facade**                 | Unified API + pluggable backend adapters                                                | `docen/pdf` — `generatePDF(json, options)` (server-only subpath) |
| **S1 — Chromium reference**            | Renders through the real editor in headless Chromium; identical to File → Export as PDF | `scripts/pdf-render.mjs` (CLI)                                   |

## Which backend?

|          | Pure-Node (`@docen/pdf/node`)                         | Chromium (`scripts/pdf-render.mjs`)                  |
| -------- | ----------------------------------------------------- | ---------------------------------------------------- |
| Browser  | **none**                                              | Chromium + Playwright                                |
| Memory   | **~84 MB at-rest**; small per-document transients     | ~1.3 GB (pool=1) … ~2.3 GB (pool=2) per browser tree |
| Fidelity | Same engine and writer (PaintKit + shared PDF writer) | Bit-identical to the in-app export                   |
| Use when | Default server export, light containers, scale        | Fallback/oracle, or reproducing app-exact output     |

Both produce deterministic bytes for the same input when a fixed `date` is passed, so both are safe to content-hash cache.

## Quick start — Node (no browser)

```ts
import { renderPdf } from "@docen/pdf/node";

const bytes = await renderPdf(docJson, {
  date: "1970-01-01T00:00:00.000Z", // fixed date → byte-identical, cacheable
  metadata: { title: "Invoice 42" },
  // Optional conformance:
  pdfa: "2b", // PDF/A-2b ("2u" | "ua" | true = "2b")
  pdfUa: true, // PDF/UA-1 (implies embedded text mode)
});

// Or via the server facade with backend adapters (default: "node"):
import { generatePDF } from "docen/pdf"; // SERVER-ONLY subpath
const bytes2 = await generatePDF(docJson, { backend: "node", date: "..." });
```

Key options: `metadata` (title/author/subject/keywords/creator/producer/creationDate/modDate), `date` (determinism), `pdfa`, `pdfUa`, `textMode` (`"outlines"` default | `"embedded"`), `language`, `outline`, `pageLabels`, `destinations`, `structElements`, `formFields`, `embeddedFonts`, `measurer`, `imageCache`.

## Server-only contract

- **`docen/pdf` is server-only.** Its `./pdf` export is gated to the `node` condition and the Chromium adapter lazily imports `node:child_process`; bundlers targeting browsers will fail to resolve it **on purpose**. Never import it from client code.
- Browser exports live in the editor and the browser-safe `@docen/pdf` main entry (`exportPdf()` on `<docen-document>`, `pagesToPdf`, `pdf-scene`, `pdf-structure`, `srgb-icc`).

## Requirements

| Requirement                                                | Notes                                                                                                                                                                                                                                         |
| ---------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Node.js                                                    | 18+ (tested v24). The Node path adds no browser dependency.                                                                                                                                                                                   |
| **veraPDF** (optional but required for conformance claims) | PDF/A-2b and PDF/UA-1 exports are produced without it, but **must be validated** with the veraPDF CLI. The fidelity harness fails when `verapdf` is missing; the unit test skips with a loud warning unless `DOCEN_REQUIRE_VERAPDF=1` is set. |
| Playwright + Chromium (Chromium path only)                 | `PLAYWRIGHT_MODULE=/path/to/playwright/index.mjs`, `CHROME_PATH=/usr/bin/chromium`.                                                                                                                                                           |
| Editor URL (Chromium path only)                            | Any page mounting `<docen-document>`: `pnpm --filter @docen/editor demo -- --port 5184 --strictPort` or a static production bundle.                                                                                                           |

### Docker dependency hint (Chromium path)

```dockerfile
FROM mcr.microsoft.com/playwright:v1.56.0-noble
ENV CHROME_PATH=/usr/bin/chromium
ENV PLAYWRIGHT_MODULE=/opt/playwright/node_modules/playwright/index.mjs
```

## Memory & throughput (measured, loaded dev machine)

- **Pure-Node:** ~84 MB RSS at-rest; a sustained 100-document loop stays flat (`packages/pdf/src/render.spec.ts`); small document ≈ 169 ms one-shot including first font/wasm load.
- **Chromium:** ~1.3 GB (pool=1) / ~2.3 GB (pool=2) per browser tree → run pool 1–2 per container and scale out; cache hits are file reads (~0–1 ms).
- Numbers are host-dependent — re-measure on your hardware before sizing containers.

## Determinism & caching

- Pass a fixed `date` (or `SOURCE_DATE_EPOCH`) → identical input yields identical bytes.
- Cache key: `sha256(json + options + date + docen version + backend)`. The S1 CLI already implements content-hash caching (`--cache`, LRU via `--cache-max-bytes`).

## Conformance (PDF/A-2b, PDF/UA-1)

- Export-side: XMP metadata (`pdfaid:part/conformance`), sRGB ICC output intent, all fonts embedded (base-14 replaced), no encryption/JS; UA adds `/Lang`, `/ViewerPreferences /DisplayDocTitle`, tagged structure + alt text.
- Verified in R12 with veraPDF 1.30.2: `PASS` for `2b` and `ua1`; 0 failed checks.
- Gates: `packages/editor/scripts/pdf-fidelity.mjs` (SSIM/PSNR/geometry/pdftotext/fonts, browser + Node parity) — run it against a dev server; it fails when veraPDF is unavailable.

## Chromium reference CLI

```bash
# 1. serve the editor (any port; the CLI default is 5184)
pnpm --filter @docen/editor demo -- --port 5184 --strictPort

# 2. render a document
PLAYWRIGHT_MODULE=/path/to/playwright/index.mjs CHROME_PATH=/usr/bin/chromium \
node scripts/pdf-render.mjs --json doc.json --out /tmp/doc.pdf

# pipe to another process (reports go to stderr)
node scripts/pdf-render.mjs --json doc.json --out - > doc.pdf
```

Main flags: `--json <file>` (repeatable), `--out <file|dir|->`, `--url <editor-url>`, `--concurrency <n>`, `--timeout <ms>`, `--date <iso|epoch-s>`, `--cache <dir>`, `--cache-max-bytes`, `--pdfa <2b|2u|ua>`, `--pdf-ua`, `--mode <outlines|embedded>`, `--repeat <n>`, `--version`. Run with `--help` for the full list.

Integration shape: your service owns the HTTP endpoint, queue/backpressure, cache, observability and deployment; docen provides the renderer (Node or Chromium) and the deterministic contract.
