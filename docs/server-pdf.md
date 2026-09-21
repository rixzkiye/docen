# Server-side PDF rendering (R11-S1)

`scripts/pdf-render.mjs` renders document JSON to PDF on the server through the
**real editor** in headless Chromium. It is the pragmatic path that is available
now, with fidelity identical to File → Export as PDF in the app; a pure-Node
backend (no browser) is the planned lightweight target (R11-S2) and the export
API will stay the same (R11-S3: `generatePDF(json, options)` in
`packages/docen`).

The renderer is deterministic (byte-identical PDFs for identical input + fixed
date) and content-address cached, so a repeated request is a file read instead
of a render.

## Requirements

| Requirement | Notes                                                                                                                                                 |
| ----------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| Node.js     | Tested on v24.14.0. No npm dependency is added to the repo.                                                                                           |
| Playwright  | The CLI imports the module named by `PLAYWRIGHT_MODULE`, default specifier `playwright`. Any Playwright >= 1.40 install works.                        |
| Chromium    | `CHROME_PATH=/usr/bin/chromium` (or let Playwright use its bundled build).                                                                            |
| Editor URL  | A page that mounts `<docen-document>`: the dev server (`pnpm --filter @docen/editor demo -- --port 5184 --strictPort`) or a static production bundle. |
| Fonts       | Document fidelity follows the OS fonts. Install the fonts your documents use, or register them on the page (`host.registerFont`) before rendering.    |

### Docker dependency hint

Chromium needs system libraries and fonts; Playwright's official image has
everything preinstalled:

```dockerfile
FROM mcr.microsoft.com/playwright:v1.56.0-noble
# or, on a distro image:
#   apt-get update && apt-get install -y chromium fonts-liberation fonts-noto-cjk
ENV CHROME_PATH=/usr/bin/chromium
ENV PLAYWRIGHT_MODULE=/opt/playwright/node_modules/playwright/index.mjs
```

Run as root by adding `--no-sandbox` (the CLI always passes
`--disable-dev-shm-usage`, which avoids the small `/dev/shm` crash).

## Quick start

```bash
# 1. serve the editor (any port; the CLI default is 5184)
pnpm --filter @docen/editor demo -- --port 5184 --strictPort

# 2. render a document
PLAYWRIGHT_MODULE=/path/to/playwright/index.mjs \
CHROME_PATH=/usr/bin/chromium \
node scripts/pdf-render.mjs \
  --json scripts/fixtures/pdf-render/demo-basic.json \
  --out /tmp/demo-basic.pdf

# pipe to another process (all reports go to stderr)
node scripts/pdf-render.mjs --json doc.json --out - > doc.pdf

# render the editor's own open demo document (no --json)
node scripts/pdf-render.mjs --out /tmp/demo.pdf
```

## CLI

```
--json <file>          document JSON (repeatable; positional paths also work)
--repeat <n>           render each input n times (benchmarks; cache dedupes when on)
--out <file|dir|->     '-' = stdout (single input); with several inputs pass a dir
--url <url>            editor URL (default http://localhost:5184/)
--concurrency <n>      editor pages in the pool (default 2)
--timeout <ms>         per-render timeout (default 60000; the page is recycled after)
--date <iso|epoch-s>   fixed PDF creation/mod date (default SOURCE_DATE_EPOCH
                       or 1970-01-01T00:00:00Z)
--version <string>     build identifier in the cache key (default: editor package
                       version + git HEAD)
--cache <dir>          cache directory (default /tmp/opencode/pdf-render-cache)
--no-cache             disable the cache
--cache-max-bytes <n>  LRU evict above this size (default 512 MiB; 0 = never)
--no-sandbox           Chromium --no-sandbox (root/containers)
--keep-open            keep the browser open after rendering (debug)
```

Exit codes: `0` ok, `1` render/runtime failure, `2` usage error, `130`
interrupted (SIGINT finishes in-flight renders, waits up to 5 s, then closes
Chromium).

Every render logs one line to stderr:

```
pdf-render: [4/12] miss demo-basic-r4      wall=1321ms render=406ms export=893ms pages=2 bytes=267189 rss=159.3MB chromium=1606.6MB/8p
```

`render` = in-page `setJSON` + repaginate; `export` = `host.exportPdf()`;
`rss` = Node process RSS after the render; `chromium` = whole Chromium process
tree RSS (Linux; summed over the browser tree via `/proc`, null elsewhere).
The summary line carries min/median/p95 wall time and cache-hit median.

## Determinism

`formatPdfDate` (the `/Info` dictionary) is the exporter's only wall-clock
input. The CLI pins it with `--date`, so **the same JSON always produces the
same SHA-256** — verified by rendering the same fixture twice in different
processes (`1d888e0b…`, 267,189 bytes, identical) and 12× in one pool run.
Set `SOURCE_DATE_EPOCH` instead of `--date` if your build system already
provides it.

Cache keys are `SHA-256(stableStringify({ schema, build, json, export options,
date }))`, i.e. document content + editor build + export options + fixed date.
Because the date is in the key, a cache hit is byte-identical to a fresh render
by construction.

## Cache

- `<key>.pdf` + `<key>.json` sidecar (build, date, page count, wall time, byte
  length, SHA-256). PDFs are written to a temp file and renamed, so a crashed
  writer never leaves a half PDF under the final name.
- A hit validates the sidecar key and byte length, then reads the PDF; it never
  launches Chromium. Measured: **0–1 ms** vs ~1.5 s median render.
- Identical inputs inside one run (`--repeat`, duplicate `--json`) share a
  single render when the cache is on (`dup` in the log). With `--no-cache`,
  `--repeat` measures N real renders.
- Eviction is LRU by mtime once the directory exceeds `--cache-max-bytes`
  (removes oldest entries down to 90% of the limit). Accounting is approximate
  (sidecar sizes are not subtracted); the directory is disposable, so deleting
  it is always safe.
- Verified key behaviour: same input + `--version` → hit; changing `--version`
  or `--date` → miss (both are part of the key).

## Measured memory & throughput

Environment: Linux, 8 vCPU, `/usr/lib/chromium` + the repo's Playwright,
**machine load average ~8–14** (several other agent lanes active during the
run), fixture `scripts/fixtures/pdf-render/demo-basic.json` (A4, 2 pages,
267,189 bytes). Treat absolute times as an upper bound typical for a contended
CI box.

| Scenario                                             | Result                                                                                                                                                 |
| ---------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Sequential, `--no-cache --concurrency 1 --repeat 12` | cold r1 **4213 ms**; warm **min 1063 / median 1552 / p95 (12 incl. cold) 4213 ms**; all 12 PDFs byte-identical                                         |
| Node RSS, pool=1                                     | **171 → 183 MB** over 12 renders (slow growth, no plateau yet)                                                                                         |
| Chromium tree RSS, pool=1                            | **~1.6 GB / 8 processes** (main renderer ~490 MB, network service ~145 MB, GPU ~82 MB)                                                                 |
| `--concurrency 2 --repeat 10`                        | per-render wall **median 1821 ms, p95 4255 ms**; sum 21.5 s → **~0.93 renders/s** throughput; CPU-bound (layout + export), so 2 pages ≈ 1.5× of 1 page |
| Chromium tree RSS, pool=2                            | **~2.3 GB / 9 processes** (+~0.7 GB per added editor page)                                                                                             |
| Cache hit (separate process, 12/12)                  | **0–1 ms**, Node RSS **58 MB**, no browser launched                                                                                                    |
| Demo document (4 pages, images + charts)             | cold **3995 ms** (render 830 ms, export 2887 ms), 861,518 bytes, Node RSS 162 MB, Chromium ~1.4 GB                                                     |

Cross-check (plan gate): a direct in-page `host.exportPdf({ metadata })` render
of the fixture produces the same 267,189 bytes with SHA-256
`1d888e0b1b475d75c3cdc6c143810a98a027038eb3ed3ab51c346b2e76fcfb12` as the CLI
output for the same input and date (`cmp` byte-identical).

Concurrency guidance:

- The browser tree dominates memory, not Node. Budget **~1.1–1.6 GB per
  browser** plus **~0.6–0.7 GB per additional editor page** on this app, and
  prefer **pool 1–2 per container** with horizontal replicas over a large pool.
- Rendering is CPU-bound; concurrency > cores only adds contention. Use
  `--concurrency 1` for predictable latency and 2 for modest overlap.
- Cold start is 4–7 s per pool (page load + first layout/export); keep the
  process warm if you serve a steady stream, and remember the cache skips the
  browser entirely for repeats.

## Integrating with Fastify

Calling the CLI per request keeps the renderer isolated and matches the cache
semantics. Keep a small queue so at most `CONCURRENCY` child processes run —
each child is a Chromium instance, so this is the memory throttle.

```js
import { spawn } from "node:child_process";
import { mkdtempSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const RENDERER = new URL("../scripts/pdf-render.mjs", import.meta.url).pathname;
const CONCURRENCY = 2;
let active = 0;
const waiting = [];
const acquire = () =>
  active < CONCURRENCY
    ? (active++, Promise.resolve())
    : new Promise((resolve) => waiting.push(() => (active++, resolve())));
const release = () => {
  active--;
  waiting.shift()?.();
};

function renderPdf(json, { date = "1970-01-01T00:00:00Z" } = {}) {
  const dir = mkdtempSync(join(tmpdir(), "docen-pdf-"));
  const input = join(dir, "doc.json");
  writeFileSync(input, JSON.stringify(json));
  return new Promise((resolve, reject) => {
    const child = spawn(
      process.execPath,
      [RENDERER, "--json", input, "--out", "-", "--date", date],
      {
        env: process.env, // PLAYWRIGHT_MODULE / CHROME_PATH
        stdio: ["ignore", "pipe", "pipe"],
      },
    );
    const chunks = [];
    let stderr = "";
    child.stdout.on("data", (c) => chunks.push(c));
    child.stderr.on("data", (c) => (stderr += c));
    child.on("error", reject);
    child.on("close", (code) => {
      rmSync(dir, { recursive: true, force: true });
      if (code === 0) resolve(Buffer.concat(chunks));
      else reject(new Error(`pdf-render exited ${code}: ${stderr.trim()}`));
    });
  });
}

export default async function pdfRoutes(app) {
  app.post("/pdf", async (req, reply) => {
    await acquire();
    try {
      const pdf = await renderPdf(req.body, { date: req.body.date ?? "1970-01-01T00:00:00Z" });
      reply.header("content-type", "application/pdf").send(pdf);
    } finally {
      release();
    }
  });
}
```

Notes for production:

- Pass a stable per-document `--date` (e.g. the document's modified time) to
  make the PDF cacheable; omit it and every request is a cache miss.
- Point `--cache` at a mounted volume and cap it with `--cache-max-bytes`.
- Set `--timeout` below your HTTP timeout. A timed-out page is recycled for the
  next request.
- `host.exportPdf({ metadata })` accepts an optional `metadata`
  (`creationDate`/`modDate`/title/author/…) — an additive, backwards-compatible
  hook added by this lane; the menu path still uses the wall clock.

## Roadmap

- **R11-S2** replaces the browser with a pure-Node layout → paint → PDF path
  (target 50–150 MB/process, no Chromium).
- **R11-S3** exposes `generatePDF(json, options)` / `renderPDFStream` from
  `packages/docen`, backend-configurable (`node | chromium`). The JSON-in,
  bytes-out contract documented above does not change.
