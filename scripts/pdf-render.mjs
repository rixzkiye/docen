#!/usr/bin/env node
/**
 * docen server-side PDF render service (R11-S1).
 *
 * Renders document JSON to PDF through the real editor (headless Chromium via
 * Playwright): one persistent browser, a pool of editor pages, a public
 * `host.setJSON` → `host.exportPdf` round-trip per request, and a content-hash
 * cache in front. Deterministic with a fixed `--date` (the PDF Info dict is
 * the only wall-clock input), so identical inputs produce identical bytes.
 *
 * Requirements (zero npm deps):
 *   PLAYWRIGHT_MODULE  path to a playwright ESM entry (default: "playwright")
 *   CHROME_PATH        Chromium executable (else Playwright's bundled one)
 *   a running editor URL (dev server or static bundle) — see docs/server-pdf.md
 *
 * Usage:
 *   node scripts/pdf-render.mjs --json doc.json --out doc.pdf
 *   node scripts/pdf-render.mjs --json doc.json --out - > doc.pdf
 *   node scripts/pdf-render.mjs --json a.json --json b.json --out /tmp/pdfs
 *   node scripts/pdf-render.mjs --json doc.json --repeat 12 --no-cache --out /tmp/bench
 *   node scripts/pdf-render.mjs                  # render the editor's open demo document
 *
 * Exit codes: 0 ok, 1 render/runtime failure, 2 usage error, 130 interrupted.
 * All reporting goes to stderr so `--out -` pipes a clean PDF on stdout.
 */
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { basename, dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const CACHE_SCHEMA = 1;
const DEFAULT_CACHE = "/tmp/opencode/pdf-render-cache";
const DEFAULT_URL = "http://localhost:5184/";
const DEFAULT_TIMEOUT_MS = 60_000;
const DEFAULT_CACHE_MAX_BYTES = 512 * 1024 * 1024;
const SIGINT_GRACE_MS = 5_000;

// ── CLI ─────────────────────────────────────────────────────────────────────

const HELP = `Usage: node scripts/pdf-render.mjs [options] [document.json ...]

Inputs (one or more; no input renders the editor's open demo document):
  --json <file>          document JSON file (repeatable)
  --repeat <n>           render each input n times (default 1; benchmarks/cache)

Output:
  --out <file|dir|->     PDF path; '-' = stdout (default for a single input);
                         a directory (default with multiple inputs) names each
                         PDF after its input; without --out, multiple inputs
                         write <input>.pdf beside each source file

Render pool:
  --url <url>            editor URL (default ${DEFAULT_URL})
  --concurrency <n>      editor pages in the pool (default 2)
  --timeout <ms>         per-render timeout (default ${DEFAULT_TIMEOUT_MS})
  --date <iso|epoch-s>   fixed PDF creation/mod date; default SOURCE_DATE_EPOCH
                         or 1970-01-01T00:00:00Z (deterministic + cacheable)
  --version <string>     build identifier in the cache key (default: editor
                         package version + git HEAD of this repo)
  --no-sandbox           launch Chromium with --no-sandbox (root/containers)
  --keep-open            keep the browser open after rendering (debug)

Cache (default ${DEFAULT_CACHE}):
  --cache <dir>          cache directory
  --no-cache             disable the cache
  --cache-max-bytes <n>  evict oldest entries above this size (0 = never)

  -h, --help             this text`;

const argv = process.argv.slice(2);
const has = (name) => argv.includes(name);
const takeValues = (name) => {
  const out = [];
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === name && argv[i + 1] !== undefined) out.push(argv[i + 1]);
  }
  return out;
};
const takeValue = (name, fallback) => takeValues(name).at(-1) ?? fallback;

if (has("-h") || has("--help")) {
  console.log(HELP);
  process.exit(0);
}

const VALUE_FLAGS = new Set([
  "--json",
  "--out",
  "--url",
  "--concurrency",
  "--timeout",
  "--date",
  "--version",
  "--cache",
  "--cache-max-bytes",
  "--repeat",
]);
const positional = argv.filter((a, i) => !a.startsWith("-") && !VALUE_FLAGS.has(argv[i - 1]));
const jsonPaths = [...takeValues("--json"), ...positional].map((p) => resolve(p));

const usageError = (message) => {
  console.error(`pdf-render: ${message}\n\n${HELP}`);
  process.exit(2);
};

const intFlag = (name, fallback, { min = 1 } = {}) => {
  const raw = takeValue(name);
  if (raw === undefined) return fallback;
  const n = Number(raw);
  if (!Number.isInteger(n) || n < min)
    usageError(`${name} expects an integer >= ${min}, got "${raw}"`);
  return n;
};

const url = takeValue("--url", DEFAULT_URL);
const concurrency = intFlag("--concurrency", 2);
const repeat = intFlag("--repeat", 1);
const timeoutMs = intFlag("--timeout", DEFAULT_TIMEOUT_MS, { min: 1000 });
const noCache = has("--no-cache");
const cacheDir = resolve(takeValue("--cache", DEFAULT_CACHE));
const cacheMaxBytes = intFlag("--cache-max-bytes", DEFAULT_CACHE_MAX_BYTES, { min: 0 });
const outFlag = takeValue("--out");
const keepOpen = has("--keep-open");
const noSandbox = has("--no-sandbox");

const parseDate = (raw) => {
  if (/^\d+$/.test(raw)) return new Date(Number(raw) * 1000);
  const d = new Date(raw);
  if (Number.isNaN(d.getTime()))
    usageError(`--date is not a valid ISO date or epoch seconds: "${raw}"`);
  return d;
};
const exportDate = parseDate(
  takeValue("--date", process.env.SOURCE_DATE_EPOCH ?? "1970-01-01T00:00:00Z"),
);
const exportDateISO = exportDate.toISOString();

// ── small helpers ───────────────────────────────────────────────────────────

const sha256 = (data) => createHash("sha256").update(data).digest("hex");
const mb = (bytes) => `${(bytes / 1048576).toFixed(1)}MB`;

/** Canonical JSON (sorted keys) so semantically equal documents share a cache
 *  key regardless of key order in the source file. */
function stableStringify(value) {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const keys = Object.keys(value).sort();
  return `{${keys.map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`).join(",")}}`;
}

/** Build identifier in the cache key: editor package version + repo commit.
 *  Overridable with --version (e.g. a container image tag in production). */
function buildIdentifier() {
  const override = takeValue("--version");
  if (override) return override;
  const root = fileURLToPath(new URL("..", import.meta.url));
  let version = "unknown";
  try {
    version = JSON.parse(readFileSync(join(root, "packages/editor/package.json"), "utf8")).version;
  } catch {
    // keep "unknown"
  }
  let commit = "unknown";
  try {
    commit = execFileSync("git", ["-C", root, "rev-parse", "--short", "HEAD"], {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "ignore"],
    }).trim();
  } catch {
    // keep "unknown"
  }
  return `@docen/editor@${version}+${commit}`;
}

const buildId = buildIdentifier();

/** Total RSS of the Chromium process tree, located by a unique launch switch
 *  (custom switches only show on the main process — descendants are found by
 *  walking /proc ppid links). Returns null off Linux or when not found. */
function chromiumTreeRss(token) {
  if (process.platform !== "linux") return null;
  const children = new Map();
  const roots = [];
  for (const entry of readdirSync("/proc")) {
    if (!/^\d+$/.test(entry)) continue;
    const pid = Number(entry);
    try {
      const stat = readFileSync(`/proc/${pid}/stat`, "utf8");
      const ppid = Number(stat.slice(stat.lastIndexOf(")") + 2).split(" ")[1]);
      if (!children.has(ppid)) children.set(ppid, []);
      children.get(ppid).push(pid);
      if (readFileSync(`/proc/${pid}/cmdline`, "utf8").includes(token)) roots.push(pid);
    } catch {
      // process vanished mid-scan
    }
  }
  if (roots.length === 0) return null;
  let totalPages = 0;
  const stack = [...roots];
  const seen = new Set();
  while (stack.length > 0) {
    const pid = stack.pop();
    if (seen.has(pid)) continue;
    seen.add(pid);
    try {
      totalPages += Number(readFileSync(`/proc/${pid}/statm`, "utf8").split(" ")[1]);
    } catch {
      // gone
    }
    stack.push(...(children.get(pid) ?? []));
  }
  return { bytes: totalPages * 4096, processes: seen.size };
}

// ── content-hash cache ──────────────────────────────────────────────────────

class PdfCache {
  #bytes = 0;

  constructor(dir, maxBytes) {
    this.dir = dir;
    mkdirSync(dir, { recursive: true });
    this.maxBytes = maxBytes;
    for (const file of readdirSync(dir)) {
      try {
        this.#bytes += statSync(join(dir, file)).size;
      } catch {
        // racing writer
      }
    }
  }

  static keyFor(json) {
    return sha256(
      stableStringify({
        schema: CACHE_SCHEMA,
        build: buildId,
        json,
        export: { tagged: true, textMode: "outlines", creationDate: exportDateISO },
        date: exportDateISO,
      }),
    );
  }

  #paths(key) {
    return { pdf: join(this.dir, `${key}.pdf`), meta: join(this.dir, `${key}.json`) };
  }

  /** Cache hit → { bytes, meta, ms }; miss or corrupt entry → null. */
  get(key) {
    const t0 = performance.now();
    const { pdf, meta } = this.#paths(key);
    if (!existsSync(pdf) || !existsSync(meta)) return null;
    try {
      const bytes = readFileSync(pdf);
      const info = JSON.parse(readFileSync(meta, "utf8"));
      if (info.key !== key || info.pdfBytes !== bytes.byteLength) return null;
      return { bytes, meta: info, ms: performance.now() - t0 };
    } catch {
      return null;
    }
  }

  put(key, bytes, info) {
    const { pdf, meta } = this.#paths(key);
    if (existsSync(pdf)) {
      // Replacing an entry — drop its accounted size first.
      try {
        this.#bytes -= statSync(pdf).size;
      } catch {
        // gone
      }
    }
    const tmp = `${pdf}.tmp-${process.pid}`;
    writeFileSync(tmp, bytes);
    writeFileSync(
      meta,
      JSON.stringify(
        {
          ...info,
          schema: CACHE_SCHEMA,
          key,
          pdfBytes: bytes.byteLength,
          pdfSha256: sha256(bytes),
        },
        null,
        2,
      ),
    );
    renameSync(tmp, pdf);
    try {
      this.#bytes += statSync(pdf).size + statSync(meta).size;
    } catch {
      // approximate accounting if the stat races
    }
    if (this.maxBytes > 0) this.#evict();
  }

  /** LRU by mtime: remove whole entries until comfortably under the limit. */
  #evict() {
    const target = this.maxBytes * 0.9;
    const entries = [];
    for (const file of readdirSync(this.dir)) {
      if (!file.endsWith(".pdf")) continue;
      try {
        const stats = statSync(join(this.dir, file));
        entries.push({ file, size: stats.size, mtime: stats.mtimeMs });
      } catch {
        // gone
      }
    }
    let total = this.#bytes;
    entries.sort((a, b) => a.mtime - b.mtime);
    for (const entry of entries) {
      if (total <= target) break;
      try {
        rmSync(join(this.dir, `${entry.file.slice(0, -4)}.json`), { force: true });
        rmSync(join(this.dir, entry.file), { force: true });
        total -= entry.size;
      } catch {
        // best effort
      }
    }
    this.#bytes = Math.max(0, total);
  }

  get sizeBytes() {
    return this.#bytes;
  }
}

// ── browser pool ────────────────────────────────────────────────────────────

/** Runs inside the editor page — must stay self-contained (Playwright
 *  serializes the source). Waits for a settled canvas, injects the document
 *  through the public host methods, and returns the PDF as base64. */
async function renderInPage({ json, iso, capture }) {
  const host = document.querySelector("docen-document");
  if (!host) throw new Error("docen-document element not found on the page");
  const canvasPages = () => host.shadowRoot?.querySelectorAll(".canvas-pages > *").length ?? 0;
  const settle = async (frames) => {
    let last = -1;
    let stable = 0;
    while (stable < frames) {
      await new Promise((r) => requestAnimationFrame(r));
      const n = canvasPages();
      if (n > 0 && n === last) stable++;
      else stable = 0;
      last = n;
    }
    return last;
  };
  if (!window.__docenPdfSettled) {
    await settle(6);
    window.__docenPdfSettled = true;
  }
  if (capture) return { json: host.getJSON(), canvasPages: canvasPages() };

  const renderStart = performance.now();
  host.setJSON(json);
  host.repaginate();
  const settled = await settle(3);
  const renderMs = performance.now() - renderStart;

  const exportStart = performance.now();
  const result = await host.exportPdf({
    metadata: { creationDate: new Date(iso), modDate: new Date(iso) },
  });
  const exportMs = performance.now() - exportStart;
  if (settled !== result.pages.length) {
    throw new Error(
      `render/canvas mismatch: ${settled} canvas pages vs ${result.pages.length} exported pages`,
    );
  }

  const bytes = result.data;
  let binary = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  }
  return {
    b64: btoa(binary),
    bytes: bytes.byteLength,
    pages: result.pages.length,
    renderMs,
    exportMs,
  };
}

class RenderPool {
  #workers = [];
  #stopped = false;

  constructor({ browser, size }) {
    this.browser = browser;
    this.size = size;
  }

  async start() {
    for (let i = 0; i < this.size; i++) {
      this.#workers.push(await this.#spawn(i));
    }
  }

  async #spawn(index) {
    const context = await this.browser.newContext({ viewport: { width: 1440, height: 900 } });
    const page = await context.newPage();
    page.on("pageerror", (err) => console.error(`pdf-render: [page ${index}] error: ${err}`));
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: timeoutMs });
    await page.waitForSelector("docen-document", { timeout: timeoutMs });
    return { index, context, page };
  }

  async #recycle(worker) {
    try {
      await worker.context.close();
    } catch {
      // already gone
    }
    const fresh = await this.#spawn(worker.index);
    Object.assign(worker, fresh);
  }

  /** One job on a free worker: { json, iso, capture } → renderInPage result. */
  async render(job) {
    const worker = this.#workers.shift();
    if (!worker) throw new Error("render pool exhausted");
    try {
      const call = worker.page.evaluate(renderInPage, job);
      let timer;
      try {
        return await Promise.race([
          call,
          new Promise((_, reject) => {
            timer = setTimeout(
              () =>
                reject(
                  new Error(`render timed out after ${timeoutMs} ms (worker ${worker.index})`),
                ),
              timeoutMs,
            );
          }),
        ]);
      } catch (err) {
        // A timed-out evaluate can't be cancelled — recycle the page so the
        // next job starts from a clean editor (keep the original error).
        try {
          await this.#recycle(worker);
        } catch {
          // browser already gone (shutdown) — the original error is reported
        }
        throw err;
      } finally {
        clearTimeout(timer);
      }
    } finally {
      if (!this.#stopped) this.#workers.push(worker);
    }
  }

  /** N consumers pull from one queue; results land in row order (row.index). */
  async run(rows) {
    const queue = [...rows];
    const results = [];
    const consumer = async () => {
      while (!this.#stopped) {
        const row = queue.shift();
        if (!row) return;
        const started = performance.now();
        try {
          const out = await this.render({ json: row.json, iso: exportDateISO });
          results.push({ row, out, wallMs: performance.now() - started });
        } catch (err) {
          results.push({
            row,
            error: err instanceof Error ? err.message : String(err),
            wallMs: performance.now() - started,
          });
        }
      }
    };
    await Promise.all(this.#workers.map(() => consumer()));
    return results;
  }

  stop() {
    this.#stopped = true;
  }

  async close() {
    const workers = this.#workers;
    this.#workers = [];
    for (const worker of workers) {
      try {
        await worker.context.close();
      } catch {
        // best effort
      }
    }
  }
}

// ── job planning ────────────────────────────────────────────────────────────

const inputs = jsonPaths.map((path) => {
  if (!existsSync(path)) usageError(`no such document JSON: ${path}`);
  try {
    return { name: basename(path, ".json"), path, json: JSON.parse(readFileSync(path, "utf8")) };
  } catch (err) {
    return usageError(`cannot parse ${path}: ${err instanceof Error ? err.message : String(err)}`);
  }
});

const planned = [];
for (let rep = 0; rep < repeat; rep++) {
  const source = inputs.length > 0 ? inputs : [{ name: "demo", path: null, json: null }];
  for (const input of source) {
    planned.push({
      index: planned.length,
      name: repeat > 1 ? `${input.name}-r${rep + 1}` : input.name,
      path: input.path,
      json: input.json,
    });
  }
}
if (planned.length > 1 && outFlag === "-") {
  usageError("--out - cannot receive more than one PDF; use --out <dir> or --out <file>");
}
if (planned.length > 1 && outFlag && outFlag.endsWith(".pdf")) {
  usageError(
    `--out "${outFlag}" is one file but ${planned.length} renders were requested; pass a directory`,
  );
}
if (inputs.length === 0 && repeat > 1) {
  usageError("--repeat needs --json inputs (the demo document renders once)");
}

const outPathFor = (row) => {
  if (outFlag === "-") return "-";
  if (!outFlag) return row.path ? `${row.path.replace(/\.json$/i, "")}.pdf` : "-";
  const isDir = existsSync(outFlag) ? statSync(outFlag).isDirectory() : outFlag.endsWith("/");
  if (isDir || planned.length > 1) return join(outFlag, `${row.name}.pdf`);
  return outFlag;
};

// Cache preflight: content-addressed hits need no browser at all.
const cache = noCache ? null : new PdfCache(cacheDir, cacheMaxBytes);
const rows = [];
for (const row of planned) {
  if (row.json !== null) row.key = PdfCache.keyFor(row.json);
  if (row.key && cache) {
    const hit = cache.get(row.key);
    if (hit) {
      row.cached = hit;
      row.wallMs = hit.ms;
      continue;
    }
  }
  rows.push(row);
}
// Single-flight: identical content in one batch shares one render (only when
// the cache is on — without it, --repeat is a benchmark of N real renders).
if (cache) {
  const byKey = new Map();
  for (const row of rows) {
    if (!row.key) continue;
    const first = byKey.get(row.key);
    if (first) row.sharedWith = first;
    else byKey.set(row.key, row);
  }
}
const preflightHits = planned.length - rows.length;
if (preflightHits > 0) {
  console.error(
    `pdf-render: ${preflightHits}/${planned.length} request(s) served from cache (${cacheDir})`,
  );
}
// The all-cache-hit path needs no browser at all (the caller's document JSON
// is already known); the demo-document capture below needs one live page and
// therefore always launches.

async function emitOutputs(allRows) {
  for (const row of allRows) {
    const payload = row.result ? Buffer.from(row.result.b64, "base64") : row.cached?.bytes;
    if (!payload) continue;
    const out = outPathFor(row);
    if (out === "-") {
      process.stdout.write(payload);
    } else {
      mkdirSync(dirname(resolve(out)), { recursive: true });
      writeFileSync(out, payload);
    }
  }
}

function percentile(sorted, p) {
  if (sorted.length === 0) return 0;
  const index = Math.min(sorted.length - 1, Math.ceil((p / 100) * sorted.length) - 1);
  return sorted[index];
}

function report(allRows, activeCache) {
  let hits = 0;
  let misses = 0;
  let failures = 0;
  let deduped = 0;
  const missWall = [];
  const hitWall = [];
  for (const row of allRows) {
    const bytes = row.result?.bytes ?? row.cached?.bytes.byteLength;
    const status = row.error
      ? "FAIL"
      : row.interrupted
        ? "intr"
        : row.deduped
          ? "dup"
          : row.cached
            ? "hit"
            : row.result
              ? "miss"
              : "skip";
    const wallMs = Math.round(row.wallMs ?? 0);
    if (row.error) failures++;
    else if (row.cached) {
      hits++;
      hitWall.push(wallMs);
    } else if (row.deduped) {
      deduped++;
    } else if (row.result) {
      misses++;
      missWall.push(wallMs);
    }
    const rss = row.rss ?? process.memoryUsage().rss;
    console.error(
      `pdf-render: [${row.index + 1}/${allRows.length}] ${status.padEnd(4)} ${row.name.padEnd(18)}` +
        ` wall=${wallMs}ms` +
        (row.result
          ? ` render=${Math.round(row.result.renderMs)}ms export=${Math.round(row.result.exportMs)}ms`
          : "") +
        ` pages=${row.result?.pages ?? "?"}` +
        ` bytes=${bytes ?? "?"}` +
        ` rss=${mb(rss)}` +
        (row.chromium ? ` chromium=${mb(row.chromium.bytes)}/${row.chromium.processes}p` : "") +
        (row.error ? ` error="${row.error}"` : ""),
    );
  }
  const sortedMiss = [...missWall].sort((a, b) => a - b);
  const sortedHits = [...hitWall].sort((a, b) => a - b);
  console.error(
    `pdf-render: summary renders=${misses} hits=${hits} deduped=${deduped} failures=${failures}` +
      (sortedMiss.length > 0
        ? ` | miss wall min=${Math.round(sortedMiss[0])}ms median=${Math.round(percentile(sortedMiss, 50))}ms p95=${Math.round(percentile(sortedMiss, 95))}ms`
        : "") +
      (sortedHits.length > 0
        ? ` | cache hit median=${Math.round(percentile(sortedHits, 50))}ms`
        : "") +
      ` | node rss=${mb(process.memoryUsage().rss)}` +
      (activeCache ? ` | cache=${cacheDir} ${mb(activeCache.sizeBytes)}` : ""),
  );
}

// ── main ────────────────────────────────────────────────────────────────────

async function main() {
  const token = `--docen-pdf-render=${process.pid}-${Date.now()}`;
  const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || "playwright");
  const browser = await chromium.launch({
    headless: true,
    args: [token, "--disable-dev-shm-usage", ...(noSandbox ? ["--no-sandbox"] : [])],
    ...(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {}),
  });
  console.error(
    `pdf-render: chromium ready (pool=${Math.min(concurrency, rows.length)}, page=${url}, build=${buildId}, date=${exportDateISO})`,
  );

  const pool = new RenderPool({ browser, size: Math.min(concurrency, rows.length) });
  let interrupted = false;
  let started = false;
  let sigints = 0;
  process.on("SIGINT", () => {
    sigints++;
    if (sigints > 1) {
      console.error("pdf-render: second interrupt — exiting");
      process.exit(130);
    }
    interrupted = true;
    if (!started) {
      console.error("pdf-render: interrupt during pool startup — shutting down");
      void browser.close().finally(() => process.exit(130));
      return;
    }
    console.error("pdf-render: interrupt — finishing in-flight renders, then shutting down");
    pool.stop();
    setTimeout(() => {
      console.error("pdf-render: interrupt grace elapsed — shutting down");
      void browser.close().finally(() => process.exit(130));
    }, SIGINT_GRACE_MS).unref();
  });

  try {
    await pool.start();
    started = true;

    // No --json: render the editor's own open document once (captured through
    // getJSON so the cache key is content-based). Needs one live page.
    const deferred = rows.filter((row) => row.json === null);
    for (const row of deferred) {
      const captured = await pool.render({ json: null, iso: exportDateISO, capture: true });
      row.json = captured.json;
      row.key = PdfCache.keyFor(row.json);
      const hit = cache?.get(row.key);
      if (hit) {
        row.cached = hit;
        row.wallMs = hit.ms;
      }
    }
    const runnable = rows.filter((row) => !row.cached && !row.sharedWith);

    const results = await pool.run(runnable);
    for (const { row, out, error, wallMs } of results) {
      row.wallMs = wallMs;
      row.rss = process.memoryUsage().rss;
      row.chromium = chromiumTreeRss(token);
      if (out) {
        row.result = out;
      } else if (interrupted) {
        row.interrupted = true;
      } else {
        row.error = error;
      }
    }
    // Identical requests in this batch share the representative's bytes.
    for (const row of rows) {
      if (!row.sharedWith) continue;
      row.result = row.sharedWith.result;
      row.error = row.sharedWith.error;
      row.rss = row.sharedWith.rss;
      row.chromium = row.sharedWith.chromium;
      row.deduped = true;
      row.wallMs = 0;
    }

    if (cache) {
      for (const row of planned) {
        if (!row.result || row.cached || row.deduped) continue;
        cache.put(row.key, Buffer.from(row.result.b64, "base64"), {
          createdAt: new Date().toISOString(),
          build: buildId,
          date: exportDateISO,
          editorUrl: url,
          pages: row.result.pages,
          wallMs: Math.round(row.wallMs ?? 0),
        });
      }
    }

    await emitOutputs(planned);
    report(planned, cache);
    const failed = planned.filter((row) => row.error);
    if (keepOpen) {
      console.error(
        `pdf-render: --keep-open — browser stays open; Ctrl-C to exit (token ${token})`,
      );
      await new Promise(() => {});
    }
    await browser.close();
    process.exitCode = failed.length > 0 ? 1 : interrupted ? 130 : 0;
  } catch (err) {
    if (!interrupted) {
      console.error(`pdf-render: ${err instanceof Error ? err.message : String(err)}`);
    }
    try {
      await browser.close();
    } catch {
      // already closed
    }
    process.exitCode = interrupted ? 130 : 1;
  }
}

if (rows.length === 0) {
  await emitOutputs(planned);
  report(planned, cache);
} else {
  await main();
}
