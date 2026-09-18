#!/usr/bin/env node
// Streaming benchmark for @docen/docx — time and peak RSS for a 100–300 page
// document across the three generation paths: `generateDOCXSync`,
// `generateDOCX` (async) and `generateDOCXStream`.
//
// Reproduce:
//   pnpm --filter @docen/docx build
//   node --expose-gc packages/docx/bench/streaming-bench.mjs --pages 100
//   node --expose-gc packages/docx/bench/streaming-bench.mjs --pages 300 --json
//
// Each phase runs in its own child process (re-exec with `--phase`), so the
// reported RSS is that phase's own footprint, not the max over all phases.
// `--expose-gc` is required: the pre-phase collection keeps the baseline
// comparable. `rssDelta` is the RSS retained at the end of the phase (the
// output is still live for sync/async, dropped per chunk for stream);
// `peakRss` samples RSS after every await for the async/stream phases and
// equals the end-of-phase RSS for the blocking sync phase.

import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";

const args = process.argv.slice(2);
const option = (name, fallback) => {
  const index = args.indexOf(name);
  return index >= 0 && args[index + 1] !== undefined ? args[index + 1] : fallback;
};
const phase = option("--phase", "all");
const pages = Number.parseInt(option("--pages", "100"), 10);
const asJson = args.includes("--json");

/**
 * One sentence paragraph measures ~3 lines at the default 11 pt Letter page
 * (LibreOffice-verified: 2 300 paragraphs → 154 pages), so 15 paragraphs
 * ≈ one page. `bench/README.md` records the PDF page count per size.
 */
const PARAGRAPHS_PER_PAGE = 15;

function buildModel(pageCount) {
  const content = [];
  for (let paragraph = 0; paragraph < pageCount * PARAGRAPHS_PER_PAGE; paragraph++) {
    content.push({
      type: "paragraph",
      attrs: paragraph % 40 === 0 ? { alignment: "center" } : {},
      content: [
        {
          type: "text",
          text: `Paragraph ${paragraph} — the streaming benchmark body exercises compile, XML stringification and ZIP compression with a realistic amount of inline text per line.`,
        },
      ],
    });
  }
  return { type: "doc", content };
}

async function runPhase(mode) {
  const { generateDOCX, generateDOCXStream, generateDOCXSync } = await import("@docen/docx");
  const model = buildModel(pages);
  if (globalThis.gc) globalThis.gc();
  const baselineRss = process.memoryUsage().rss;
  let peakRss = baselineRss;
  const sample = () => {
    peakRss = Math.max(peakRss, process.memoryUsage().rss);
  };
  const start = process.hrtime.bigint();

  let outputBytes = 0;
  let hash;
  if (mode === "sync") {
    const bytes = generateDOCXSync(model);
    sample();
    outputBytes = bytes.byteLength;
    hash = createHash("sha256").update(bytes).digest("hex");
  } else if (mode === "async") {
    const bytes = await generateDOCX(model);
    sample();
    outputBytes = bytes.byteLength;
    hash = createHash("sha256").update(bytes).digest("hex");
  } else {
    const stream = await generateDOCXStream(model);
    const reader = stream.getReader();
    const digest = createHash("sha256");
    for (;;) {
      const { done, value } = await reader.read();
      sample();
      if (done) break;
      outputBytes += value.byteLength;
      digest.update(value);
    }
    hash = digest.digest("hex");
  }

  const elapsedMs = Number(process.hrtime.bigint() - start) / 1e6;
  return {
    phase: mode,
    pages,
    elapsedMs,
    outputBytes,
    hash,
    rssDelta: process.memoryUsage().rss - baselineRss,
    peakRss: peakRss - baselineRss,
  };
}

if (phase === "all") {
  const self = fileURLToPath(import.meta.url);
  const runChild = (mode) =>
    JSON.parse(
      execFileSync(
        process.execPath,
        ["--expose-gc", self, "--phase", mode, "--pages", String(pages), "--json"],
        { encoding: "utf8", maxBuffer: 16 * 1024 * 1024 },
      ),
    );
  const results = {
    node: process.version,
    pages,
    sync: runChild("sync"),
    async: runChild("async"),
    stream: runChild("stream"),
  };
  results.streamMatchesSync = results.sync.hash === results.stream.hash;
  if (asJson) console.log(JSON.stringify(results, null, 2));
  else {
    const mib = (value) => `${(value / 1024 / 1024).toFixed(1)} MiB`;
    console.log(
      `node ${results.node}, ${pages} pages (~${pages * PARAGRAPHS_PER_PAGE} paragraphs)`,
    );
    for (const phaseName of ["sync", "async", "stream"]) {
      const result = results[phaseName];
      console.log(
        `  ${phaseName.padEnd(6)} ${result.elapsedMs.toFixed(0).padStart(5)} ms   output ${mib(result.outputBytes)}   RSS delta ${mib(result.rssDelta)}   peak ${mib(result.peakRss)}`,
      );
    }
    console.log(`  bytes identical: ${results.streamMatchesSync}`);
  }
} else {
  console.log(JSON.stringify(await runPhase(phase), null, 2));
}
