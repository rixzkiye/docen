# @docen/docx benchmarks

## Streaming benchmark (`streaming-bench.mjs`)

Measures wall time, output bytes and RSS for a 100–300 page document across all
three generation paths. Each phase runs in its own child process so the
reported RSS is that phase's own footprint.

```bash
pnpm --filter @docen/docx build
node --expose-gc packages/docx/bench/streaming-bench.mjs --pages 100
node --expose-gc packages/docx/bench/streaming-bench.mjs --pages 300 --json
```

The model is generated in code (no I/O): `--pages N` emits `N × 15`
one-sentence paragraphs. The calibration is verified with LibreOffice PDF
conversion — 1 500 paragraphs → 100 pages, 4 500 → 300 pages (Letter, default
styles).

### Recorded run

`streaming-results.json` (2026-09-18, Linux x86_64, Node v24.14.0):

| Pages | Path   | Time   | Output  | RSS delta | Peak RSS |
| ----- | ------ | ------ | ------- | --------- | -------- |
| 100   | sync   | 154 ms | 14.3 KB | 2.4 MiB   | 2.3 MiB  |
| 100   | async  | 318 ms | 14.3 KB | 8.8 MiB   | 8.5 MiB  |
| 100   | stream | 95 ms  | 14.3 KB | 8.4 MiB   | 8.4 MiB  |
| 300   | sync   | 70 ms  | 24.2 KB | 10.3 MiB  | 10.0 MiB |
| 300   | async  | 180 ms | 24.2 KB | 28.2 MiB  | 28.1 MiB |
| 300   | stream | 218 ms | 24.2 KB | 27.3 MiB  | 27.3 MiB |

### Thresholds

- **Byte parity** — `stream` and `sync` outputs hash equal for the same model
  (asserted in `src/converters/streaming.spec.ts`).
- **Time** — each path stays under 2 s for 300 pages (recorded worst: 218 ms).
- **Peak RSS** — each path stays under 64 MiB delta for 300 pages (recorded
  worst: 28 MiB).
- **Streaming retention** — the stream consumer releases each chunk; its peak
  RSS stays within 16 MiB of the async path (recorded: stream 27.3 vs async
  28.1 MiB at 300 pages), and the streaming test asserts the output arrives in
  more than one chunk.

### Measuring note (Node)

`generateDOCXStream` delivers the archive in 64 KiB chunks, but on Node the
native ZIP writer compresses the whole package before the first chunk is
enqueued, and `prepareDocument` clones the model — so streaming is not
_less_ memory-hungry than `generateDOCX` here; the win is incremental delivery
and constant consumer-side retention (verified by the peak-RSS sampling). The
compressed output of text documents is small (24 KB for 300 pages) because XML
text deflates ~30×, so the peaks are dominated by the compiled parts, not the
archive. The per-path bound above is the contract the benchmark gates.

> The time figures come from single runs and are noisy (a cold first run was
> ~2× slower); rerun before comparing machines. `--json` emits the raw numbers
> for recording.
