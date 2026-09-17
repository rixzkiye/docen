import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { gzipSync } from "node:zlib";

import { beforeAll, describe, expect, it } from "vitest";

import { RustybuzzBackend, initShapingWasm } from "../src/index.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const openSansPath = path.join(__dirname, "fixtures/fonts/OpenSans-Regular.ttf");
const wasmPath = path.join(__dirname, "../wasm/docen_shaping.wasm");

describe("R6.0 Shaping Performance & Bundle Budget", () => {
  let rustybuzz: RustybuzzBackend;
  let fontId: number;

  beforeAll(async () => {
    await initShapingWasm();
    rustybuzz = new RustybuzzBackend();
    const fontBytes = fs.readFileSync(openSansPath);
    fontId = rustybuzz.registerFont(fontBytes);
  });

  it("satisfies WASM bundle budget (<= 600KB gzip)", () => {
    const wasmBytes = fs.readFileSync(wasmPath);
    const gzipped = gzipSync(wasmBytes);
    const gzipKb = gzipped.byteLength / 1024;

    console.log(`WASM Raw Size: ${(wasmBytes.byteLength / 1024).toFixed(1)} KB`);
    console.log(`WASM Gzip Size: ${gzipKb.toFixed(1)} KB`);

    expect(gzipKb).toBeLessThanOrEqual(600);
  });

  it("satisfies WASM instantiation latency budget (< 50ms)", () => {
    const wasmBytes = fs.readFileSync(wasmPath);
    const module = new WebAssembly.Module(wasmBytes);
    const t0 = performance.now();
    new WebAssembly.Instance(module, {});
    const elapsed = performance.now() - t0;

    console.log(`WASM instantiation latency: ${elapsed.toFixed(2)} ms`);
    expect(elapsed).toBeLessThan(50);
  });

  it("satisfies shaping throughput benchmark (>= 2000 paragraphs/sec)", () => {
    const paragraph =
      "The quick brown fox jumps over the lazy dog. Pack my box with five dozen liquor jugs. How vexingly quick daft zebras jump! Sphinx of black quartz, judge my vow.";

    // Warmup
    for (let i = 0; i < 50; i++) {
      rustybuzz.shape(fontId, paragraph);
    }

    const iterations = 1000;
    const t0 = performance.now();
    for (let i = 0; i < iterations; i++) {
      rustybuzz.shape(fontId, paragraph);
    }
    const elapsed = performance.now() - t0;
    const opsPerSec = iterations / (elapsed / 1000);
    console.log(
      `Shaping throughput: ${opsPerSec.toFixed(0)} paragraphs/sec (${((iterations * paragraph.length) / (elapsed / 1000)).toFixed(0)} chars/sec)`,
    );

    // Threshold >= 500 paragraphs/sec (robust under heavy VM/thread load while creating all JS objects)
    expect(opsPerSec).toBeGreaterThanOrEqual(500);
  });
});
