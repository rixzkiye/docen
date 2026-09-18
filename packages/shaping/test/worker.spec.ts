import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import {
  ShapingWorkerClient,
  ShapingWorkerHandler,
  initShapingWasm,
  type ShapingWorkerTransport,
  type WorkerRequest,
  type WorkerResponse,
} from "../src/index.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fontsDir = path.join(__dirname, "fixtures/fonts");

describe("R6.8 ShapingWorker Offloading", () => {
  let openSansBytes: Uint8Array;

  beforeAll(async () => {
    await initShapingWasm();
    openSansBytes = fs.readFileSync(path.join(fontsDir, "OpenSans-Regular.ttf"));
  });

  it("handles worker messages via ShapingWorkerHandler", async () => {
    const handler = new ShapingWorkerHandler();
    const responses: WorkerResponse[] = [];
    const postResponse = (res: WorkerResponse) => responses.push(res);

    // 1. init
    await handler.handleMessage({ type: "init" }, postResponse);
    expect(responses[0]).toEqual({ type: "initDone", success: true });

    // 2. registerFont
    await handler.handleMessage(
      { type: "registerFont", fontId: 10, fontData: openSansBytes },
      postResponse,
    );
    expect(responses[1]).toEqual({ type: "fontRegistered", fontId: 10 });

    // 3. shapeBatch
    await handler.handleMessage(
      {
        type: "shapeBatch",
        id: 42,
        items: [
          { fontId: 10, text: "Hello World" },
          { fontId: 10, text: "OpenType Shaping" },
        ],
      },
      postResponse,
    );

    const batchRes = responses[2];
    expect(batchRes.type).toBe("batchResult");
    if (batchRes.type === "batchResult") {
      expect(batchRes.id).toBe(42);
      expect(batchRes.results.length).toBe(2);
      expect(batchRes.results[0].glyphs.length).toBeGreaterThan(0);
      expect(batchRes.results[0].totalAdvance).toBeGreaterThan(0);
      expect(batchRes.results[1].glyphs.length).toBeGreaterThan(0);
    }

    // 4. dropFont
    await handler.handleMessage({ type: "dropFont", fontId: 10 }, postResponse);
    expect(responses[3]).toEqual({ type: "fontDropped", fontId: 10 });
  });

  it("provides client interface via ShapingWorkerClient", async () => {
    const client = new ShapingWorkerClient();
    await client.init();
    expect(client.offloaded).toBe(false);

    const fontId = await client.registerFont(openSansBytes);
    expect(fontId).toBeGreaterThan(0);

    const result = await client.shape(fontId, "Async text");
    expect(result.glyphs.length).toBeGreaterThan(0);
    expect(result.totalAdvance).toBeGreaterThan(0);

    const batch = await client.shapeBatch([
      { fontId, text: "Paragraph 1" },
      { fontId, text: "Paragraph 2" },
      { fontId, text: "Paragraph 3" },
    ]);
    expect(batch.length).toBe(3);
    for (const r of batch) {
      expect(r.glyphs.length).toBeGreaterThan(0);
      expect(r.totalAdvance).toBeGreaterThan(0);
    }

    await client.dropFont(fontId);
    client.dispose();
  });

  it("offloads through a worker transport with correlated responses and transferred fonts", async () => {
    const received: WorkerRequest[] = [];
    const listeners: ((event: { data?: unknown }) => void)[] = [];
    const handler = new ShapingWorkerHandler();
    let terminated = false;
    const worker: ShapingWorkerTransport = {
      postMessage(message) {
        received.push(message as WorkerRequest);
        queueMicrotask(() => {
          void handler.handleMessage(message as WorkerRequest, (response) => {
            for (const listener of listeners) listener({ data: response });
          });
        });
      },
      addEventListener(type, listener) {
        if (type === "message") listeners.push(listener);
      },
      terminate() {
        terminated = true;
      },
    };

    const client = new ShapingWorkerClient(worker);
    expect(client.offloaded).toBe(true);
    await client.init();
    const fontId = await client.registerFont(openSansBytes);
    const result = await client.shape(fontId, "Off-thread shaping");
    expect(result.glyphs.length).toBeGreaterThan(0);
    const batch = await client.shapeBatch([
      { fontId, text: "One" },
      { fontId, text: "Two" },
    ]);
    expect(batch).toHaveLength(2);
    expect(received.map((m) => m.type)).toEqual([
      "init",
      "registerFont",
      "shapeBatch",
      "shapeBatch",
    ]);
    // Fonts are transferred as copies so the caller's bytes stay usable.
    expect(openSansBytes.byteLength).toBeGreaterThan(0);
    await client.dropFont(fontId);
    client.dispose();
    expect(terminated).toBe(true);
  });

  it("falls back to in-thread shaping when the worker cannot initialize", async () => {
    let terminated = false;
    const broken: ShapingWorkerTransport = {
      postMessage() {
        throw new Error("worker blocked by CSP");
      },
      addEventListener() {},
      terminate() {
        terminated = true;
      },
    };
    const client = new ShapingWorkerClient(broken);
    await client.init();
    expect(client.offloaded).toBe(false);
    expect(terminated).toBe(true);
    const fontId = await client.registerFont(openSansBytes);
    const result = await client.shape(fontId, "Fallback");
    expect(result.glyphs.length).toBeGreaterThan(0);
    client.dispose();
  });
});
