import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import {
  ShapingWorkerClient,
  ShapingWorkerHandler,
  initShapingWasm,
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
});
