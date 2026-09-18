import { createHash } from "node:crypto";

import { unzipSync } from "@office-open/core";
import type { JSONContent } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import { decodedSrcCacheSize, mediaOfSrc } from "../extensions/image";
import { generateDOCX, generateDOCXStream, generateDOCXSync, parseDOCXSync } from "../index";

const sha = (bytes: Uint8Array): string => createHash("sha256").update(bytes).digest("hex");

/** A distinct (content-wise) image data URL per seed. */
const pngSrc = (seed: number): string =>
  `data:image/png;base64,${Buffer.from(`docen-image-${seed}-${(seed * 2654435761) % 2 ** 31}`).toString("base64")}`;

function document(seed: number, images = 2): JSONContent {
  return {
    type: "doc",
    content: [
      {
        type: "paragraph",
        content: [
          { type: "text", text: `unique-body-${seed}` },
          ...Array.from({ length: images }, (_, index) => ({
            type: "image",
            attrs: { src: pngSrc(seed * 100 + index), width: 1, height: 1 },
          })),
        ],
      },
    ],
  };
}

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.byteLength;
  }
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}

const docPrIds = (bytes: Uint8Array): string[] => {
  const xml = new TextDecoder().decode(unzipSync(bytes)["word/document.xml"]!);
  return [...xml.matchAll(/<wp:docPr id="(\d+)"/g)].map((match) => match[1]!);
};

describe("parallel generation with the shared defaultManager", () => {
  it("keeps every document's content and id sequence isolated", async () => {
    const docs = Array.from({ length: 8 }, (_, seed) => document(seed));
    const outputs = await Promise.all(docs.map((doc) => generateDOCX(doc)));

    outputs.forEach((output: unknown, seed: number) => {
      const bytes = output as Uint8Array;
      const text = JSON.stringify(parseDOCXSync(bytes));
      expect(text).toContain(`unique-body-${seed}`);
      for (let other = 0; other < docs.length; other++) {
        if (other !== seed) expect(text).not.toContain(`unique-body-${other}`);
      }
      // A fresh per-generation scope restarts the drawing id sequence.
      expect(docPrIds(bytes)).toEqual(["1", "2"]);
    });
  });

  it("matches sequential bytes for a mixed async/stream batch", async () => {
    const docs = Array.from({ length: 6 }, (_, seed) => document(seed));
    const concurrent = await Promise.all(
      docs.map((doc, index) => (index % 2 === 0 ? generateDOCX(doc) : generateDOCXStream(doc))),
    );
    const concurrentBytes = await Promise.all(
      concurrent.map((value: Uint8Array | ReadableStream<Uint8Array>) =>
        value instanceof ReadableStream ? collect(value) : Promise.resolve(value as Uint8Array),
      ),
    );

    for (const [index, doc] of docs.entries()) {
      const sequential = generateDOCXSync(doc) as Uint8Array;
      expect(sha(concurrentBytes[index]!)).toBe(sha(sequential));
    }
  });

  it("bounds the shared decoded-src cache at 64 entries", async () => {
    for (let seed = 0; seed < 80; seed++) {
      const media = mediaOfSrc(pngSrc(seed));
      expect(media?.bytes.byteLength).toBeGreaterThan(0);
    }
    expect(decodedSrcCacheSize()).toBeLessThanOrEqual(64);
    expect(decodedSrcCacheSize()).toBeGreaterThan(0);

    // Parallel exports also share the cache without exceeding the FIFO cap.
    await Promise.all(Array.from({ length: 8 }, (_, seed) => generateDOCX(document(seed, 10))));
    expect(decodedSrcCacheSize()).toBeLessThanOrEqual(64);
  });
});
