import { createHash } from "node:crypto";

import type { JSONContent } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import { generateDOCXStream, generateDOCXSync } from "../index";

/** ~15 paragraphs per page (LibreOffice-verified; see bench/README.md).
 *  High-entropy tokens keep the archive above the 64 KiB stream chunk size. */
const PARAGRAPHS_PER_PAGE = 15;

function largeDocument(pages: number): JSONContent {
  const content: JSONContent[] = [];
  for (let paragraph = 0; paragraph < pages * PARAGRAPHS_PER_PAGE; paragraph++) {
    const tokens: string[] = [];
    for (let token = 0; token < 12; token++) {
      tokens.push(((paragraph * 2654435761 + token * 40503) >>> 0).toString(36));
    }
    content.push({
      type: "paragraph",
      content: [{ type: "text", text: `${paragraph} ${tokens.join(" ")}` }],
    });
  }
  return { type: "doc", content };
}

const sha = (bytes: Uint8Array): string => createHash("sha256").update(bytes).digest("hex");

describe("generateDOCXStream", () => {
  it("delivers a 100-page document as chunked bytes identical to sync", async () => {
    const model = largeDocument(100);
    const sync = generateDOCXSync(model) as Uint8Array;
    const stream = await generateDOCXStream(model);
    const reader = stream.getReader();
    const chunks: Uint8Array[] = [];
    let total = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      total += value.byteLength;
    }
    // Streaming must be incremental, not one giant chunk.
    expect(chunks.length).toBeGreaterThan(1);
    const streamed = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
      streamed.set(chunk, offset);
      offset += chunk.byteLength;
    }
    expect(sha(streamed)).toBe(sha(sync));
  }, 30_000);
});
