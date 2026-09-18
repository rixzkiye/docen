import { createHash } from "node:crypto";

import { unzipSync } from "@office-open/core";
import type { JSONContent } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import { DOCX_EPOCH, generateDOCX, generateDOCXStream, generateDOCXSync } from "../index";

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const sha = (bytes: Uint8Array): string => createHash("sha256").update(bytes).digest("hex");
const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

/** The same model in every process (see tests/determinism-child.mjs). */
function model(): JSONContent {
  return {
    type: "doc",
    attrs: { core: { title: "Reproducible", creator: "docen" } },
    content: [
      {
        type: "paragraph",
        content: [
          { type: "text", text: "deterministic output " },
          { type: "image", attrs: { src: PNG, width: 1, height: 1 } },
          { type: "image", attrs: { src: PNG, width: 2, height: 2 } },
        ],
      },
      { type: "paragraph", content: [{ type: "text", text: "second paragraph" }] },
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

describe("deterministic generation", () => {
  it("generates byte-identical DOCX twice in one process (async, sync, stream)", async () => {
    const async1 = (await generateDOCX(model())) as Uint8Array;
    const async2 = (await generateDOCX(model())) as Uint8Array;
    const sync1 = generateDOCXSync(model()) as Uint8Array;
    const sync2 = generateDOCXSync(model()) as Uint8Array;
    const streamed = await collect(await generateDOCXStream(model()));

    expect(sha(async2)).toBe(sha(async1));
    expect(sha(sync2)).toBe(sha(sync1));
    expect(sha(sync1)).toBe(sha(async1));
    expect(sha(streamed)).toBe(sha(async1));
  });

  it("uses the fixed epoch for missing core-property dates", async () => {
    const bytes = (await generateDOCX(model())) as Uint8Array;
    const core = decode(unzipSync(bytes)["docProps/core.xml"]!);
    expect(core).toContain(`<dcterms:created xsi:type="dcterms:W3CDTF">${DOCX_EPOCH}`);
    expect(core).toContain(`<dcterms:modified xsi:type="dcterms:W3CDTF">${DOCX_EPOCH}`);
  });

  it("numbers drawings from one per generation (no cross-generation counter drift)", async () => {
    const first = decode(
      unzipSync((await generateDOCX(model())) as Uint8Array)["word/document.xml"]!,
    );
    const second = decode(
      unzipSync((await generateDOCX(model())) as Uint8Array)["word/document.xml"]!,
    );
    const ids = (xml: string) => [...xml.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1]);
    expect(ids(first)).toEqual(["1", "2"]);
    expect(ids(second)).toEqual(["1", "2"]);
  });

  it("honours an explicit date and omits dates for date: null", async () => {
    const fixed = "2024-05-06T07:08:09.000Z";
    const dated = (await generateDOCX(model(), { date: fixed })) as Uint8Array;
    const datedCore = decode(unzipSync(dated)["docProps/core.xml"]!);
    expect(datedCore).toContain(`<dcterms:created xsi:type="dcterms:W3CDTF">${fixed}`);
    expect(datedCore).toContain(`<dcterms:modified xsi:type="dcterms:W3CDTF">${fixed}`);

    const undated = (await generateDOCX(model(), { date: null })) as Uint8Array;
    const undatedCore = decode(unzipSync(undated)["docProps/core.xml"]!);
    expect(undatedCore).not.toContain("<dcterms:created");
    expect(undatedCore).not.toContain("<dcterms:modified");

    // A source-carried date wins over the generation clock.
    const fromSource = (await generateDOCX(model(), {
      document: { created: "2001-02-03T04:05:06.000Z" },
    })) as Uint8Array;
    const sourceCore = decode(unzipSync(fromSource)["docProps/core.xml"]!);
    expect(sourceCore).toContain("2001-02-03T04:05:06.000Z");
  });

  it("never mutates the input JSON (default prepare works on a copy)", async () => {
    const json = model();
    const snapshot = structuredClone(json);
    await generateDOCX(json);
    generateDOCXSync(json);
    expect(json).toEqual(snapshot);
  });
});
