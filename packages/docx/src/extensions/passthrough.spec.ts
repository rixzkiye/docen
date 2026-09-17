import { unzipSync } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCXSync, resolveDocument } from "../index";
import { decodePassthroughData, encodePassthroughData } from "./passthrough";

/**
 * Passthrough data codec + OLE (w:object) round-trip. A passthrough branch may
 * carry binary leaves (the OLE embed payload, a VML picture, …); plain
 * JSON.stringify corrupts them into `{"0": …}` index objects and compile then
 * crashed in office-open's media reader ("Unsupported data type: object").
 * The codec tags each binary leaf as base64 so the JSON→JSON persistence
 * round-trip (autosave, v-model) keeps the exact bytes.
 */

const decoder = new TextDecoder();

/** A fake OLE compound file: CFB magic + deterministic payload. */
function oleBytes(): Uint8Array {
  const bytes = new Uint8Array(Array.from({ length: 256 }, (_, i) => (i * 7 + 3) % 256));
  bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
  return bytes;
}

function oleDoc(): DocumentOptions {
  return {
    sections: [
      {
        children: [
          { paragraph: { text: "before ole" } },
          {
            paragraph: {
              children: [
                {
                  object: {
                    width: 120,
                    height: 80,
                    embed: {
                      data: oleBytes(),
                      progId: "Package",
                      fileName: "oleObject1.bin",
                      relationshipType: "oleObject",
                    },
                  },
                },
              ],
            },
          },
          { paragraph: { text: "after ole" } },
        ],
      },
    ],
  };
}

describe("passthrough binary codec", () => {
  it("restores Uint8Array, ArrayBuffer and DataView leaves as bytes", () => {
    const bytes = oleBytes();
    const encoded = encodePassthroughData({
      embed: { data: bytes },
      buffer: bytes.buffer.slice(0),
      view: new DataView(bytes.buffer, 8, 4),
    });
    // JSON-safe on the wire: no index-object corruption.
    expect(encoded).toContain("$docenBinary");
    const decoded = decodePassthroughData<{
      embed: { data: Uint8Array };
      buffer: Uint8Array;
      view: Uint8Array;
    }>(encoded);
    // Decoding normalizes every binary leaf to bytes (Buffer under Node,
    // Uint8Array in browsers) — compare contents, not the view class.
    expect(new Uint8Array(decoded.embed.data)).toEqual(bytes);
    expect(new Uint8Array(decoded.buffer)).toEqual(bytes);
    expect(new Uint8Array(decoded.view)).toEqual(bytes.subarray(8, 12));
  });

  it("keeps plain JSON byte-identical in both directions", () => {
    const value = { complexField: { instruction: 'PAGEREF "目标书签" \\h', result: "1" } };
    expect(encodePassthroughData(value)).toBe(JSON.stringify(value));
    expect(decodePassthroughData(encodePassthroughData(value))).toEqual(value);
  });
});

describe("OLE w:object round-trip", () => {
  it("exports w:object, the embedded part and its relationship", () => {
    const bytes = generateDOCXSync(resolveDocument(oleDoc()), { prepare: false }) as Uint8Array;
    const zip = unzipSync(bytes);
    expect(zip["word/embeddings/oleObject1.bin"]).toEqual(oleBytes());
    const documentXml = decoder.decode(zip["word/document.xml"]);
    expect(documentXml).toContain("<w:object>");
    expect(documentXml).toContain("<o:OLEObject");
    const rels = decoder.decode(zip["word/_rels/document.xml.rels"]);
    expect(rels).toContain('Target="embeddings/oleObject1.bin"');
  });

  it("survives a JSON persistence round-trip and re-export (no crash)", () => {
    // The autosave shape: resolve → JSON.stringify → JSON.parse → export.
    const persisted = JSON.parse(JSON.stringify(resolveDocument(oleDoc())));
    const gen1 = generateDOCXSync(persisted, { prepare: false }) as Uint8Array;
    const reparsed = parseDOCXSync(gen1);
    // The OLE branch came back as an inline passthrough atom with live bytes.
    const atom = reparsed.content?.[1]?.content?.find((c) => c.type === "inlinePassthrough");
    expect(atom).toBeDefined();
    const gen2 = generateDOCXSync(reparsed, { prepare: false }) as Uint8Array;
    const zip = unzipSync(gen2);
    expect(zip["word/embeddings/oleObject1.bin"]).toEqual(oleBytes());
    expect(decoder.decode(zip["word/document.xml"])).toContain("<w:object>");
  });
});
