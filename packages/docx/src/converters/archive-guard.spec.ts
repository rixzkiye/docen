import { crc32 } from "node:zlib";

import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCXSync } from "../index";
import { ARCHIVE_LIMITS, ArchiveRejection, assertArchiveWithinLimits } from "./archive-guard";

const encoder = new TextEncoder();
const bytes = (value: string | Uint8Array | number): Uint8Array =>
  typeof value === "number"
    ? new Uint8Array(value)
    : typeof value === "string"
      ? encoder.encode(value)
      : value;

function zip(files: Record<string, Uint8Array | string | number>): Uint8Array {
  const data: Record<string, Uint8Array> = {};
  for (const [name, value] of Object.entries(files)) data[name] = bytes(value);
  return zipSync(data, { level: 6 });
}

/**
 * A minimal stored-entry ZIP whose single entry streams through a data
 * descriptor (general-purpose bit 3): local header CRC/sizes zeroed, the true
 * values after the payload, no ZIP64. Mirrors what LibreOffice's DOCX export
 * produces (`_rels/.rels`, `word/document.xml`, … all carry bit 3).
 */
function zipWithDataDescriptor(name: string, content: Uint8Array): Uint8Array {
  const nameBytes = encoder.encode(name);
  const checksum = crc32(content) >>> 0;
  const local = new Uint8Array(30 + nameBytes.length);
  const lv = new DataView(local.buffer);
  lv.setUint32(0, 0x04034b50, true);
  lv.setUint16(4, 20, true); // version needed
  lv.setUint16(6, 0x08, true); // flags: data descriptor
  lv.setUint16(8, 0, true); // stored
  lv.setUint32(14, 0, true); // CRC deferred
  lv.setUint32(18, 0, true); // compressed size deferred
  lv.setUint32(22, 0, true); // uncompressed size deferred
  lv.setUint16(26, nameBytes.length, true);
  local.set(nameBytes, 30);
  const descriptor = new Uint8Array(16);
  const dv = new DataView(descriptor.buffer);
  dv.setUint32(0, 0x08074b50, true);
  dv.setUint32(4, checksum, true);
  dv.setUint32(8, content.length, true);
  dv.setUint32(12, content.length, true);
  const central = new Uint8Array(46 + nameBytes.length);
  const cv = new DataView(central.buffer);
  cv.setUint32(0, 0x02014b50, true);
  cv.setUint16(4, 20, true); // version made by
  cv.setUint16(6, 20, true); // version needed
  cv.setUint16(8, 0x08, true); // flags
  cv.setUint16(10, 0, true); // stored
  cv.setUint32(16, checksum, true);
  cv.setUint32(20, content.length, true);
  cv.setUint32(24, content.length, true);
  cv.setUint16(28, nameBytes.length, true);
  cv.setUint32(42, 0, true); // local header offset
  central.set(nameBytes, 46);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, 1, true); // entries on disk
  ev.setUint16(10, 1, true); // entries total
  ev.setUint32(12, central.length, true);
  ev.setUint32(16, local.length + content.length + descriptor.length, true);
  const archive = new Uint8Array(
    local.length + content.length + descriptor.length + central.length + 22,
  );
  let cursor = 0;
  for (const part of [local, content, descriptor, central, eocd]) {
    archive.set(part, cursor);
    cursor += part.length;
  }
  return archive;
}

interface EntryHeader {
  name: string;
  flags: number;
  method: number;
  crc: number;
  compressed: number;
  uncompressed: number;
  localOffset: number;
  centralOffset: number;
}

function entriesOf(input: Uint8Array): EntryHeader[] {
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
  let eocd = input.length - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
  const count = view.getUint16(eocd + 10, true);
  let cursor = view.getUint32(eocd + 16, true);
  const entries: EntryHeader[] = [];
  for (let index = 0; index < count; index++) {
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    entries.push({
      name: new TextDecoder().decode(input.subarray(cursor + 46, cursor + 46 + nameLength)),
      flags: view.getUint16(cursor + 8, true),
      method: view.getUint16(cursor + 10, true),
      crc: view.getUint32(cursor + 16, true),
      compressed: view.getUint32(cursor + 20, true),
      uncompressed: view.getUint32(cursor + 24, true),
      localOffset: view.getUint32(cursor + 42, true),
      centralOffset: cursor,
    });
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

/** Patch one entry's fields in BOTH the central directory and its local header. */
function patchEntry(
  input: Uint8Array,
  name: string,
  patch: Partial<Pick<EntryHeader, "flags" | "method" | "crc" | "uncompressed">>,
): Uint8Array {
  const copy = input.slice();
  const view = new DataView(copy.buffer, copy.byteOffset, copy.byteLength);
  const entry = entriesOf(copy).find((candidate) => candidate.name === name);
  if (!entry) throw new Error(`entry not found: ${name}`);
  const set = (offset: number, value: number, size: 2 | 4) =>
    size === 2 ? view.setUint16(offset, value, true) : view.setUint32(offset, value, true);
  if (patch.flags !== undefined) {
    set(entry.centralOffset + 8, patch.flags, 2);
    set(entry.localOffset + 6, patch.flags, 2);
  }
  if (patch.method !== undefined) {
    set(entry.centralOffset + 10, patch.method, 2);
    set(entry.localOffset + 8, patch.method, 2);
  }
  if (patch.crc !== undefined) {
    set(entry.centralOffset + 16, patch.crc, 4);
    set(entry.localOffset + 14, patch.crc, 4);
  }
  if (patch.uncompressed !== undefined) {
    set(entry.centralOffset + 24, patch.uncompressed, 4);
    set(entry.localOffset + 22, patch.uncompressed, 4);
  }
  return copy;
}

function patchEocd(input: Uint8Array, patch: (view: DataView, eocd: number) => void): Uint8Array {
  const copy = input.slice();
  const view = new DataView(copy.buffer, copy.byteOffset, copy.byteLength);
  let eocd = copy.length - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
  patch(view, eocd);
  return copy;
}

function expectRejection(run: () => void, code: string): void {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(ArchiveRejection);
    expect((error as ArchiveRejection).code).toBe(code);
    return;
  }
  throw new Error(`expected rejection ${code}`);
}

/** A valid DOCX generated by the engine (guard must accept it). */
function validDocx(): Uint8Array {
  return generateDOCXSync({
    type: "doc",
    content: [
      { type: "paragraph", content: [{ type: "text", text: "admission" }] },
      { type: "paragraph", content: [{ type: "text", text: "control" }] },
    ],
  }) as Uint8Array;
}

describe("archive admission limits", () => {
  it("mirrors the Akademi importer's published limits", () => {
    expect(ARCHIVE_LIMITS).toEqual({
      archiveBytes: 32 * 1024 * 1024,
      entries: 512,
      filenameBytes: 255,
      entryUncompressedBytes: 16 * 1024 * 1024,
      aggregateUncompressedBytes: 64 * 1024 * 1024,
      compressionRatio: 100,
      xmlBytes: 4 * 1024 * 1024,
      xmlAggregateBytes: 16 * 1024 * 1024,
      xmlNodes: 100_000,
      xmlAttributes: 200_000,
      xmlAttributeBytes: 8 * 1024 * 1024,
      xmlDepth: 64,
      xmlTextBytes: 8 * 1024 * 1024,
      relationships: 1_024,
      externalTargetBytes: 2_048,
      mediaEntries: 256,
      mediaEntryBytes: 8 * 1024 * 1024,
      mediaAggregateBytes: 32 * 1024 * 1024,
    });
  });

  it("accepts a valid engine-generated DOCX", () => {
    expect(() => assertArchiveWithinLimits(validDocx())).not.toThrow();
  });

  it("rejects non-ZIP input and truncated archives", () => {
    expectRejection(
      () => assertArchiveWithinLimits(encoder.encode("not a zip")),
      "ZIP_MAGIC_INVALID",
    );
    const truncated = validDocx().slice(0, 40);
    expectRejection(() => assertArchiveWithinLimits(truncated), "ZIP_CENTRAL_DIRECTORY_INVALID");
  });

  it("rejects too many entries", () => {
    const files: Record<string, Uint8Array> = {};
    for (let index = 0; index < 600; index++) files[`part-${index}.txt`] = bytes(`x${index}`);
    expectRejection(() => assertArchiveWithinLimits(zip(files)), "ZIP_ENTRY_LIMIT_EXCEEDED");
  });

  it("rejects traversal paths and duplicates", () => {
    expectRejection(
      () => assertArchiveWithinLimits(zip({ "../evil.xml": "<a/>" })),
      "ZIP_FILENAME_INVALID",
    );
    expectRejection(
      () => assertArchiveWithinLimits(zip({ "Word/Doc.xml": "<a/>", "word/doc.xml": "<a/>" })),
      "ZIP_DUPLICATE_PATH",
    );
  });

  it("rejects encrypted entries", () => {
    const archive = zip({ "word/document.xml": "<a/>" });
    expectRejection(
      () => assertArchiveWithinLimits(patchEntry(archive, "word/document.xml", { flags: 1 })),
      "ZIP_ENCRYPTED_ENTRY",
    );
  });

  it("accepts streamed data-descriptor entries (the LibreOffice export shape)", () => {
    // LibreOffice writes every part with general-purpose bit 3: local CRC and
    // sizes are zeroed and a data descriptor follows the payload.
    const archive = zipWithDataDescriptor("word/document.xml", bytes("<a/>"));
    expect(() => assertArchiveWithinLimits(archive)).not.toThrow();
  });

  it("rejects a data descriptor that disagrees with the central directory", () => {
    const archive = zipWithDataDescriptor("word/document.xml", bytes("<a/>"));
    // Byte 4 of the descriptor is the CRC-32 — flip it.
    const descriptorCrc = 30 + "word/document.xml".length + "<a/>".length + 4;
    const corrupted = archive.slice();
    corrupted[descriptorCrc] = (corrupted[descriptorCrc]! ^ 0xff) & 0xff;
    expectRejection(() => assertArchiveWithinLimits(corrupted), "ZIP_LOCAL_HEADER_INVALID");
  });

  it("rejects unsupported compression methods", () => {
    const archive = zip({ "word/document.xml": "<a/>" });
    expectRejection(
      () => assertArchiveWithinLimits(patchEntry(archive, "word/document.xml", { method: 12 })),
      "ZIP_COMPRESSION_UNSUPPORTED",
    );
  });

  it("stops a size-lying DEFLATE bomb at the declared cap", () => {
    // 8 MiB of zeros compresses tiny; the header then claims only 64 KiB. A
    // declared-size-only guard would let office-open inflate the full 8 MiB,
    // so the guard must catch the actual overshoot.
    const archive = zip({ "word/document.xml": new Uint8Array(8 * 1024 * 1024) });
    const lying = patchEntry(archive, "word/document.xml", { uncompressed: 64 * 1024 });
    expectRejection(() => assertArchiveWithinLimits(lying), "ZIP_ENTRY_SIZE_EXCEEDED");
  });

  it("rejects a high declared compression ratio", () => {
    const archive = zip({ "word/document.xml": new Uint8Array(1024 * 1024) });
    expectRejection(
      () => assertArchiveWithinLimits(archive, { ...ARCHIVE_LIMITS, compressionRatio: 10 }),
      "ZIP_COMPRESSION_RATIO_EXCEEDED",
    );
  });

  it("enforces per-entry and aggregate size caps", () => {
    const archive = zip({ "a.bin": new Uint8Array(200), "b.bin": new Uint8Array(200) });
    expectRejection(
      () => assertArchiveWithinLimits(archive, { ...ARCHIVE_LIMITS, entryUncompressedBytes: 100 }),
      "ZIP_ENTRY_SIZE_EXCEEDED",
    );
    expectRejection(
      () =>
        assertArchiveWithinLimits(archive, {
          ...ARCHIVE_LIMITS,
          entryUncompressedBytes: 1000,
          aggregateUncompressedBytes: 300,
        }),
      "ZIP_AGGREGATE_SIZE_EXCEEDED",
    );
  });

  it("enforces the compressed-input cap", () => {
    const archive = zip({ "word/document.xml": "<a/>" });
    expectRejection(
      () => assertArchiveWithinLimits(archive, { ...ARCHIVE_LIMITS, archiveBytes: 16 }),
      "ZIP_ENTRY_SIZE_EXCEEDED",
    );
  });

  it("rejects CRC mismatches", () => {
    const archive = zip({ "word/document.xml": "<a/>" });
    const entry = entriesOf(archive)[0]!;
    expectRejection(
      () =>
        assertArchiveWithinLimits(
          patchEntry(archive, "word/document.xml", { crc: (entry.crc ^ 0xffffffff) >>> 0 }),
        ),
      "ZIP_CRC_MISMATCH",
    );
  });

  it("rejects multi-disk and ZIP64 archives", () => {
    const archive = zip({ "word/document.xml": "<a/>" });
    expectRejection(
      () =>
        assertArchiveWithinLimits(
          patchEocd(archive, (view, eocd) => view.setUint16(eocd + 4, 1, true)),
        ),
      "ZIP_MULTIDISK_UNSUPPORTED",
    );
    expectRejection(
      () =>
        assertArchiveWithinLimits(
          patchEocd(archive, (view, eocd) => {
            view.setUint16(eocd + 8, 0xffff, true);
            view.setUint16(eocd + 10, 0xffff, true);
          }),
        ),
      "ZIP64_UNSUPPORTED",
    );
  });
});

describe("XML admission budgets", () => {
  it("rejects excessive XML depth", () => {
    const depth = 80;
    const xml = `${"<a>".repeat(depth)}text${"</a>".repeat(depth)}`;
    expectRejection(
      () => assertArchiveWithinLimits(zip({ "word/document.xml": xml })),
      "OOXML_XML_DEPTH_EXCEEDED",
    );
  });

  it("rejects excessive XML text, node and attribute counts", () => {
    const text = `<a>${"x".repeat(200)}</a>`;
    expectRejection(
      () =>
        assertArchiveWithinLimits(zip({ "word/document.xml": text }), {
          ...ARCHIVE_LIMITS,
          xmlTextBytes: 100,
        }),
      "OOXML_XML_TEXT_LIMIT_EXCEEDED",
    );

    const nodes = `<a>${"<b/>".repeat(10)}</a>`;
    expectRejection(
      () =>
        assertArchiveWithinLimits(zip({ "word/document.xml": nodes }), {
          ...ARCHIVE_LIMITS,
          xmlNodes: 5,
        }),
      "OOXML_XML_NODE_LIMIT_EXCEEDED",
    );

    const attributes = `<a ${Array.from({ length: 10 }, (_, i) => `x${i}="v"`).join(" ")}/>`;
    expectRejection(
      () =>
        assertArchiveWithinLimits(zip({ "word/document.xml": attributes }), {
          ...ARCHIVE_LIMITS,
          xmlAttributes: 5,
        }),
      "OOXML_XML_ATTRIBUTE_LIMIT_EXCEEDED",
    );
  });

  it("rejects DTDs, malformed XML and relationship floods", () => {
    expectRejection(
      () =>
        assertArchiveWithinLimits(
          zip({ "word/document.xml": `<!DOCTYPE a [<!ENTITY x "y">]><a/>` }),
        ),
      "OOXML_DTD_FORBIDDEN",
    );
    expectRejection(
      () => assertArchiveWithinLimits(zip({ "word/document.xml": "<a>" })),
      "OOXML_XML_INVALID",
    );

    const rels = `<Relationships>${Array.from(
      { length: 20 },
      (_, i) => `<Relationship Id="rId${i}" Type="t" Target="https://example.com/${i}"/>`,
    ).join("")}</Relationships>`;
    expectRejection(
      () =>
        assertArchiveWithinLimits(zip({ "word/_rels/document.xml.rels": rels }), {
          ...ARCHIVE_LIMITS,
          relationships: 10,
        }),
      "OOXML_RELATIONSHIP_LIMIT_EXCEEDED",
    );
  });

  it("rejects oversized external relationship targets", () => {
    const rels = `<Relationships><Relationship Id="rId1" Type="t" TargetMode="External" Target="https://example.com/${"x".repeat(100)}"/></Relationships>`;
    expectRejection(
      () =>
        assertArchiveWithinLimits(zip({ "word/_rels/document.xml.rels": rels }), {
          ...ARCHIVE_LIMITS,
          externalTargetBytes: 32,
        }),
      "OOXML_XML_TEXT_LIMIT_EXCEEDED",
    );
  });
});

describe("media admission budgets", () => {
  it("enforces per-media, count and aggregate caps", () => {
    const one = zip({ "word/media/image1.png": new Uint8Array(100) });
    expectRejection(
      () => assertArchiveWithinLimits(one, { ...ARCHIVE_LIMITS, mediaEntryBytes: 10 }),
      "ZIP_MEDIA_LIMIT_EXCEEDED",
    );

    const two = zip({
      "word/media/image1.png": new Uint8Array(100),
      "word/media/image2.png": new Uint8Array(100),
    });
    expectRejection(
      () => assertArchiveWithinLimits(two, { ...ARCHIVE_LIMITS, mediaEntries: 1 }),
      "ZIP_MEDIA_LIMIT_EXCEEDED",
    );
    expectRejection(
      () => assertArchiveWithinLimits(two, { ...ARCHIVE_LIMITS, mediaAggregateBytes: 150 }),
      "ZIP_MEDIA_LIMIT_EXCEEDED",
    );
  });
});

describe("fuzz corpus", () => {
  const seed = validDocx();

  // Deterministic xorshift — the corpus must be identical on every run.
  let state = 0x12345678;
  const random = (): number => {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    return state >>> 0;
  };

  it("never crashes or hangs on mutated archives", () => {
    const candidates: Uint8Array[] = [];
    for (let round = 0; round < 150; round++) {
      const mutated = seed.slice();
      const flips = 1 + (random() % 8);
      for (let flip = 0; flip < flips; flip++) {
        const offset = random() % mutated.length;
        mutated[offset] = (mutated[offset]! ^ (1 + (random() % 255))) & 0xff;
      }
      candidates.push(mutated);
    }
    for (let round = 0; round < 30; round++) {
      candidates.push(seed.slice(0, random() % seed.length));
    }

    for (const candidate of candidates) {
      try {
        // Malformed inputs must either parse or raise a normal Error — never
        // OOM, hang, or escape as a non-Error.
        parseDOCXSync(candidate);
      } catch (error) {
        expect(error).toBeInstanceOf(Error);
        expect(String((error as Error).message).length).toBeGreaterThan(0);
      }
    }
  }, 60_000);
});
