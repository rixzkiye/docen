// Font-collection (.ttc/.otc) support: the engine registers a face by index,
// which is how system CJK fonts ship (Noto Sans CJK `.ttc`, Windows
// msyh.ttc/simsun.ttc). The oracle below parses the collection's own cmap and
// hmtx tables for the requested face — no WASM — so a face-index mistake in
// any consumer (shape, metrics, outlines) surfaces as a mismatch.

// @vitest-environment node
import fs from "node:fs";

import { beforeAll, describe, expect, it } from "vitest";

import { createFontRefSync, initShapingWasm } from "../src/index.js";

/** System Noto Sans CJK collection (JP/KR/SC/TC faces); absent on some CI. */
const TTC_PATH = "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc";
const CJK_CHAR = "直"; // U+76F4 — differs between the JP and SC faces.
const CJK_CODE = 0x76f4;

interface SfntTables {
  [tag: string]: { offset: number; length: number };
}

/** The table directory of face `index` inside a TTC (offsets are absolute). */
function faceTables(bytes: Uint8Array, index: number): SfntTables {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (String.fromCharCode(bytes[0]!, bytes[1]!, bytes[2]!, bytes[3]!) !== "ttcf") {
    throw new Error("not a collection");
  }
  const faceOffset = view.getUint32(12 + index * 4);
  const numTables = view.getUint16(faceOffset + 4);
  const tables: SfntTables = {};
  for (let i = 0; i < numTables; i++) {
    const rec = faceOffset + 12 + i * 16;
    const tag = String.fromCharCode(bytes[rec]!, bytes[rec + 1]!, bytes[rec + 2]!, bytes[rec + 3]!);
    tables[tag] = { offset: view.getUint32(rec + 8), length: view.getUint32(rec + 12) };
  }
  return tables;
}

/** Unicode → glyph id from a face's cmap (formats 12 and 4). */
function glyphOf(bytes: Uint8Array, tables: SfntTables, code: number): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const cmap = tables.cmap!.offset;
  const numTables = view.getUint16(cmap + 2);
  let best = -1;
  for (let i = 0; i < numTables; i++) {
    const rec = cmap + 4 + i * 8;
    const platform = view.getUint16(rec);
    const encoding = view.getUint16(rec + 2);
    const offset = view.getUint32(rec + 4);
    // Prefer a full-Unicode subtable (platform 0, or Windows UCS-4).
    if (platform === 0 || (platform === 3 && encoding === 10)) best = cmap + offset;
    else if (best < 0 && platform === 3 && encoding === 1) best = cmap + offset;
  }
  if (best < 0) throw new Error("no usable cmap");
  const format = view.getUint16(best);
  if (format === 12) {
    const groups = view.getUint32(best + 12);
    for (let i = 0; i < groups; i++) {
      const g = best + 16 + i * 12;
      const start = view.getUint32(g);
      const end = view.getUint32(g + 4);
      if (code >= start && code <= end) return view.getUint32(g + 8) + (code - start);
    }
    throw new Error(`U+${code.toString(16)} unmapped`);
  }
  if (format === 4) {
    const segCountX2 = view.getUint16(best + 6);
    const endPos = best + 14;
    const startPos = endPos + segCountX2 + 2;
    const deltaPos = startPos + segCountX2;
    const rangePos = deltaPos + segCountX2;
    for (let i = 0; i < segCountX2 / 2; i++) {
      const end = view.getUint16(endPos + i * 2);
      if (code > end) continue;
      const start = view.getUint16(startPos + i * 2);
      if (code < start) break;
      const rangeOffset = view.getUint16(rangePos + i * 2);
      const delta = view.getInt16(deltaPos + i * 2);
      if (rangeOffset === 0) return (code + delta) & 0xffff;
      const at = rangePos + i * 2 + rangeOffset + (code - start) * 2;
      const gid = view.getUint16(at);
      return gid === 0 ? 0 : (gid + delta) & 0xffff;
    }
    throw new Error(`U+${code.toString(16)} unmapped`);
  }
  throw new Error(`unhandled cmap format ${format}`);
}

/** Advance width in font units for a glyph from a face's hmtx. */
function advanceOf(bytes: Uint8Array, tables: SfntTables, gid: number): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const numHMetrics = view.getUint16(tables.hhea!.offset + 34);
  const hmtx = tables.hmtx!.offset;
  const at = gid < numHMetrics ? hmtx + gid * 4 : hmtx + (numHMetrics - 1) * 4;
  return view.getUint16(at);
}

const hasTtc = fs.existsSync(TTC_PATH);

describe.skipIf(!hasTtc)("font collection (.ttc) face indexing", () => {
  beforeAll(async () => {
    await initShapingWasm();
  });

  it("registers face 0 of a collection when no index is given", () => {
    const bytes = new Uint8Array(fs.readFileSync(TTC_PATH));
    const font = createFontRefSync(bytes);
    try {
      expect(font.familyName).toBe("Noto Sans CJK JP");
      expect(font.shape(CJK_CHAR).glyphs[0]!.glyphId).toBe(
        glyphOf(bytes, faceTables(bytes, 0), CJK_CODE),
      );
    } finally {
      font.dispose();
    }
  });

  it("shapes, measures and outlines face 2 (SC) at its own index", () => {
    const bytes = new Uint8Array(fs.readFileSync(TTC_PATH));
    const tables = faceTables(bytes, 2);
    const oracleGid = glyphOf(bytes, tables, CJK_CODE);
    const oracleAdvance = advanceOf(bytes, tables, oracleGid);
    const font = createFontRefSync(bytes, { index: 2 });

    try {
      expect(font.familyName).toBe("Noto Sans CJK SC");
      const shaped = font.shape(CJK_CHAR);
      expect(shaped.glyphs.map((g) => g.glyphId)).toEqual([oracleGid]);
      expect(shaped.totalAdvance).toBe(oracleAdvance);
      expect(font.metrics.unitsPerEm).toBe(1000);
      expect(font.getGlyphOutline(oracleGid).length).toBeGreaterThan(0);
      expect(font.identity.glyphCount).toBeGreaterThan(40000);
    } finally {
      font.dispose();
    }
  });

  it("selects different faces for JP and SC", () => {
    const bytes = new Uint8Array(fs.readFileSync(TTC_PATH));
    const jp = createFontRefSync(bytes, { index: 0 });
    const sc = createFontRefSync(bytes, { index: 2 });
    try {
      expect(jp.familyName).toBe("Noto Sans CJK JP");
      expect(sc.familyName).toBe("Noto Sans CJK SC");
      expect(jp.shape(CJK_CHAR).glyphs[0]!.glyphId).not.toBe(sc.shape(CJK_CHAR).glyphs[0]!.glyphId);
    } finally {
      jp.dispose();
      sc.dispose();
    }
  });

  it("refuses to subset a collection face (flat-sfnt subsetter)", async () => {
    const bytes = new Uint8Array(fs.readFileSync(TTC_PATH));
    const font = createFontRefSync(bytes, { index: 2 });
    try {
      await expect(font.subset([1, 2, 3])).rejects.toThrow(/collection/);
    } finally {
      font.dispose();
    }
  });
});
