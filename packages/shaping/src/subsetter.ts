/**
 * Standalone TrueType font subsetter for PDF and DOCX embedding.
 *
 * TrueType (`glyf`/`loca`) fonts are rebuilt with a remapped glyph order:
 * `glyf`, `loca`, `hmtx`, `maxp`, `hhea` and `head` are rewritten, `cmap` is
 * rebuilt so every retained code point maps to its NEW glyph ID, and `post`
 * is rewritten as version 3.0 (no glyph names — the stale name index is
 * exactly what corrupted subsets before). Composite glyph components are
 * pulled in transitively.
 *
 * CFF/bitmap fonts have no `glyf` table; they pass through unchanged (valid
 * bytes, only not reduced) and the caller must use the original glyph IDs.
 *
 * `subsetFontWithPlan` returns the old→new glyph map so callers (PDF
 * `CIDToGIDMap`, DOCX fontTable) can address the remapped glyphs.
 */

interface TableRecord {
  tag: string;
  checksum: number;
  offset: number;
  length: number;
}

/** Result of a subsetting run: the font bytes plus the old→new glyph map. */
export interface SubsetPlan {
  readonly data: Uint8Array;
  /** Old glyph ID → new glyph ID (identity when the font passed through). */
  readonly glyphMap: ReadonlyMap<number, number>;
}

function calculateTableChecksum(data: Uint8Array, offset: number, length: number): number {
  let sum = 0;
  const view = new DataView(data.buffer, data.byteOffset + offset, length);
  const nWords = Math.floor(length / 4);
  for (let i = 0; i < nWords; i++) {
    sum = (sum + view.getUint32(i * 4)) >>> 0;
  }
  const remainder = length % 4;
  if (remainder > 0) {
    let lastWord = 0;
    for (let i = 0; i < remainder; i++) {
      lastWord = (lastWord | (data[offset + nWords * 4 + i]! << (24 - i * 8))) >>> 0;
    }
    sum = (sum + lastWord) >>> 0;
  }
  return sum;
}

/**
 * Read a font's `cmap` into a code point → glyph ID map (union of every
 * subtable the font exposes, so a Windows-BMP table and a full Unicode
 * format-12 table both contribute).
 */
export function readCmap(fontData: Uint8Array): Map<number, number> {
  const out = new Map<number, number>();
  const view = new DataView(fontData.buffer, fontData.byteOffset, fontData.byteLength);

  const tables = readTableDirectory(fontData);
  const cmap = tables.get("cmap");
  if (!cmap || cmap.length < 4) return out;

  const base = cmap.offset;
  const numTables = view.getUint16(base + 2);
  for (let i = 0; i < numTables; i++) {
    const rec = base + 4 + i * 8;
    if (rec + 8 > base + cmap.length) break;
    const subtableOffset = base + view.getUint32(rec + 4);
    parseCmapSubtable(view, subtableOffset, base + cmap.length, out);
  }
  return out;
}

function parseCmapSubtable(
  view: DataView,
  offset: number,
  limit: number,
  out: Map<number, number>,
): void {
  if (offset < 0 || offset + 2 > limit) return;
  const format = view.getUint16(offset);
  switch (format) {
    case 0: {
      if (offset + 262 > limit) return;
      for (let cp = 0; cp < 256; cp++) {
        const gid = view.getUint8(offset + 6 + cp);
        if (gid !== 0 && !out.has(cp)) out.set(cp, gid);
      }
      return;
    }
    case 4: {
      if (offset + 14 > limit) return;
      const segCountX2 = view.getUint16(offset + 6);
      const segCount = segCountX2 / 2;
      if (segCount === 0) return;
      const endCodes = offset + 14;
      const startCodes = endCodes + segCountX2 + 2;
      const idDeltas = startCodes + segCountX2;
      const idRangeOffsets = idDeltas + segCountX2;
      if (idRangeOffsets + segCountX2 > limit) return;
      for (let s = 0; s < segCount; s++) {
        const end = view.getUint16(endCodes + s * 2);
        const start = view.getUint16(startCodes + s * 2);
        const delta = view.getInt16(idDeltas + s * 2);
        const rangeOffset = view.getUint16(idRangeOffsets + s * 2);
        if (start > end) continue;
        for (let cp = start; cp <= end && cp !== 0xffff; cp++) {
          let gid: number;
          if (rangeOffset === 0) {
            gid = (cp + delta) & 0xffff;
          } else {
            const glyphIndex = idRangeOffsets + s * 2 + rangeOffset + (cp - start) * 2;
            if (glyphIndex + 2 > limit) continue;
            gid = view.getUint16(glyphIndex);
            if (gid !== 0) gid = (gid + delta) & 0xffff;
          }
          if (gid !== 0 && !out.has(cp)) out.set(cp, gid);
        }
      }
      return;
    }
    case 6: {
      if (offset + 10 > limit) return;
      const firstCode = view.getUint16(offset + 6);
      const entryCount = view.getUint16(offset + 8);
      for (let i = 0; i < entryCount; i++) {
        const at = offset + 10 + i * 2;
        if (at + 2 > limit) return;
        const gid = view.getUint16(at);
        if (gid !== 0) out.set(firstCode + i, gid);
      }
      return;
    }
    case 12: {
      if (offset + 16 > limit) return;
      const nGroups = view.getUint32(offset + 12);
      for (let g = 0; g < nGroups; g++) {
        const at = offset + 16 + g * 12;
        if (at + 12 > limit) return;
        const start = view.getUint32(at);
        const end = view.getUint32(at + 4);
        const startGid = view.getUint32(at + 8);
        for (let cp = start; cp <= end; cp++) {
          const gid = startGid + (cp - start);
          if (gid !== 0 && !out.has(cp)) out.set(cp, gid);
        }
      }
      return;
    }
    default:
      return;
  }
}

function readTableDirectory(fontData: Uint8Array): Map<string, TableRecord> {
  const view = new DataView(fontData.buffer, fontData.byteOffset, fontData.byteLength);
  const numTables = view.getUint16(4);
  const tables = new Map<string, TableRecord>();
  for (let i = 0; i < numTables; i++) {
    const recOffset = 12 + i * 16;
    if (recOffset + 16 > fontData.byteLength) break;
    let tag = "";
    for (let c = 0; c < 4; c++) {
      tag += String.fromCharCode(fontData[recOffset + c]!);
    }
    tables.set(tag, {
      tag,
      checksum: view.getUint32(recOffset + 4),
      offset: view.getUint32(recOffset + 8),
      length: view.getUint32(recOffset + 12),
    });
  }
  return tables;
}

/** Build the subset's `cmap`: Windows BMP (3,1 format 4) + full Unicode
 *  (3,10 format 12) tables over the retained code points. */
function buildCmapTable(entries: ReadonlyMap<number, number>): Uint8Array {
  const bmp: { cp: number; gid: number }[] = [];
  const all: { cp: number; gid: number }[] = [];
  for (const [cp, gid] of entries) {
    if (cp <= 0xffff && cp !== 0xffff) bmp.push({ cp, gid });
    all.push({ cp, gid });
  }
  bmp.sort((a, b) => a.cp - b.cp);
  all.sort((a, b) => a.cp - b.cp);

  const format4 = buildCmapFormat4(bmp);
  const format12 = bmp.length === all.length ? undefined : buildCmapFormat12(all);

  const subtables: { platformId: number; encodingId: number; data: Uint8Array }[] = [
    { platformId: 3, encodingId: 1, data: format4 },
  ];
  // The (3,10) table must carry the FULL map (Windows consults it for every
  // lookup once present), not just the supplementary planes.
  if (format12) subtables.push({ platformId: 3, encodingId: 10, data: format12 });

  const headerLength = 4 + subtables.length * 8;
  let total = headerLength;
  for (const t of subtables) total += (t.data.length + 3) & ~3;

  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);
  view.setUint16(0, 0); // version
  view.setUint16(2, subtables.length);
  let at = headerLength;
  for (let i = 0; i < subtables.length; i++) {
    const t = subtables[i]!;
    const rec = 4 + i * 8;
    view.setUint16(rec, t.platformId);
    view.setUint16(rec + 2, t.encodingId);
    view.setUint32(rec + 4, at);
    out.set(t.data, at);
    at += (t.data.length + 3) & ~3;
  }
  return out;
}

function buildCmapFormat4(entries: readonly { cp: number; gid: number }[]): Uint8Array {
  interface Segment {
    start: number;
    end: number;
    delta: number;
  }
  const segments: Segment[] = [];
  for (const { cp, gid } of entries) {
    const delta = (gid - cp) & 0xffff;
    const last = segments[segments.length - 1];
    if (last && cp === last.end + 1 && delta === last.delta) {
      last.end = cp;
    } else {
      segments.push({ start: cp, end: cp, delta });
    }
  }
  // The spec requires a final segment mapping U+FFFF to glyph 0.
  segments.push({ start: 0xffff, end: 0xffff, delta: 1 });

  const segCount = segments.length;
  const segCountX2 = segCount * 2;
  let searchRange = 1;
  let entrySelector = 0;
  while (searchRange * 2 <= segCount) {
    searchRange *= 2;
    entrySelector++;
  }
  searchRange *= 2;
  const rangeShift = segCountX2 - searchRange;

  const length = 16 + segCountX2 * 4;
  const out = new Uint8Array(length);
  const view = new DataView(out.buffer);
  view.setUint16(0, 4);
  view.setUint16(2, length);
  view.setUint16(4, 0); // language
  view.setUint16(6, segCountX2);
  view.setUint16(8, searchRange);
  view.setUint16(10, entrySelector);
  view.setUint16(12, rangeShift);

  const endCodes = 14;
  const startCodes = endCodes + segCountX2 + 2;
  const idDeltas = startCodes + segCountX2;
  const idRangeOffsets = idDeltas + segCountX2;
  for (let s = 0; s < segCount; s++) {
    const seg = segments[s]!;
    view.setUint16(endCodes + s * 2, seg.end);
    view.setUint16(startCodes + s * 2, seg.start);
    view.setInt16(idDeltas + s * 2, seg.delta > 0x7fff ? seg.delta - 0x10000 : seg.delta);
    view.setUint16(idRangeOffsets + s * 2, 0);
  }
  return out;
}

function buildCmapFormat12(entries: readonly { cp: number; gid: number }[]): Uint8Array {
  const groups: { start: number; end: number; startGid: number }[] = [];
  for (const { cp, gid } of entries) {
    const last = groups[groups.length - 1];
    if (last && cp === last.end + 1 && gid === last.startGid + (cp - last.start)) {
      last.end = cp;
    } else {
      groups.push({ start: cp, end: cp, startGid: gid });
    }
  }
  const out = new Uint8Array(16 + groups.length * 12);
  const view = new DataView(out.buffer);
  view.setUint16(0, 12);
  view.setUint16(2, 0);
  view.setUint32(4, out.length);
  view.setUint32(8, 0);
  view.setUint32(12, groups.length);
  for (let g = 0; g < groups.length; g++) {
    const group = groups[g]!;
    view.setUint32(16 + g * 12, group.start);
    view.setUint32(20 + g * 12, group.end);
    view.setUint32(24 + g * 12, group.startGid);
  }
  return out;
}

/** A `post` version 3.0 table: the 32-byte header only, no glyph names. */
function buildPostV3(original: Uint8Array | undefined): Uint8Array {
  const out = new Uint8Array(32);
  if (original && original.length >= 32) out.set(original.subarray(0, 32));
  const view = new DataView(out.buffer);
  view.setUint32(0, 0x00030000);
  view.setUint32(12, 0); // minMemType42
  view.setUint32(16, 0); // maxMemType42
  view.setUint32(20, 0); // minMemType1
  view.setUint32(24, 0); // maxMemType1
  return out;
}

function assembleSfnt(tables: { tag: string; data: Uint8Array }[]): Uint8Array {
  const sorted = [...tables].sort((a, b) => a.tag.localeCompare(b.tag));
  const outNumTables = sorted.length;
  let searchRange = 1;
  let entrySelector = 0;
  while (searchRange * 2 <= outNumTables) {
    searchRange *= 2;
    entrySelector++;
  }
  searchRange *= 16;
  const rangeShift = outNumTables * 16 - searchRange;

  let currentOffset = 12 + outNumTables * 16;
  const placements = sorted.map((t) => {
    const paddedLen = (t.data.length + 3) & ~3;
    const aligned = new Uint8Array(paddedLen);
    aligned.set(t.data);
    const placement = {
      tag: t.tag,
      data: aligned,
      offset: currentOffset,
      length: t.data.length,
      checksum: calculateTableChecksum(aligned, 0, t.data.length),
    };
    currentOffset += paddedLen;
    return placement;
  });

  const finalFont = new Uint8Array(currentOffset);
  const finalView = new DataView(finalFont.buffer);
  finalView.setUint32(0, 0x00010000);
  finalView.setUint16(4, outNumTables);
  finalView.setUint16(6, searchRange);
  finalView.setUint16(8, entrySelector);
  finalView.setUint16(10, rangeShift);

  for (let i = 0; i < outNumTables; i++) {
    const p = placements[i]!;
    const recOffset = 12 + i * 16;
    for (let c = 0; c < 4; c++) finalFont[recOffset + c] = p.tag.charCodeAt(c);
    finalView.setUint32(recOffset + 4, p.checksum);
    finalView.setUint32(recOffset + 8, p.offset);
    finalView.setUint32(recOffset + 12, p.length);
    finalFont.set(p.data, p.offset);
  }

  const totalSum = calculateTableChecksum(finalFont, 0, finalFont.length);
  const checkSumAdjustment = (0xb1b0afba - totalSum) >>> 0;
  const headPlacement = placements.find((p) => p.tag === "head");
  if (headPlacement) finalView.setUint32(headPlacement.offset + 8, checkSumAdjustment);

  return finalFont;
}

/** TrueType subsetting with the old→new glyph map (identity for pass-through fonts). */
export function subsetFontWithPlan(fontData: Uint8Array, glyphIds: readonly number[]): SubsetPlan {
  const view = new DataView(fontData.buffer, fontData.byteOffset, fontData.byteLength);
  const tables = readTableDirectory(fontData);

  const head = tables.get("head");
  const maxp = tables.get("maxp");
  const hhea = tables.get("hhea");
  const hmtx = tables.get("hmtx");
  const loca = tables.get("loca");
  const glyf = tables.get("glyf");

  if (!head || !maxp || !hhea || !hmtx || !loca || !glyf) {
    // CFF/bitmap fonts: return the original bytes (valid, simply not reduced);
    // the glyph IDs are unchanged, hence the identity map.
    const identity = new Map<number, number>();
    for (const gid of glyphIds) identity.set(gid, gid);
    return { data: fontData, glyphMap: identity };
  }

  const indexToLocFormat = view.getInt16(head.offset + 50);
  const numOriginalGlyphs = view.getUint16(maxp.offset + 4);
  const numOfHMetrics = view.getUint16(hhea.offset + 34);

  const originalLoca: number[] = [];
  if (indexToLocFormat === 0) {
    for (let i = 0; i <= numOriginalGlyphs; i++) {
      originalLoca.push(view.getUint16(loca.offset + i * 2) * 2);
    }
  } else {
    for (let i = 0; i <= numOriginalGlyphs; i++) {
      originalLoca.push(view.getUint32(loca.offset + i * 4));
    }
  }

  const usedGlyphSet = new Set<number>([0]);
  for (const gid of glyphIds) {
    if (gid >= 0 && gid < numOriginalGlyphs) usedGlyphSet.add(gid);
  }

  // Composite closure: pull in every component glyph transitively.
  const toScan = Array.from(usedGlyphSet);
  while (toScan.length > 0) {
    const gid = toScan.pop()!;
    const gStart = originalLoca[gid]!;
    const gEnd = originalLoca[gid + 1]!;
    if (gEnd <= gStart) continue;

    const glyphOffset = glyf.offset + gStart;
    const numContours = view.getInt16(glyphOffset);
    if (numContours < 0) {
      let compOffset = glyphOffset + 10;
      let flags = 0x0020;
      while (flags & 0x0020) {
        flags = view.getUint16(compOffset);
        const compGlyphId = view.getUint16(compOffset + 2);
        if (!usedGlyphSet.has(compGlyphId) && compGlyphId < numOriginalGlyphs) {
          usedGlyphSet.add(compGlyphId);
          toScan.push(compGlyphId);
        }
        compOffset += 4;
        if (flags & 0x0001) compOffset += 4;
        else compOffset += 2;
        if (flags & 0x0008) compOffset += 2;
        else if (flags & 0x0040) compOffset += 4;
        else if (flags & 0x0080) compOffset += 8;
      }
    }
  }

  const sortedGids = Array.from(usedGlyphSet).sort((a, b) => a - b);
  const oldToNew = new Map<number, number>();
  sortedGids.forEach((oldGid, newGid) => oldToNew.set(oldGid, newGid));
  const newNumGlyphs = sortedGids.length;

  const newGlyfChunks: Uint8Array[] = [];
  const newLocaOffsets: number[] = [0];
  let currentGlyfOffset = 0;

  for (const oldGid of sortedGids) {
    const gStart = originalLoca[oldGid]!;
    const gEnd = originalLoca[oldGid + 1]!;
    const gLen = gEnd - gStart;

    if (gLen === 0) {
      newLocaOffsets.push(currentGlyfOffset);
      continue;
    }

    const glyphBytes = new Uint8Array(fontData.subarray(glyf.offset + gStart, glyf.offset + gEnd));
    const gView = new DataView(glyphBytes.buffer, glyphBytes.byteOffset, glyphBytes.byteLength);
    const numContours = gView.getInt16(0);

    if (numContours < 0) {
      let compOffset = 10;
      let flags = 0x0020;
      while (flags & 0x0020) {
        flags = gView.getUint16(compOffset);
        const oldCompGid = gView.getUint16(compOffset + 2);
        gView.setUint16(compOffset + 2, oldToNew.get(oldCompGid) ?? 0);
        compOffset += 4;
        if (flags & 0x0001) compOffset += 4;
        else compOffset += 2;
        if (flags & 0x0008) compOffset += 2;
        else if (flags & 0x0040) compOffset += 4;
        else if (flags & 0x0080) compOffset += 8;
      }
    }

    const paddedLen = (gLen + 1) & ~1;
    let chunk = glyphBytes;
    if (paddedLen > gLen) {
      chunk = new Uint8Array(paddedLen);
      chunk.set(glyphBytes);
    }
    newGlyfChunks.push(chunk);
    currentGlyfOffset += paddedLen;
    newLocaOffsets.push(currentGlyfOffset);
  }

  const newGlyfTable = new Uint8Array(currentGlyfOffset);
  let writeOffset = 0;
  for (const chunk of newGlyfChunks) {
    newGlyfTable.set(chunk, writeOffset);
    writeOffset += chunk.length;
  }

  const newLocaTable = new Uint8Array(newLocaOffsets.length * 4);
  const locaView = new DataView(newLocaTable.buffer);
  for (let i = 0; i < newLocaOffsets.length; i++) {
    locaView.setUint32(i * 4, newLocaOffsets[i]!);
  }

  const newHmtxTable = new Uint8Array(newNumGlyphs * 4);
  const hmtxView = new DataView(newHmtxTable.buffer);
  for (let newGid = 0; newGid < newNumGlyphs; newGid++) {
    const oldGid = sortedGids[newGid]!;
    let advance = 0;
    let lsb = 0;
    if (oldGid < numOfHMetrics) {
      advance = view.getUint16(hmtx.offset + oldGid * 4);
      lsb = view.getInt16(hmtx.offset + oldGid * 4 + 2);
    } else {
      advance = view.getUint16(hmtx.offset + (numOfHMetrics - 1) * 4);
      lsb = view.getInt16(hmtx.offset + numOfHMetrics * 4 + (oldGid - numOfHMetrics) * 2);
    }
    hmtxView.setUint16(newGid * 4, advance);
    hmtxView.setInt16(newGid * 4 + 2, lsb);
  }

  const newMaxpTable = new Uint8Array(fontData.subarray(maxp.offset, maxp.offset + maxp.length));
  new DataView(newMaxpTable.buffer, newMaxpTable.byteOffset).setUint16(4, newNumGlyphs);

  const newHheaTable = new Uint8Array(fontData.subarray(hhea.offset, hhea.offset + hhea.length));
  new DataView(newHheaTable.buffer, newHheaTable.byteOffset).setUint16(34, newNumGlyphs);

  const newHeadTable = new Uint8Array(fontData.subarray(head.offset, head.offset + head.length));
  const headView = new DataView(newHeadTable.buffer, newHeadTable.byteOffset);
  headView.setInt16(50, 1);
  headView.setUint32(8, 0);

  // Rebuild the cmap over the retained glyphs so lookup hits the new IDs.
  const cmapEntries = new Map<number, number>();
  for (const [cp, oldGid] of readCmap(fontData)) {
    const newGid = oldToNew.get(oldGid);
    if (newGid !== undefined) cmapEntries.set(cp, newGid);
  }

  const postRec = tables.get("post");
  const outputTables: { tag: string; data: Uint8Array }[] = [
    { tag: "glyf", data: newGlyfTable },
    { tag: "head", data: newHeadTable },
    { tag: "hhea", data: newHheaTable },
    { tag: "hmtx", data: newHmtxTable },
    { tag: "loca", data: newLocaTable },
    { tag: "maxp", data: newMaxpTable },
    { tag: "cmap", data: buildCmapTable(cmapEntries) },
    {
      tag: "post",
      data: buildPostV3(
        postRec ? fontData.subarray(postRec.offset, postRec.offset + postRec.length) : undefined,
      ),
    },
  ];

  // name and OS/2 carry identity/licensing data the embedder needs.
  for (const tag of ["OS/2", "name"]) {
    const rec = tables.get(tag);
    if (rec)
      outputTables.push({ tag, data: fontData.subarray(rec.offset, rec.offset + rec.length) });
  }
  // Keep TrueType hinting programs (small; PDF ignores them but other consumers may not).
  for (const tag of ["cvt ", "fpgm", "prep", "gasp"]) {
    const rec = tables.get(tag);
    if (rec)
      outputTables.push({ tag, data: fontData.subarray(rec.offset, rec.offset + rec.length) });
  }

  // First/last BMP code point in OS/2, where present (offsets are version-stable).
  const os2 = outputTables.find((t) => t.tag === "OS/2");
  if (os2 && os2.data.length >= 68) {
    let first = 0xffff;
    let last = 0;
    for (const [cp] of cmapEntries) {
      if (cp > 0xffff) continue;
      if (cp < first) first = cp;
      if (cp > last) last = cp;
    }
    const os2View = new DataView(os2.data.buffer, os2.data.byteOffset, os2.data.length);
    os2View.setUint16(64, first === 0xffff ? 0 : first);
    os2View.setUint16(66, last);
  }

  return { data: assembleSfnt(outputTables), glyphMap: oldToNew };
}

/** Subset a TrueType font down to `glyphIds` (plus composite components and
 *  `.notdef`). CFF/bitmap fonts pass through unchanged. */
export function subsetFont(fontData: Uint8Array, glyphIds: readonly number[]): Uint8Array {
  return subsetFontWithPlan(fontData, glyphIds).data;
}
