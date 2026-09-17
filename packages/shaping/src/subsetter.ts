/**
 * Standalone TrueType font subsetter for PDF and DOCX embedding.
 * Extracts minimal required tables (head, hhea, maxp, OS/2, hmtx, cmap, loca, glyf, post, name)
 * for a specific subset of glyph IDs.
 */

interface TableRecord {
  tag: string;
  checksum: number;
  offset: number;
  length: number;
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

export function subsetFont(fontData: Uint8Array, glyphIds: readonly number[]): Uint8Array {
  const view = new DataView(fontData.buffer, fontData.byteOffset, fontData.byteLength);

  // 1. Read Table Directory
  const numTables = view.getUint16(4);
  const tables = new Map<string, TableRecord>();

  for (let i = 0; i < numTables; i++) {
    const recOffset = 12 + i * 16;
    let tag = "";
    for (let c = 0; c < 4; c++) {
      tag += String.fromCharCode(fontData[recOffset + c]!);
    }
    const checksum = view.getUint32(recOffset + 4);
    const offset = view.getUint32(recOffset + 8);
    const length = view.getUint32(recOffset + 12);
    tables.set(tag, { tag, checksum, offset, length });
  }

  const head = tables.get("head");
  const maxp = tables.get("maxp");
  const hhea = tables.get("hhea");
  const hmtx = tables.get("hmtx");
  const loca = tables.get("loca");
  const glyf = tables.get("glyf");

  if (!head || !maxp || !hhea || !hmtx || !loca || !glyf) {
    // If essential TrueType tables are missing (e.g. CFF or bitmap font), return original
    return fontData;
  }

  const indexToLocFormat = view.getInt16(head.offset + 50);
  const numOriginalGlyphs = view.getUint16(maxp.offset + 4);
  const numOfHMetrics = view.getUint16(hhea.offset + 34);

  // 2. Read original loca offsets
  const originalLoca: number[] = [];
  if (indexToLocFormat === 0) {
    // 16-bit offset / 2
    for (let i = 0; i <= numOriginalGlyphs; i++) {
      originalLoca.push(view.getUint16(loca.offset + i * 2) * 2);
    }
  } else {
    // 32-bit offset
    for (let i = 0; i <= numOriginalGlyphs; i++) {
      originalLoca.push(view.getUint32(loca.offset + i * 4));
    }
  }

  // 3. Resolve all used glyph IDs including composite dependencies
  const usedGlyphSet = new Set<number>([0]); // Always include .notdef
  for (const gid of glyphIds) {
    if (gid >= 0 && gid < numOriginalGlyphs) {
      usedGlyphSet.add(gid);
    }
  }

  // Scan composite glyphs to pull in child components
  const toScan = Array.from(usedGlyphSet);
  while (toScan.length > 0) {
    const gid = toScan.pop()!;
    const gStart = originalLoca[gid]!;
    const gEnd = originalLoca[gid + 1]!;
    if (gEnd <= gStart) continue; // Empty glyph

    const glyphOffset = glyf.offset + gStart;
    const numContours = view.getInt16(glyphOffset);
    if (numContours < 0) {
      // Composite glyph
      let compOffset = glyphOffset + 10;
      let flags = 0x0020; // MORE_COMPONENTS initially true
      while (flags & 0x0020) {
        flags = view.getUint16(compOffset);
        const compGlyphId = view.getUint16(compOffset + 2);
        if (!usedGlyphSet.has(compGlyphId) && compGlyphId < numOriginalGlyphs) {
          usedGlyphSet.add(compGlyphId);
          toScan.push(compGlyphId);
        }

        compOffset += 4;
        if (flags & 0x0001) compOffset += 4; // ARG_1_AND_2_ARE_WORDS
        else compOffset += 2;
        if (flags & 0x0008) compOffset += 2; // WE_HAVE_A_SCALE
        else if (flags & 0x0040) compOffset += 4; // WE_HAVE_AN_X_AND_Y_SCALE
        else if (flags & 0x0080) compOffset += 8; // WE_HAVE_A_TWO_BY_TWO
      }
    }
  }

  // 4. Build glyph remapping: oldGid -> newGid
  const sortedGids = Array.from(usedGlyphSet).sort((a, b) => a - b);
  const oldToNew = new Map<number, number>();
  sortedGids.forEach((oldGid, newGid) => oldToNew.set(oldGid, newGid));
  const newNumGlyphs = sortedGids.length;

  // 5. Construct new glyf and loca tables
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
      // Update composite glyph component IDs
      let compOffset = 10;
      let flags = 0x0020;
      while (flags & 0x0020) {
        flags = gView.getUint16(compOffset);
        const oldCompGid = gView.getUint16(compOffset + 2);
        const newCompGid = oldToNew.get(oldCompGid) ?? 0;
        gView.setUint16(compOffset + 2, newCompGid);

        compOffset += 4;
        if (flags & 0x0001) compOffset += 4;
        else compOffset += 2;
        if (flags & 0x0008) compOffset += 2;
        else if (flags & 0x0040) compOffset += 4;
        else if (flags & 0x0080) compOffset += 8;
      }
    }

    // Pad to 2 bytes
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

  // Concatenate new glyf table
  const newGlyfTable = new Uint8Array(currentGlyfOffset);
  let writeOffset = 0;
  for (const chunk of newGlyfChunks) {
    newGlyfTable.set(chunk, writeOffset);
    writeOffset += chunk.length;
  }

  // Construct new loca table (32-bit format)
  const newLocaTable = new Uint8Array(newLocaOffsets.length * 4);
  const locaView = new DataView(newLocaTable.buffer);
  for (let i = 0; i < newLocaOffsets.length; i++) {
    locaView.setUint32(i * 4, newLocaOffsets[i]!);
  }

  // 6. Construct new hmtx table
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

  // 7. Update maxp table
  const newMaxpTable = new Uint8Array(fontData.subarray(maxp.offset, maxp.offset + maxp.length));
  new DataView(newMaxpTable.buffer, newMaxpTable.byteOffset).setUint16(4, newNumGlyphs);

  // 8. Update hhea table (numberOfHMetrics = newNumGlyphs)
  const newHheaTable = new Uint8Array(fontData.subarray(hhea.offset, hhea.offset + hhea.length));
  new DataView(newHheaTable.buffer, newHheaTable.byteOffset).setUint16(34, newNumGlyphs);

  // 9. Update head table (indexToLocFormat = 1)
  const newHeadTable = new Uint8Array(fontData.subarray(head.offset, head.offset + head.length));
  const headView = new DataView(newHeadTable.buffer, newHeadTable.byteOffset);
  headView.setInt16(50, 1); // 32-bit loca
  headView.setUint32(8, 0); // checkSumAdjustment = 0 temporarily

  // 10. Gather output tables
  const outputTables: { tag: string; data: Uint8Array }[] = [
    { tag: "glyf", data: newGlyfTable },
    { tag: "head", data: newHeadTable },
    { tag: "hhea", data: newHheaTable },
    { tag: "hmtx", data: newHmtxTable },
    { tag: "loca", data: newLocaTable },
    { tag: "maxp", data: newMaxpTable },
  ];

  // Pass through other tables if present (OS/2, name, post, cmap)
  const passThroughTags = ["OS/2", "name", "post", "cmap"];
  for (const tag of passThroughTags) {
    const rec = tables.get(tag);
    if (rec) {
      outputTables.push({
        tag,
        data: fontData.subarray(rec.offset, rec.offset + rec.length),
      });
    }
  }

  // Sort tables alphabetically by tag
  outputTables.sort((a, b) => a.tag.localeCompare(b.tag));

  // 11. Calculate layout and assemble final font file
  const outNumTables = outputTables.length;
  let searchRange = 1;
  let entrySelector = 0;
  while (searchRange * 2 <= outNumTables) {
    searchRange *= 2;
    entrySelector++;
  }
  searchRange *= 16;
  const rangeShift = outNumTables * 16 - searchRange;

  const headerSize = 12 + outNumTables * 16;
  let currentOffset = headerSize;

  const tablePlacements: {
    tag: string;
    data: Uint8Array;
    offset: number;
    length: number;
    checksum: number;
  }[] = [];

  for (const t of outputTables) {
    const paddedLen = (t.data.length + 3) & ~3;
    const alignedData = new Uint8Array(paddedLen);
    alignedData.set(t.data);
    const checksum = calculateTableChecksum(alignedData, 0, t.data.length);
    tablePlacements.push({
      tag: t.tag,
      data: alignedData,
      offset: currentOffset,
      length: t.data.length,
      checksum,
    });
    currentOffset += paddedLen;
  }

  const finalFont = new Uint8Array(currentOffset);
  const finalView = new DataView(finalFont.buffer);

  // Write TrueType header
  finalView.setUint32(0, 0x00010000); // sfntVersion
  finalView.setUint16(4, outNumTables);
  finalView.setUint16(6, searchRange);
  finalView.setUint16(8, entrySelector);
  finalView.setUint16(10, rangeShift);

  // Write table records
  for (let i = 0; i < outNumTables; i++) {
    const p = tablePlacements[i]!;
    const recOffset = 12 + i * 16;
    for (let c = 0; c < 4; c++) {
      finalFont[recOffset + c] = p.tag.charCodeAt(c);
    }
    finalView.setUint32(recOffset + 4, p.checksum);
    finalView.setUint32(recOffset + 8, p.offset);
    finalView.setUint32(recOffset + 12, p.length);

    // Copy table data
    finalFont.set(p.data, p.offset);
  }

  // 12. Calculate whole-font checkSumAdjustment
  const totalSum = calculateTableChecksum(finalFont, 0, finalFont.length);
  const checkSumAdjustment = (0xb1b0afba - totalSum) >>> 0;

  const headPlacement = tablePlacements.find((p) => p.tag === "head");
  if (headPlacement) {
    finalView.setUint32(headPlacement.offset + 8, checkSumAdjustment);
  }

  return finalFont;
}
