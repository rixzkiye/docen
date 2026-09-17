/**
 * Generates an Adobe PostScript ToUnicode CMap stream for PDF Type 0 fonts.
 * Maps CIDs (glyph IDs) to Unicode UTF-16 code units.
 */
export function generateToUnicodeCMap(
  cidToUnicodeMap: Map<number, number> | [number, number][],
  cmapName = "Custom-ToUnicode",
): string {
  const entries: [number, number][] = Array.isArray(cidToUnicodeMap)
    ? cidToUnicodeMap
    : Array.from(cidToUnicodeMap.entries());

  // Sort by CID
  entries.sort((a, b) => a[0] - b[0]);

  const toHex4 = (n: number) => n.toString(16).toUpperCase().padStart(4, "0");

  const toUnicodeHex = (cp: number) => {
    if (cp <= 0xffff) {
      return toHex4(cp);
    }
    // Surrogate pair
    const s = String.fromCodePoint(cp);
    return toHex4(s.charCodeAt(0)) + toHex4(s.charCodeAt(1));
  };

  const lines: string[] = [
    "/CIDInit /ProcSet findresource begin",
    "12 dict begin",
    "begincmap",
    "/CIDSystemInfo <<",
    "  /Registry (Adobe)",
    "  /Ordering (UCS)",
    "  /Supplement 0",
    ">> def",
    `/CMapName /${cmapName} def`,
    "/CMapType 2 def",
    "1 begincodespacerange",
    "<0000> <FFFF>",
    "endcodespacerange",
  ];

  // Adobe spec recommends chunks of max 100 entries per beginbfchar
  const CHUNK_SIZE = 100;
  for (let i = 0; i < entries.length; i += CHUNK_SIZE) {
    const chunk = entries.slice(i, i + CHUNK_SIZE);
    lines.push(`${chunk.length} beginbfchar`);
    for (const [cid, cp] of chunk) {
      lines.push(`<${toHex4(cid)}> <${toUnicodeHex(cp)}>`);
    }
    lines.push("endbfchar");
  }

  lines.push("endcmap");
  lines.push("CMapName currentdict /CMap defineresource pop");
  lines.push("end");
  lines.push("end\n");

  return lines.join("\n");
}
