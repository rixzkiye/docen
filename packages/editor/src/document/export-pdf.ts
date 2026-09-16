// PDF export — the paginated canvas snapshots (printSnapshots) flatten into a
// PDF 1.4 file with a searchable, selectable text layer and clickable links.
// Each page combines an opaque JPEG image XObject for visual fidelity with
// invisible (3 Tr) text operators aligned to the layout's character and word
// coordinates. Supports standard Type 1 fonts for Latin and a universal Type 0
// composite font with a /ToUnicode CMap for CJK and extended Unicode.
// Hand-rolled object/xref layout with zero external dependencies.

import type { FlowPage, LaidOutBlock } from "@docen/layout";
import {
  fieldLabelOf,
  lineBaselineDepthPx,
  vertAlignBaselineShiftPx,
  vertAlignedSizePx,
} from "@docen/layout";

import type { CanvasStageSection } from "./canvas/stage";

/** One text span placed in the PDF text layer with exact coordinates (in pt). */
export interface PdfTextSpan {
  text: string;
  /** PDF x in pt (distance from page left edge). */
  x: number;
  /** PDF baseline y in pt (distance from page bottom edge). */
  y: number;
  /** Laid out run width in pt. */
  width: number;
  /** Line/run height in pt. */
  height: number;
  /** Font size in pt. */
  fontSize: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  /** Structure element tag (P, H1, H2, H3, Table). */
  tag?: string;
}

/** One interactive link annotation in the PDF page (in pt). */
export interface PdfLinkAnnotation {
  /** [left, bottom, right, top] in PDF points. */
  rect: [number, number, number, number];
  /** Target URL (http/https) or internal bookmark anchor (#bookmark). */
  url: string;
  title?: string;
}

/** One page snapshot from the stage's printSnapshots(): the page's CSS-pixel
 *  paper size, its image (PNG data URL or raw JPEG buffer), and optional
 *  text layer and link annotations. */
export interface PdfPageShot {
  width: number;
  height: number;
  url?: string;
  jpeg?: Uint8Array<ArrayBuffer>;
  textSpans?: PdfTextSpan[];
  links?: PdfLinkAnnotation[];
}

/** Options for PDF document generation. */
export interface PdfExportOptions {
  metadata?: {
    title?: string;
    author?: string;
    subject?: string;
    keywords?: string;
    creator?: string;
    producer?: string;
  };
  /** Produce a Tagged PDF with /MarkInfo and /StructTreeRoot (default: true). */
  tagged?: boolean;
}

/** Decode a snapshot PNG and flatten it onto white — the print canvases are
 *  transparent, JPEG has no alpha, and PDF viewers composite images on black,
 *  so the alpha must become paper white before encoding. Returns the JPEG bytes. */
async function jpegOf(
  shot: PdfPageShot,
): Promise<{ jpeg: Uint8Array<ArrayBuffer>; width: number; height: number }> {
  if (shot.jpeg) {
    return { jpeg: shot.jpeg, width: shot.width, height: shot.height };
  }
  if (!shot.url) {
    throw new Error("PdfPageShot must provide either url or jpeg");
  }
  if (typeof Image === "undefined" || typeof document === "undefined") {
    // Non-browser or headless fallback without Canvas
    if (shot.url.startsWith("data:image/")) {
      const comma = shot.url.indexOf(",");
      const bin = atob(shot.url.slice(comma + 1));
      const jpeg = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) jpeg[i] = bin.charCodeAt(i);
      return { jpeg, width: shot.width, height: shot.height };
    }
    throw new Error("Canvas environment unavailable for snapshot decode");
  }
  const img = new Image();
  await new Promise<void>((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error("page snapshot failed to decode"));
    img.src = shot.url!;
  });
  const canvas = document.createElement("canvas");
  canvas.width = img.naturalWidth;
  canvas.height = img.naturalHeight;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("canvas 2d context unavailable");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0);
  const dataUrl = canvas.toDataURL("image/jpeg", 0.92);
  const bin = atob(dataUrl.slice(dataUrl.indexOf(",") + 1));
  const jpeg = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) jpeg[i] = bin.charCodeAt(i);
  return { jpeg, width: img.naturalWidth, height: img.naturalHeight };
}

/** Serialize the whole file: string parts stream through TextEncoder, binary
 *  parts (the JPEG streams) pass through as-is; every object's byte offset is
 *  recorded for the xref table. */
function pdfWriter(): {
  push: (part: string | Uint8Array) => void;
  offset: () => number;
  bytes: () => Uint8Array<ArrayBuffer>;
} {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  let length = 0;
  const push = (part: string | Uint8Array): void => {
    const bytes = typeof part === "string" ? encoder.encode(part) : part;
    chunks.push(bytes);
    length += bytes.length;
  };
  const bytes = (): Uint8Array<ArrayBuffer> => {
    const out = new Uint8Array(new ArrayBuffer(length));
    let at = 0;
    for (const c of chunks) {
      out.set(c, at);
      at += c.length;
    }
    return out;
  };
  return { push, offset: () => length, bytes };
}

/** The 10-digit zero-padded xref entry offset. */
const xrefAt = (offset: number): string => String(offset).padStart(10, "0");

/** Escape special characters for PDF literal strings ( ... ). */
export function escapePdfString(str: string): string {
  return str
    .replace(/\\/g, "\\\\")
    .replace(/\(/g, "\\(")
    .replace(/\)/g, "\\)")
    .replace(/\r/g, "\\r")
    .replace(/\n/g, "\\n");
}

/** Encode a Unicode string as 2-byte hexadecimal CIDs for /Identity-H fonts. */
export function encodeHexUtf16(str: string): string {
  let hex = "";
  for (let i = 0; i < str.length; i++) {
    hex += str.charCodeAt(i).toString(16).padStart(4, "0");
  }
  return hex;
}

/** Determine whether a string can be safely represented in standard 7-bit ASCII. */
function isAsciiPrintable(str: string): boolean {
  return /^[\x20-\x7E]*$/.test(str);
}

/** Format a Date for PDF metadata (D:YYYYMMDDHHmmSSZ). */
function formatPdfDate(date: Date = new Date()): string {
  const pad = (n: number) => String(n).padStart(2, "0");
  const y = date.getUTCFullYear();
  const m = pad(date.getUTCMonth() + 1);
  const d = pad(date.getUTCDate());
  const h = pad(date.getUTCHours());
  const min = pad(date.getUTCMinutes());
  const s = pad(date.getUTCSeconds());
  return `D:${y}${m}${d}${h}${min}${s}Z`;
}

/** Select the best matching PDF font name and encoding mode for a text span. */
function fontForSpan(span: PdfTextSpan): { fontName: string; isUnicode: boolean } {
  if (!isAsciiPrintable(span.text)) {
    return { fontName: "F_Uni", isUnicode: true };
  }
  const fam = (span.fontFamily ?? "").toLowerCase();
  const bold = !!span.bold;
  const italic = !!span.italic;

  if (
    fam.includes("times") ||
    fam.includes("serif") ||
    fam.includes("georgia") ||
    fam.includes("garamond")
  ) {
    if (bold && italic) return { fontName: "F8", isUnicode: false };
    if (bold) return { fontName: "F6", isUnicode: false };
    if (italic) return { fontName: "F7", isUnicode: false };
    return { fontName: "F5", isUnicode: false };
  }
  if (fam.includes("courier") || fam.includes("mono") || fam.includes("consolas")) {
    if (bold) return { fontName: "F10", isUnicode: false };
    return { fontName: "F9", isUnicode: false };
  }
  // Default: Helvetica / Sans-serif
  if (bold && italic) return { fontName: "F4", isUnicode: false };
  if (bold) return { fontName: "F2", isUnicode: false };
  if (italic) return { fontName: "F3", isUnicode: false };
  return { fontName: "F1", isUnicode: false };
}

/** Build the PDF blob from page snapshots, text layers, and link annotations.
 *  Pages keep their paper size (CSS px → pt at 72/96). The canvas image fills
 *  the page, with invisible text operators and link annotations overlaid. */
export async function pagesToPdf(
  shots: readonly PdfPageShot[],
  options?: PdfExportOptions,
): Promise<Blob> {
  const { push, offset, bytes } = pdfWriter();
  // %PDF-1.4 plus binary-marker comment line (raw bytes)
  push("%PDF-1.4\n");
  push(new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]));

  const offsets: number[] = [0]; // object 0 is free
  const record = (): void => {
    offsets.push(offset());
  };

  const jpegs = await Promise.all(shots.map((s) => jpegOf(s)));
  const pt = (px: number): number => (px * 72) / 96;

  // Object numbering allocator
  let nextId = 1;
  const allocId = () => nextId++;

  const catalogId = allocId();
  const pagesId = allocId();
  const isTagged = options?.tagged !== false;
  const structTreeRootId = isTagged ? allocId() : undefined;

  // Standard Type 1 fonts + Unicode Type 0 font
  const fontIds: Record<string, number> = {
    F1: allocId(),
    F2: allocId(),
    F3: allocId(),
    F4: allocId(),
    F5: allocId(),
    F6: allocId(),
    F7: allocId(),
    F8: allocId(),
    F9: allocId(),
    F10: allocId(),
    F_Uni: allocId(),
  };
  const cidFontId = allocId();
  const fontDescId = allocId();
  const toUnicodeId = allocId();

  const infoId = allocId();

  interface PageAlloc {
    pageId: number;
    contentId: number;
    imageId: number;
    annotIds: number[];
    structElemIds: number[];
  }

  const pageAllocs: PageAlloc[] = [];
  for (const shot of shots) {
    const pageId = allocId();
    const contentId = allocId();
    const imageId = allocId();
    const annotIds = (shot.links ?? []).map(() => allocId());
    const spans = shot.textSpans ?? [];
    const structElemIds = isTagged && spans.length > 0 ? [allocId()] : [];
    pageAllocs.push({ pageId, contentId, imageId, annotIds, structElemIds });
  }

  // 1. Catalog Object
  record();
  let catDict = `<< /Type /Catalog /Pages ${pagesId} 0 R`;
  if (isTagged && structTreeRootId) {
    catDict += ` /MarkInfo << /Marked true >> /StructTreeRoot ${structTreeRootId} 0 R`;
  }
  catDict += ` >>\nendobj\n`;
  push(`${catalogId} 0 obj\n${catDict}`);

  // 2. Pages Object
  record();
  const kids = pageAllocs.map((p) => `${p.pageId} 0 R`).join(" ");
  push(`${pagesId} 0 obj\n<< /Type /Pages /Kids [ ${kids} ] /Count ${shots.length} >>\nendobj\n`);

  // 3. StructTreeRoot Object (if tagged)
  if (isTagged && structTreeRootId) {
    record();
    const allStructKids = pageAllocs
      .flatMap((p) => p.structElemIds.map((id) => `${id} 0 R`))
      .join(" ");
    push(
      `${structTreeRootId} 0 obj\n<< /Type /StructTreeRoot /RoleMap << /H1 /H /H2 /H /H3 /H /H4 /H /P /P /Table /Table >> /K [ ${allStructKids} ] >>\nendobj\n`,
    );
  }

  // 4. Shared Font Dictionaries
  const type1Fonts: [string, string][] = [
    ["F1", "Helvetica"],
    ["F2", "Helvetica-Bold"],
    ["F3", "Helvetica-Oblique"],
    ["F4", "Helvetica-BoldOblique"],
    ["F5", "Times-Roman"],
    ["F6", "Times-Bold"],
    ["F7", "Times-Italic"],
    ["F8", "Times-BoldItalic"],
    ["F9", "Courier"],
    ["F10", "Courier-Bold"],
  ];
  for (const [key, baseFont] of type1Fonts) {
    record();
    push(
      `${fontIds[key]} 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /${baseFont} /Encoding /WinAnsiEncoding >>\nendobj\n`,
    );
  }

  // Type 0 Unicode Font
  record();
  push(
    `${fontIds.F_Uni} 0 obj\n<< /Type /Font /Subtype /Type0 /BaseFont /Helvetica /Encoding /Identity-H /DescendantFonts [ ${cidFontId} 0 R ] /ToUnicode ${toUnicodeId} 0 R >>\nendobj\n`,
  );

  // Descendant CIDFont
  record();
  push(
    `${cidFontId} 0 obj\n<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Helvetica /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${fontDescId} 0 R /DW 1000 >>\nendobj\n`,
  );

  // FontDescriptor
  record();
  push(
    `${fontDescId} 0 obj\n<< /Type /FontDescriptor /FontName /Helvetica /Flags 4 /FontBBox [ -1000 -1000 2000 2000 ] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 800 /StemV 80 >>\nendobj\n`,
  );

  // ToUnicode CMap Stream
  const toUnicodeCMap =
    `/CIDInit /ProcSet findresource begin\n` +
    `12 dict begin\n` +
    `begincmap\n` +
    `/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n` +
    `/CMapName /Custom-ToUnicode def\n` +
    `/CMapType 2 def\n` +
    `1 begincodespacerange\n` +
    `<0000> <FFFF>\n` +
    `endcodespacerange\n` +
    `1 beginbfrange\n` +
    `<0000> <FFFF> <0000>\n` +
    `endbfrange\n` +
    `endcmap\n` +
    `CMapName currentdict /CMap defineresource pop\n` +
    `end\n` +
    `end\n`;

  record();
  push(
    `${toUnicodeId} 0 obj\n<< /Length ${toUnicodeCMap.length} >>\nstream\n${toUnicodeCMap}\nendstream\nendobj\n`,
  );

  // 5. Document Information Object (/Info)
  const title = options?.metadata?.title ?? "Document";
  const author = options?.metadata?.author ?? "Docen";
  const creator = options?.metadata?.creator ?? "Docen Word Processor";
  const producer = options?.metadata?.producer ?? "Docen PDF Engine";
  const dateStr = formatPdfDate();

  record();
  push(
    `${infoId} 0 obj\n<< /Title (${escapePdfString(title)}) /Author (${escapePdfString(author)}) /Creator (${escapePdfString(creator)}) /Producer (${escapePdfString(producer)}) /CreationDate (${dateStr}) >>\nendobj\n`,
  );

  // 6. Per-Page Objects
  const fontEntries = Object.entries(fontIds)
    .map(([k, id]) => `/${k} ${id} 0 R`)
    .join(" ");

  for (const [i, page] of jpegs.entries()) {
    const shot = shots[i]!;
    const alloc = pageAllocs[i]!;
    const w = pt(shot.width).toFixed(2);
    const h = pt(shot.height).toFixed(2);

    // Page object
    let pageDict =
      `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [ 0 0 ${w} ${h} ] ` +
      `/Resources << /XObject << /Im0 ${alloc.imageId} 0 R >> /Font << ${fontEntries} >> >> ` +
      `/Contents ${alloc.contentId} 0 R`;

    if (alloc.annotIds.length > 0) {
      const annots = alloc.annotIds.map((id) => `${id} 0 R`).join(" ");
      pageDict += ` /Annots [ ${annots} ]`;
    }
    if (isTagged) {
      pageDict += ` /StructParents ${i}`;
    }
    pageDict += ` >>\nendobj\n`;

    record();
    push(`${alloc.pageId} 0 obj\n${pageDict}`);

    // Contents Stream
    let content = `q ${w} 0 0 ${h} 0 0 cm /Im0 Do Q\n`;
    const spans = shot.textSpans ?? [];
    if (spans.length > 0) {
      if (isTagged && alloc.structElemIds.length > 0) {
        content += `/P << /MCID 0 >> BDC\n`;
      }
      content += `BT\n3 Tr\n`;
      let currentFont = "";
      let currentSize = "";
      for (let sIdx = 0; sIdx < spans.length; sIdx++) {
        const span = spans[sIdx]!;
        if (!span.text) continue;
        const { fontName, isUnicode } = fontForSpan(span);
        const sizeStr = span.fontSize.toFixed(2);
        if (currentFont !== fontName || currentSize !== sizeStr) {
          content += `/${fontName} ${sizeStr} Tf\n`;
          currentFont = fontName;
          currentSize = sizeStr;
        }
        content += `1 0 0 1 ${span.x.toFixed(2)} ${span.y.toFixed(2)} Tm\n`;

        // Horizontal scaling Tz to match rendered word bounding box
        const estCharWidth = isUnicode ? span.fontSize : span.fontSize * 0.52;
        const estWidth = estCharWidth * span.text.length;
        if (estWidth > 0 && span.width > 0) {
          const scale = (span.width / estWidth) * 100;
          if (scale >= 30 && scale <= 400 && Math.abs(scale - 100) > 2) {
            content += `${scale.toFixed(1)} Tz\n`;
          } else {
            content += `100 Tz\n`;
          }
        }

        if (isUnicode) {
          content += `<${encodeHexUtf16(span.text)}> Tj\n`;
        } else {
          content += `(${escapePdfString(span.text)}) Tj\n`;
        }
      }
      content += `ET\n`;
      if (isTagged && alloc.structElemIds.length > 0) {
        content += `EMC\n`;
      }
    }

    record();
    push(
      `${alloc.contentId} 0 obj\n<< /Length ${content.length} >>\nstream\n${content}\nendstream\nendobj\n`,
    );

    // Image XObject
    record();
    push(
      `${alloc.imageId} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${page.width} /Height ${page.height} ` +
        `/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${page.jpeg.length} >>\nstream\n`,
    );
    push(page.jpeg);
    push("\nendstream\nendobj\n");

    // Link Annotations
    for (let lIdx = 0; lIdx < (shot.links ?? []).length; lIdx++) {
      const link = shot.links![lIdx]!;
      const annotId = alloc.annotIds[lIdx]!;
      const [x1, y1, x2, y2] = link.rect;
      record();
      let annotObj =
        `${annotId} 0 obj\n<< /Type /Annot /Subtype /Link ` +
        `/Rect [ ${x1.toFixed(2)} ${y1.toFixed(2)} ${x2.toFixed(2)} ${y2.toFixed(2)} ] ` +
        `/Border [ 0 0 0 ] /F 4`;
      if (link.url.startsWith("#")) {
        annotObj += ` /Dest (${escapePdfString(link.url.slice(1))})`;
      } else {
        annotObj += ` /A << /Type /Action /S /URI /URI (${escapePdfString(link.url)}) >>`;
      }
      annotObj += ` >>\nendobj\n`;
      push(annotObj);
    }

    // Structure Element (if tagged)
    if (isTagged && alloc.structElemIds.length > 0 && structTreeRootId) {
      const structElemId = alloc.structElemIds[0]!;
      record();
      push(
        `${structElemId} 0 obj\n<< /Type /StructElem /S /P /P ${structTreeRootId} 0 R /Pg ${alloc.pageId} 0 R /K 0 >>\nendobj\n`,
      );
    }
  }

  // 7. Xref Table & Trailer
  const xrefStart = offset();
  const size = offsets.length;
  let xref = `xref\n0 ${size}\n0000000000 65535 f \n`;
  for (let i = 1; i < size; i++) xref += `${xrefAt(offsets[i]!)} 00000 n \n`;
  xref += `trailer\n<< /Size ${size} /Root ${catalogId} 0 R /Info ${infoId} 0 R >>\nstartxref\n${xrefStart}\n%%EOF\n`;
  push(xref);

  return new Blob([bytes()], { type: "application/pdf" });
}

/** Extract text spans and link annotations from laid-out document pages for PDF export. */
export function extractPdfPageLayers(
  pages: readonly FlowPage[],
  sections: readonly CanvasStageSection[],
  sectionOfPage: readonly number[],
): { textSpans: PdfTextSpan[]; links: PdfLinkAnnotation[] }[] {
  const toPt = (px: number): number => (px * 72) / 96;

  return pages.map((page, pageIndex) => {
    const secIndex = sectionOfPage[pageIndex] ?? 0;
    const section = sections[secIndex] ?? sections[0];
    const flow = section?.flow ?? {
      pageWidthPx: 816,
      pageHeightPx: 1056,
      contentLeftPx: 96,
      contentTopPx: 96,
      contentWidthPx: 624,
      contentHeightPx: 864,
    };
    const pageH = flow.pageHeightPx;
    const toPdfY = (pxY: number): number => ((pageH - pxY) * 72) / 96;

    const textSpans: PdfTextSpan[] = [];
    const links: PdfLinkAnnotation[] = [];

    const pageCtx = {
      pageIndex,
      pageCount: pages.length,
      pageNumber: section?.pageNumbering
        ? {
            fmt: section.pageNumbering.format,
            offset: (section.pageNumbering.start ?? 1) - 1,
          }
        : undefined,
    };

    const walkBlock = (block: LaidOutBlock, originX: number, originY: number): void => {
      switch (block.kind) {
        case "paragraph": {
          const para = block;
          for (const line of para.lines) {
            const lineY = originY + line.yPx;
            const lineX =
              originX +
              (para.indent?.leftPx ?? 0) +
              (line.firstLineIndentPx ?? 0) +
              (line.xOffsetPx ?? 0);

            for (const item of line.items) {
              if (item.kind !== "text") continue;
              const inline = para.inline[item.inlineIndex];
              if (inline && "suppressed" in inline && inline.suppressed) continue;

              const measured = item.displayText ?? item.text;
              const text =
                inline && "field" in inline && inline.field
                  ? fieldLabelOf(inline, pageCtx, measured)
                  : measured;
              if (!text) continue;

              const sizePx =
                item.fontSizePx ??
                (inline && "style" in inline ? vertAlignedSizePx(inline.style) : 16);
              const baselineDepth = lineBaselineDepthPx(line, sizePx);
              const shiftPx =
                inline && "style" in inline
                  ? vertAlignBaselineShiftPx(inline.style) + (inline.style.baselineShiftPx ?? 0)
                  : 0;
              const baselineYPx = lineY + baselineDepth + (item.rubyLiftPx ?? 0) + shiftPx;

              const xPt = toPt(lineX + item.xPx);
              const yPt = toPdfY(baselineYPx);
              const widthPt = toPt(item.widthPx);
              const heightPt = toPt(line.heightPx);
              const fontSizePt = toPt(sizePx);

              const style = inline && "style" in inline ? inline.style : undefined;
              const fontFamily =
                typeof style?.family === "string"
                  ? style.family
                  : (style?.family?.latin ?? style?.family?.eastAsia);

              textSpans.push({
                text,
                x: xPt,
                y: yPt,
                width: widthPt,
                height: heightPt,
                fontSize: fontSizePt,
                fontFamily,
                bold: style?.bold,
                italic: style?.italic,
                tag: "P",
              });

              // Hyperlinks
              const link =
                inline && "link" in inline && inline.link
                  ? (inline.link as { url?: string; anchor?: string; tooltip?: string })
                  : undefined;
              const target = link?.url ?? (link?.anchor ? `#${link.anchor}` : undefined);
              if (target) {
                links.push({
                  rect: [xPt, toPdfY(lineY + line.heightPx), xPt + widthPt, toPdfY(lineY)],
                  url: target,
                  title: link?.tooltip,
                });
              }
            }
          }
          break;
        }
        case "table": {
          let rowY = originY;
          for (const row of block.rows) {
            let colX = originX + (block.offsetXPx ?? 0);
            let colIndex = 0;
            for (const cell of row.cells) {
              const cellLeft = colX + (cell.insets.left ?? 0);
              const cellTop = rowY + (cell.insets.top ?? 0) + (cell.contentOffsetYPx ?? 0);
              for (const stackItem of cell.stack) {
                walkBlock(stackItem.block, cellLeft, cellTop + stackItem.yPx);
              }
              const spannedWidth = block.columnWidthsPx
                .slice(colIndex, colIndex + cell.colspan)
                .reduce((a, b) => a + b, 0);
              colX += spannedWidth;
              colIndex += cell.colspan;
            }
            rowY += row.heightPx;
          }
          break;
        }
        case "group": {
          for (const child of block.children) {
            walkBlock(child.block, originX, originY + child.yPx);
          }
          break;
        }
      }
    };

    // 1. Page items
    for (const item of page.items) {
      walkBlock(item.block, flow.contentLeftPx + (item.xPx ?? 0), flow.contentTopPx + item.yPx);
    }

    // 2. Footnotes & Endnotes
    if (page.footnotes) {
      const footY = flow.contentTopPx + page.footnotes.yPx;
      for (const item of page.footnotes.items) {
        walkBlock(item.block, flow.contentLeftPx, footY + item.yPx);
      }
    }
    if (page.endnotes) {
      const endY = flow.contentTopPx + page.endnotes.yPx;
      for (const item of page.endnotes.items) {
        walkBlock(item.block, flow.contentLeftPx, endY + item.yPx);
      }
    }

    // 3. Section furniture (header / footer)
    const laid = section?.furnitureLaid;
    if (laid) {
      const slotIndex =
        pageIndex === 0 && laid.header[1] ? 1 : pageIndex % 2 === 1 && laid.header[2] ? 2 : 0;
      const header = laid.header[slotIndex];
      if (header) {
        const headTop = section.furniture?.headerDistancePx ?? 48;
        for (const item of header.stack) {
          walkBlock(item.block, flow.contentLeftPx, headTop + item.yPx);
        }
      }
      const footer = laid.footer[slotIndex];
      if (footer) {
        const footBottom = flow.pageHeightPx - (section.furniture?.footerDistancePx ?? 48);
        for (const item of footer.stack) {
          walkBlock(item.block, flow.contentLeftPx, footBottom - footer.heightPx + item.yPx);
        }
      }
    }

    return { textSpans, links };
  });
}
