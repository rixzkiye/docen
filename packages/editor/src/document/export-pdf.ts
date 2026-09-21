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
import {
  generateToUnicodeCMap,
  isFontEmbeddingAllowed,
  isFontSubsettingAllowed,
  readFontFsType,
  readFontUnitsPerEm,
  readGlyphAdvances,
} from "@docen/shaping";
import type { SubsetPlan } from "@docen/shaping/subsetter";

import type { CanvasStageSection } from "./canvas/stage";
import {
  extGStateDict,
  planSceneContent,
  standardFontFor,
  wrapSceneContent,
  type PdfSceneImageData,
  type PdfSceneNode,
  type PdfScenePage,
  type PdfSceneTextNode,
  type PdfTextFontRef,
} from "./pdf-scene";

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
  letterSpacing?: number;
  /** Run color (OOXML hex without '#') — the visible text mode's paint. */
  color?: string;
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
  /** The page's serialized vector scene (sceneSnapshots). When present the
   *  page paints as a vector content stream; absent falls back to the legacy
   *  JPEG page image. */
  scene?: PdfScenePage;
  textSpans?: PdfTextSpan[];
  links?: PdfLinkAnnotation[];
  formFields?: readonly PdfFormField[];
}

/** One embedded subset font with its raw bytes and optional mappings.
 *
 *  Text spans are encoded as Identity-H with the UTF-16 code unit as the CID.
 *  `cidToGid` maps each code unit to the subset's glyph ID (the PDF writer
 *  emits it as a `/CIDToGIDMap` stream) and `toUnicodeMap` maps code units to
 *  Unicode for extraction; without either, the identity ToUnicode CMap is
 *  used and CIDs resolve through Identity-H substitution. */
export interface PdfEmbeddedFont {
  /** Base font name written into the PDF (spaces stripped). */
  readonly fontName: string;
  /** CSS family used to route text spans to this face (substring match,
   *  case-insensitive); defaults to `fontName`. */
  readonly fontFamily?: string;
  readonly fontData: Uint8Array;
  /** CID (UTF-16 code unit) → glyph ID in the embedded subset. */
  readonly cidToGid?: ReadonlyMap<number, number>;
  /** Per-glyph horizontal advances (1000/em, index = glyph ID in fontData) —
   *  emitted as the CIDFont's /W array so visible embedded-font text lays out
   *  at the face's true widths. */
  readonly glyphAdvances?: readonly number[];
  readonly toUnicodeMap?: Map<number, number | string> | [number, number][];
  /** Font descriptor values (font units scaled to 1000/em by the writer). */
  readonly unitsPerEm?: number;
  readonly ascent?: number;
  readonly descent?: number;
  readonly bbox?: readonly [number, number, number, number];
  /** OS/2 fsType the embedding decision was based on (diagnostics). */
  readonly fsType?: number;
}

/** A host-registered font the PDF exporter may embed. */
export interface PdfEmbeddableFontSource {
  readonly family: string;
  readonly fontData: Uint8Array;
  /** OS/2.fsType override; read from the bytes when omitted. */
  readonly fsType?: number;
}

/** One outline item (bookmark) in the document outline tree. */
export interface PdfOutlineItem {
  title: string;
  /** Destination: 0-based page index or target with optional top in pt. */
  dest: number | { pageIndex: number; top?: number };
  children?: readonly PdfOutlineItem[];
}

/** Page label range specification. */
export interface PdfPageLabelRange {
  startPageIndex: number;
  style?: "decimal" | "romanUpper" | "romanLower" | "alphaUpper" | "alphaLower" | "none";
  prefix?: string;
  startNumber?: number;
}

/** Named destination target in PDF points. */
export interface PdfDestination {
  pageIndex: number;
  x?: number;
  y?: number;
  zoom?: number;
}

/** Standard PDF viewer preferences. */
export interface PdfViewerPreferences {
  displayDocTitle?: boolean;
  hideToolbar?: boolean;
  hideMenubar?: boolean;
  centerWindow?: boolean;
  fitWindow?: boolean;
}

/** Interactive form field for AcroForm. */
export interface PdfFormField {
  name: string;
  type: "text" | "checkbox" | "dropdown";
  pageIndex: number;
  rect: [number, number, number, number]; // [left, bottom, right, top] in pt
  value?: string | boolean;
  readOnly?: boolean;
  options?: readonly string[];
}

/** Accessibility structure element with semantic tag and alt text. */
export interface PdfStructElement {
  type:
    | "Document"
    | "Part"
    | "Art"
    | "Sect"
    | "Div"
    | "H1"
    | "H2"
    | "H3"
    | "P"
    | "Table"
    | "Figure";
  pageIndex: number;
  altText?: string;
  title?: string;
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
    creationDate?: Date | string;
    modDate?: Date | string;
  };
  /** Produce a Tagged PDF with /MarkInfo and /StructTreeRoot (default: true). */
  tagged?: boolean;
  /** Embedded subset TrueType/OpenType fonts for visual fidelity and text extraction. */
  embeddedFonts?: readonly PdfEmbeddedFont[];
  /** How visible text paints in a vector scene page:
   *  - "outlines" (default): glyph outlines are drawn as vector paths and the
   *    text layer stays invisible (exact visual fidelity, font-independent).
   *  - "embedded": glyph outlines are dropped and the text layer renders
   *    visibly through the embedded/standard fonts (smaller files; used for
   *    the size/fidelity measurement — see the R10-P1 report). */
  textMode?: "outlines" | "embedded";
  /** Document outline bookmarks hierarchy. */
  outline?: readonly PdfOutlineItem[];
  /** Page labels for section page numbering. */
  pageLabels?: readonly PdfPageLabelRange[];
  /** Named destination targets (#anchor or external reference). */
  destinations?: Record<string, number | PdfDestination>;
  /** Viewer preferences for window display. */
  viewerPreferences?: PdfViewerPreferences;
  /** Interactive form fields for AcroForm. */
  formFields?: readonly PdfFormField[];
  /** Semantic structure elements with alt text for figures and tables. */
  structElements?: readonly PdfStructElement[];
}

/** Default textMode for client-side editor export (visual vector fidelity). */
export const DEFAULT_EDITOR_TEXT_MODE = "outlines" as const;

/** Default textMode for server-side headless export (compact stream, embedded fonts). */
export const DEFAULT_SERVER_TEXT_MODE = "embedded" as const;

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

/** Escape special characters for PDF name literals (/name). */
export function escapePdfName(name: string): string {
  return name.replace(
    /[^a-zA-Z0-9_-]/g,
    (c) => `#${c.charCodeAt(0).toString(16).padStart(2, "0")}`,
  );
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
  return {
    fontName: standardFontFor(span.fontFamily, span.bold, span.italic),
    isUnicode: false,
  };
}

/** Build embedded PDF fonts from the host's registered font bytes.
 *
 *  Only fonts whose OS/2 fsType allows embedding are returned; subsetting is
 *  applied when the license also allows it (otherwise the full font is
 *  embedded). Each entry carries a `cidToGid` map so the writer can emit a
 *  `/CIDToGIDMap` and text spans can keep UTF-16 code units as CIDs. */
export async function buildEmbeddedPdfFonts(
  spans: readonly PdfTextSpan[],
  sources: Iterable<PdfEmbeddableFontSource>,
): Promise<PdfEmbeddedFont[]> {
  const byFamily = new Map<string, PdfEmbeddableFontSource>();
  for (const source of sources) {
    byFamily.set(source.family.trim().toLowerCase(), source);
  }
  if (byFamily.size === 0) return [];

  // The subsetter is export-path-only: loaded on demand so the runtime
  // shaping bundle never carries it (see @docen/shaping/subsetter).
  const { readCmap, subsetFontWithPlan } = await import("@docen/shaping/subsetter");

  const used = new Map<string, { source: PdfEmbeddableFontSource; codeUnits: Set<number> }>();
  for (const span of spans) {
    const family = span.fontFamily?.trim().toLowerCase();
    if (!family) continue;
    let source = byFamily.get(family);
    if (!source) {
      for (const [key, candidate] of byFamily) {
        if (family.includes(key) || key.includes(family)) {
          source = candidate;
          break;
        }
      }
    }
    if (!source) continue;
    const entry = used.get(source.family) ?? { source, codeUnits: new Set<number>() };
    for (let i = 0; i < span.text.length; i++) entry.codeUnits.add(span.text.charCodeAt(i));
    used.set(source.family, entry);
  }

  const fonts: PdfEmbeddedFont[] = [];
  for (const { source, codeUnits } of used.values()) {
    const fsType = source.fsType ?? readFontFsType(source.fontData);
    if (!isFontEmbeddingAllowed(fsType)) continue;
    const canSubset = isFontSubsettingAllowed(fsType);

    const cmap = readCmap(source.fontData);
    const usedGids = new Set<number>([0]);
    for (const cu of codeUnits) {
      const gid = cmap.get(cu);
      if (gid !== undefined) usedGids.add(gid);
    }

    const cidToGid = new Map<number, number>();
    let fontData = source.fontData;
    let subsetted = false;
    if (canSubset) {
      // A malformed font must never break the export: fall back to the full
      // font when subsetting throws.
      let plan: SubsetPlan | undefined;
      try {
        plan = subsetFontWithPlan(source.fontData, [...usedGids]);
      } catch {
        plan = undefined;
      }
      if (plan) {
        fontData = plan.data;
        subsetted = true;
        for (const cu of codeUnits) {
          const gid = cmap.get(cu);
          const newGid = gid === undefined ? undefined : plan.glyphMap.get(gid);
          if (newGid !== undefined) cidToGid.set(cu, newGid);
        }
      }
    }
    if (!subsetted) {
      for (const cu of codeUnits) {
        const gid = cmap.get(cu);
        if (gid !== undefined) cidToGid.set(cu, gid);
      }
    }

    const cleanName = source.family.replace(/[^A-Za-z0-9-]/g, "") || "Embedded";
    fonts.push({
      fontName: `${cleanName}${subsetted ? "Subset" : ""}`,
      fontFamily: source.family,
      fontData,
      cidToGid,
      glyphAdvances: readGlyphAdvances(fontData),
      unitsPerEm: readFontUnitsPerEm(source.fontData),
      fsType,
    });
  }
  return fonts;
}

/** Flate-encode a stream (zlib wrapper, PDF /FlateDecode); environments
 *  without CompressionStream get the raw bytes and no filter. */
async function flateEncode(data: Uint8Array): Promise<{ data: Uint8Array; filter: string }> {
  if (typeof CompressionStream === "undefined" || typeof Response === "undefined") {
    return { data, filter: "" };
  }
  try {
    const stream = new Blob([data as unknown as BlobPart])
      .stream()
      .pipeThrough(new CompressionStream("deflate"));
    const encoded = new Uint8Array(await new Response(stream).arrayBuffer());
    return { data: encoded, filter: " /Filter /FlateDecode" };
  } catch {
    return { data, filter: "" };
  }
}

/** The identity ToUnicode CMap used when the embedded font's CIDs are UTF-16
 *  code units (extraction maps the CID straight back to its code point). */
function identityToUnicodeCMap(): string {
  return (
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
    `end\n`
  );
}

/** Build the PDF blob from page snapshots, text layers, and link annotations.
 *  Pages keep their paper size (CSS px → pt at 72/96). A page with a `scene`
 *  paints as a vector content stream (paths/images/text); a page without one
 *  falls back to the legacy JPEG page image. The invisible text layer, link
 *  annotations and structure tags overlay either. Scene content streams are
 *  Flate-encoded; the text layer stays uncompressed so extraction stays
 *  debuggable.
 *
 *  Every object is allocated an ID up front and then written in ascending
 *  object-number order: the xref table is indexed by object number, so the
 *  write order and the allocation order must never diverge. */
export async function pagesToPdf(
  shots: readonly PdfPageShot[],
  options?: PdfExportOptions,
): Promise<Blob> {
  const { push, offset, bytes } = pdfWriter();
  // %PDF-1.4 plus binary-marker comment line (raw bytes)
  push("%PDF-1.4\n");
  push(new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]));

  interface PdfObject {
    id: number;
    parts: (string | Uint8Array)[];
  }
  const objects: PdfObject[] = [];
  const addObject = (id: number, ...parts: (string | Uint8Array)[]): void => {
    objects.push({ id, parts });
  };

  const pt = (px: number): number => (px * 72) / 96;

  const formFields =
    options?.formFields && options.formFields.length > 0
      ? options.formFields
      : shots.flatMap((s) => s.formFields ?? []);

  // Object numbering allocator
  let nextId = 1;
  const allocId = () => nextId++;

  const catalogId = allocId();
  const pagesId = allocId();
  const isTagged = options?.tagged !== false;
  const structTreeRootId = isTagged ? allocId() : undefined;

  // Standard Type 1 fonts + universal non-embedded Unicode Type 0 font
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
  const uniCidFontId = allocId();
  const uniFontDescId = allocId();
  const uniToUnicodeId = allocId();

  interface EmbeddedFontPlan {
    readonly index: number;
    readonly font: PdfEmbeddedFont;
    readonly resourceName: string;
    readonly type0Id: number;
    readonly cidFontId: number;
    readonly fontDescId: number;
    readonly toUnicodeId: number;
    readonly fontFileId: number;
    readonly cidToGidId?: number;
  }
  const embeddedFonts: EmbeddedFontPlan[] = (options?.embeddedFonts ?? []).map((font, index) => ({
    index,
    font,
    resourceName: `F_Emb${index}`,
    type0Id: allocId(),
    cidFontId: allocId(),
    fontDescId: allocId(),
    toUnicodeId: allocId(),
    fontFileId: allocId(),
    ...(font.cidToGid ? { cidToGidId: allocId() } : {}),
  }));

  const infoId = allocId();

  // Vector scene plans — pure analysis (resource names are page-local) done
  // before any object is written, so every allocation below is final.
  const textMode = options?.textMode ?? DEFAULT_EDITOR_TEXT_MODE;
  // The face for a family: an exact family match wins, then the longest
  // substring match (so "Calibri" never resolves to "Calibri Light").
  const embeddedFontForFamily = (family: string | undefined): EmbeddedFontPlan | undefined => {
    const key = (family ?? "").trim().toLowerCase();
    if (!key || embeddedFonts.length === 0) return undefined;
    for (const emb of embeddedFonts) {
      const target = (emb.font.fontFamily ?? emb.font.fontName).trim().toLowerCase();
      if (target && target === key) return emb;
    }
    let best: { plan: EmbeddedFontPlan; length: number } | undefined;
    for (const emb of embeddedFonts) {
      const target = (emb.font.fontFamily ?? emb.font.fontName).trim().toLowerCase();
      if (!target) continue;
      if (!key.includes(target) && !target.includes(key)) continue;
      if (!best || target.length > best.length) best = { plan: emb, length: target.length };
    }
    return best?.plan;
  };
  const sceneFontForText = (node: PdfSceneTextNode): PdfTextFontRef => {
    const embedded = embeddedFontForFamily(node.fontFamily);
    if (embedded) return { resource: embedded.resourceName, isUnicode: true };
    return {
      resource: standardFontFor(node.fontFamily, node.bold, node.italic),
      isUnicode: false,
    };
  };
  const scenePlans = shots.map((shot) =>
    shot.scene
      ? planSceneContent(shot.scene, {
          fontForText: sceneFontForText,
          includeGlyphs: textMode !== "embedded",
        })
      : null,
  );
  // In outlines mode the scene renders fallback Text visibly; a layout span
  // for the same run must leave the invisible layer or extraction would return
  // it twice (see consumeVisiblePlacement).
  const pageTextSpans: PdfTextSpan[][] = shots.map((shot, index) => {
    const spans = shot.textSpans ?? [];
    if (!shot.scene || !scenePlans[index]) return spans;
    const placements = sceneTextPlacements(shot.scene);
    if (placements.length === 0) return spans;
    return spans.filter((span) => !consumeVisiblePlacement(placements, span, shot.scene!.height));
  });

  // Image XObjects dedupe globally by content key; an SMask is allocated only
  // when the raw RGBA really carries alpha.
  interface SceneImageAlloc {
    id: number;
    data: PdfSceneImageData;
    smaskId?: number;
  }
  const imageAllocs = new Map<string, SceneImageAlloc>();
  for (const plan of scenePlans) {
    if (!plan) continue;
    for (const image of plan.images) {
      if (imageAllocs.has(image.key)) continue;
      imageAllocs.set(image.key, {
        id: allocId(),
        data: image,
        ...(image.rgba && image.hasAlpha ? { smaskId: allocId() } : {}),
      });
    }
  }

  interface PageAlloc {
    pageId: number;
    /** The legacy whole-page JPEG (pages without a scene only). */
    imageId?: number;
    /** The text (or legacy) content stream. */
    contentId: number;
    /** The vector scene content stream (scene pages only). */
    sceneContentId?: number;
    /** Per-page image resources, /Im0… in plan order. */
    imageRefs: { name: string; id: number }[];
    /** Per-page ExtGState object ids, /GS0… in plan order. */
    stateIds: number[];
    annotIds: number[];
    fieldAnnotIds: number[];
    structElemIds: number[];
  }

  const pageAllocs: PageAlloc[] = [];
  for (const [index, shot] of shots.entries()) {
    const plan = scenePlans[index];
    const pageId = allocId();
    const contentId = allocId();
    const sceneContentId = plan ? allocId() : undefined;
    const imageId = plan ? undefined : allocId();
    const imageRefs = (plan?.images ?? []).map((image, k) => ({
      name: `Im${k}`,
      id: imageAllocs.get(image.key)!.id,
    }));
    const stateIds = (plan?.extGStates ?? []).map(() => allocId());
    const annotIds = (shot.links ?? []).map(() => allocId());
    const pageFields = formFields.filter(
      (f) => Math.max(0, Math.min(shots.length - 1, f.pageIndex)) === index,
    );
    const fieldAnnotIds = pageFields.map(() => allocId());
    const structElemIds = isTagged && (pageTextSpans[index]?.length ?? 0) > 0 ? [allocId()] : [];
    pageAllocs.push({
      pageId,
      contentId,
      ...(sceneContentId !== undefined ? { sceneContentId } : {}),
      ...(imageId !== undefined ? { imageId } : {}),
      imageRefs,
      stateIds,
      annotIds,
      fieldAnnotIds,
      structElemIds,
    });
  }

  // Outline Tree (/Outlines)
  let outlineRootId: number | undefined;
  interface OutlineItemAlloc {
    id: number;
    item: PdfOutlineItem;
    parentId: number;
    prevId?: number;
    nextId?: number;
    firstChildId?: number;
    lastChildId?: number;
    count: number;
  }
  const outlineAllocs: OutlineItemAlloc[] = [];

  if (options?.outline && options.outline.length > 0) {
    outlineRootId = allocId();
    const buildOutlineLevel = (
      items: readonly PdfOutlineItem[],
      parentId: number,
    ): { firstId: number; lastId: number; totalCount: number } => {
      const levelAllocs: OutlineItemAlloc[] = items.map((item) => ({
        id: allocId(),
        item,
        parentId,
        count: 0,
      }));
      for (let i = 0; i < levelAllocs.length; i++) {
        if (i > 0) levelAllocs[i]!.prevId = levelAllocs[i - 1]!.id;
        if (i < levelAllocs.length - 1) levelAllocs[i]!.nextId = levelAllocs[i + 1]!.id;
      }
      let totalCount = 0;
      for (const node of levelAllocs) {
        if (node.item.children && node.item.children.length > 0) {
          const childLevel = buildOutlineLevel(node.item.children, node.id);
          node.firstChildId = childLevel.firstId;
          node.lastChildId = childLevel.lastId;
          node.count = childLevel.totalCount;
          totalCount += childLevel.totalCount;
        }
        totalCount += 1;
        outlineAllocs.push(node);
      }
      return {
        firstId: levelAllocs[0]!.id,
        lastId: levelAllocs[levelAllocs.length - 1]!.id,
        totalCount,
      };
    };

    const rootLevel = buildOutlineLevel(options.outline, outlineRootId);
    addObject(
      outlineRootId,
      `${outlineRootId} 0 obj\n<< /Type /Outlines /First ${rootLevel.firstId} 0 R /Last ${rootLevel.lastId} 0 R /Count ${rootLevel.totalCount} >>\nendobj\n`,
    );

    for (const node of outlineAllocs) {
      const pIdx = typeof node.item.dest === "number" ? node.item.dest : node.item.dest.pageIndex;
      const targetPageId = (pageAllocs[pIdx] ?? pageAllocs[0]!).pageId;
      const top =
        typeof node.item.dest === "object" && node.item.dest.top != null ? node.item.dest.top : 792;
      let dict =
        `<< /Title (${escapePdfString(node.item.title)}) /Parent ${node.parentId} 0 R ` +
        `/Dest [ ${targetPageId} 0 R /XYZ 0 ${top.toFixed(2)} 0 ]`;
      if (node.prevId) dict += ` /Prev ${node.prevId} 0 R`;
      if (node.nextId) dict += ` /Next ${node.nextId} 0 R`;
      if (node.firstChildId && node.lastChildId) {
        dict += ` /First ${node.firstChildId} 0 R /Last ${node.lastChildId} 0 R /Count ${node.count}`;
      }
      dict += ` >>\nendobj\n`;
      addObject(node.id, `${node.id} 0 obj\n${dict}`);
    }
  }

  // Page Labels (/PageLabels)
  let pageLabelsDict = "";
  if (options?.pageLabels && options.pageLabels.length > 0) {
    const sortedLabels = [...options.pageLabels].sort(
      (a, b) => a.startPageIndex - b.startPageIndex,
    );
    const numsEntries: string[] = [];
    for (const range of sortedLabels) {
      let dict = "<<";
      if (range.style && range.style !== "none") {
        const styleCode =
          {
            decimal: "/D",
            romanUpper: "/R",
            romanLower: "/r",
            alphaUpper: "/A",
            alphaLower: "/a",
          }[range.style] ?? "/D";
        dict += ` /S ${styleCode}`;
      }
      if (range.prefix) dict += ` /P (${escapePdfString(range.prefix)})`;
      if (range.startNumber != null) dict += ` /St ${range.startNumber}`;
      dict += " >>";
      numsEntries.push(`${range.startPageIndex} ${dict}`);
    }
    pageLabelsDict = ` /PageLabels << /Nums [ ${numsEntries.join(" ")} ] >>`;
  }

  // Named Destinations (/Dests)
  let destsDict = "";
  if (options?.destinations && Object.keys(options.destinations).length > 0) {
    const destEntries: string[] = [];
    for (const [name, target] of Object.entries(options.destinations)) {
      const pIdx = typeof target === "number" ? target : target.pageIndex;
      const targetPageId = (pageAllocs[pIdx] ?? pageAllocs[0]!).pageId;
      const x = typeof target === "object" && target.x != null ? target.x : 0;
      const y = typeof target === "object" && target.y != null ? target.y : 0;
      const zoom = typeof target === "object" && target.zoom != null ? target.zoom : 0;
      destEntries.push(
        `/${escapePdfName(name)} [ ${targetPageId} 0 R /XYZ ${x.toFixed(2)} ${y.toFixed(2)} ${zoom} ]`,
      );
    }
    if (destEntries.length > 0) {
      destsDict = ` /Dests << ${destEntries.join(" ")} >>`;
    }
  }

  // Viewer Preferences (/ViewerPreferences)
  let viewerPrefsDict = "";
  if (options?.viewerPreferences) {
    const prefs: string[] = [];
    const vp = options.viewerPreferences;
    if (vp.displayDocTitle != null) prefs.push(`/DisplayDocTitle ${vp.displayDocTitle}`);
    if (vp.hideToolbar != null) prefs.push(`/HideToolbar ${vp.hideToolbar}`);
    if (vp.hideMenubar != null) prefs.push(`/HideMenubar ${vp.hideMenubar}`);
    if (vp.centerWindow != null) prefs.push(`/CenterWindow ${vp.centerWindow}`);
    if (vp.fitWindow != null) prefs.push(`/FitWindow ${vp.fitWindow}`);
    if (prefs.length > 0) viewerPrefsDict = ` /ViewerPreferences << ${prefs.join(" ")} >>`;
  }

  // AcroForm object ID
  const acroFormId = formFields.length > 0 ? allocId() : undefined;

  // 1. Catalog Object
  let catDict = `<< /Type /Catalog /Pages ${pagesId} 0 R`;
  if (isTagged && structTreeRootId) {
    catDict += ` /MarkInfo << /Marked true >> /StructTreeRoot ${structTreeRootId} 0 R`;
  }
  if (outlineRootId !== undefined) {
    catDict += ` /Outlines ${outlineRootId} 0 R`;
  }
  if (pageLabelsDict) catDict += pageLabelsDict;
  if (destsDict) catDict += destsDict;
  if (viewerPrefsDict) catDict += viewerPrefsDict;
  if (acroFormId !== undefined) {
    catDict += ` /AcroForm ${acroFormId} 0 R`;
  }
  catDict += ` >>\nendobj\n`;
  addObject(catalogId, `${catalogId} 0 obj\n${catDict}`);

  // 2. Pages Object
  const kids = pageAllocs.map((p) => `${p.pageId} 0 R`).join(" ");
  addObject(
    pagesId,
    `${pagesId} 0 obj\n<< /Type /Pages /Kids [ ${kids} ] /Count ${shots.length} >>\nendobj\n`,
  );

  // Custom accessibility struct elements
  const customStructIds: number[] = [];
  if (isTagged && structTreeRootId && options?.structElements) {
    for (const elem of options.structElements) {
      const id = allocId();
      customStructIds.push(id);
      const targetPageId = (pageAllocs[elem.pageIndex] ?? pageAllocs[0]!).pageId;
      let dict = `${id} 0 obj\n<< /Type /StructElem /S /${elem.type} /P ${structTreeRootId} 0 R /Pg ${targetPageId} 0 R /K 0`;
      if (elem.altText) dict += ` /Alt (${escapePdfString(elem.altText)})`;
      if (elem.title) dict += ` /T (${escapePdfString(elem.title)})`;
      dict += ` >>\nendobj\n`;
      addObject(id, dict);
    }
  }

  // 3. StructTreeRoot Object (if tagged)
  if (isTagged && structTreeRootId) {
    const allStructKids = [
      ...pageAllocs.flatMap((p) => p.structElemIds.map((id) => `${id} 0 R`)),
      ...customStructIds.map((id) => `${id} 0 R`),
    ].join(" ");
    addObject(
      structTreeRootId,
      `${structTreeRootId} 0 obj\n<< /Type /StructTreeRoot /RoleMap << /H1 /H /H2 /H /H3 /H /H4 /H /P /P /Table /Table /Figure /Figure >> /K [ ${allStructKids} ] >>\nendobj\n`,
    );
  }

  // 3b. AcroForm Object
  if (acroFormId !== undefined) {
    const allFieldIds = pageAllocs.flatMap((p) => p.fieldAnnotIds);
    const fieldsRef = allFieldIds.map((id) => `${id} 0 R`).join(" ");
    addObject(
      acroFormId,
      `${acroFormId} 0 obj\n<< /Fields [ ${fieldsRef} ] /NeedAppearances true /DA (/F1 12 Tf 0 g) >>\nendobj\n`,
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
    addObject(
      fontIds[key]!,
      `${fontIds[key]} 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /${baseFont} /Encoding /WinAnsiEncoding >>\nendobj\n`,
    );
  }

  // Embedded subset fonts: FontFile2 stream + ToUnicode + CIDToGIDMap + Type0.
  for (const emb of embeddedFonts) {
    const font = emb.font;
    const baseFontName = font.fontName.replace(/\s+/g, "") || `EmbeddedFont${emb.index}`;
    const data = font.fontData;
    addObject(
      emb.fontFileId,
      `${emb.fontFileId} 0 obj\n<< /Length ${data.length} /Length1 ${data.length} >>\nstream\n`,
      data,
      "\nendstream\nendobj\n",
    );

    const toUnicodeCMap = font.toUnicodeMap
      ? generateToUnicodeCMap(font.toUnicodeMap, `Docen-ToUnicode-${emb.index}`)
      : identityToUnicodeCMap();
    addObject(
      emb.toUnicodeId,
      `${emb.toUnicodeId} 0 obj\n<< /Length ${toUnicodeCMap.length} >>\nstream\n${toUnicodeCMap}\nendstream\nendobj\n`,
    );

    let cidToGidRef = "";
    if (font.cidToGid && font.cidToGid.size > 0 && emb.cidToGidId) {
      let maxCid = 0;
      for (const cid of font.cidToGid.keys()) {
        if (cid > maxCid && cid <= 0xffff) maxCid = cid;
      }
      const raw = new Uint8Array((maxCid + 1) * 2);
      const rawView = new DataView(raw.buffer);
      for (const [cid, gid] of font.cidToGid) {
        if (cid <= 0xffff) rawView.setUint16(cid * 2, gid & 0xffff);
      }
      const stream = await flateEncode(raw);
      addObject(
        emb.cidToGidId,
        `${emb.cidToGidId} 0 obj\n<< /Length ${stream.data.length}${stream.filter} >>\nstream\n`,
        stream.data,
        "\nendstream\nendobj\n",
      );
      cidToGidRef = ` /CIDToGIDMap ${emb.cidToGidId} 0 R`;
    }

    const unitsPerEm = font.unitsPerEm && font.unitsPerEm > 0 ? font.unitsPerEm : 1000;
    const fontScale = 1000 / unitsPerEm;
    const bbox = font.bbox ?? [-1000, -1000, 2000, 2000];
    const ascent = Math.round((font.ascent ?? 0.8 * unitsPerEm) * fontScale);
    const descent = Math.round((font.descent ?? -0.2 * unitsPerEm) * fontScale);
    // /W widths: without them a CIDFontType2 viewer assumes the /DW default
    // (1000/em) for every CID — visible embedded text would render one em per
    // glyph. Group consecutive CIDs into `first [w…]` runs.
    let widthsRef = "";
    if (font.glyphAdvances && font.cidToGid && font.cidToGid.size > 0) {
      const entries = [...font.cidToGid].sort((a, b) => a[0] - b[0]);
      const segments: string[] = [];
      let runStart = -1;
      let runWidths: number[] = [];
      const flush = (): void => {
        if (runStart >= 0 && runWidths.length > 0) {
          segments.push(`${runStart} [ ${runWidths.join(" ")} ]`);
        }
        runStart = -1;
        runWidths = [];
      };
      for (const [cid, gid] of entries) {
        if (cid > 0xffff) continue;
        const width = font.glyphAdvances[gid];
        if (width === undefined) continue;
        if (runStart >= 0 && cid === runStart + runWidths.length) runWidths.push(width);
        else {
          flush();
          runStart = cid;
          runWidths = [width];
        }
      }
      flush();
      if (segments.length > 0) widthsRef = ` /W [ ${segments.join(" ")} ]`;
    }
    addObject(
      emb.type0Id,
      `${emb.type0Id} 0 obj\n<< /Type /Font /Subtype /Type0 /BaseFont /${baseFontName} /Encoding /Identity-H /DescendantFonts [ ${emb.cidFontId} 0 R ] /ToUnicode ${emb.toUnicodeId} 0 R >>\nendobj\n`,
    );
    addObject(
      emb.cidFontId,
      `${emb.cidFontId} 0 obj\n<< /Type /Font /Subtype /CIDFontType2 /BaseFont /${baseFontName} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${emb.fontDescId} 0 R /DW 1000${widthsRef}${cidToGidRef} >>\nendobj\n`,
    );
    addObject(
      emb.fontDescId,
      `${emb.fontDescId} 0 obj\n<< /Type /FontDescriptor /FontName /${baseFontName} /Flags 4 /FontBBox [ ${bbox.join(" ")} ] /ItalicAngle 0 /Ascent ${ascent} /Descent ${descent} /CapHeight ${ascent} /StemV 80 /FontFile2 ${emb.fontFileId} 0 R >>\nendobj\n`,
    );
  }

  // Universal Type 0 Unicode Font (non-embedded fallback)
  addObject(
    fontIds.F_Uni!,
    `${fontIds.F_Uni} 0 obj\n<< /Type /Font /Subtype /Type0 /BaseFont /Helvetica /Encoding /Identity-H /DescendantFonts [ ${uniCidFontId} 0 R ] /ToUnicode ${uniToUnicodeId} 0 R >>\nendobj\n`,
  );

  addObject(
    uniCidFontId,
    `${uniCidFontId} 0 obj\n<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Helvetica /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${uniFontDescId} 0 R /DW 1000 /W [ 0 255 500 ] >>\nendobj\n`,
  );

  addObject(
    uniFontDescId,
    `${uniFontDescId} 0 obj\n<< /Type /FontDescriptor /FontName /Helvetica /Flags 4 /FontBBox [ -1000 -1000 2000 2000 ] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 800 /StemV 80 >>\nendobj\n`,
  );

  const uniToUnicodeCMap = identityToUnicodeCMap();
  addObject(
    uniToUnicodeId,
    `${uniToUnicodeId} 0 obj\n<< /Length ${uniToUnicodeCMap.length} >>\nstream\n${uniToUnicodeCMap}\nendstream\nendobj\n`,
  );

  // 5. Document Information Object (/Info)
  const title = options?.metadata?.title ?? "Document";
  const author = options?.metadata?.author ?? "Docen";
  const creator = options?.metadata?.creator ?? "Docen Word Processor";
  const producer = options?.metadata?.producer ?? "Docen PDF Engine";
  const creationDateStr = options?.metadata?.creationDate
    ? options.metadata.creationDate instanceof Date
      ? formatPdfDate(options.metadata.creationDate)
      : options.metadata.creationDate
    : formatPdfDate();
  const modDateStr = options?.metadata?.modDate
    ? options.metadata.modDate instanceof Date
      ? formatPdfDate(options.metadata.modDate)
      : options.metadata.modDate
    : creationDateStr;

  let infoDict =
    `<< /Title (${escapePdfString(title)}) /Author (${escapePdfString(author)}) ` +
    `/Creator (${escapePdfString(creator)}) /Producer (${escapePdfString(producer)}) ` +
    `/CreationDate (${creationDateStr}) /ModDate (${modDateStr})`;
  if (options?.metadata?.subject) {
    infoDict += ` /Subject (${escapePdfString(options.metadata.subject)})`;
  }
  if (options?.metadata?.keywords) {
    infoDict += ` /Keywords (${escapePdfString(options.metadata.keywords)})`;
  }
  infoDict += ` >>\nendobj\n`;
  addObject(infoId, `${infoId} 0 obj\n${infoDict}`);

  // 6. Per-Page Objects
  const fontEntries = [
    ...Object.entries(fontIds).map(([k, id]) => `/${k} ${id} 0 R`),
    ...embeddedFonts.map((emb) => `/${emb.resourceName} ${emb.type0Id} 0 R`),
  ].join(" ");

  const embeddedFontForSpan = (span: PdfTextSpan): EmbeddedFontPlan | undefined => {
    if (embeddedFonts.length === 0) return undefined;
    return embeddedFontForFamily(span.fontFamily) ?? embeddedFonts[0];
  };

  for (const [i, shot] of shots.entries()) {
    const alloc = pageAllocs[i]!;
    const plan = scenePlans[i];
    const scene = shot.scene;
    const widthPx = scene?.width ?? shot.width;
    const heightPx = scene?.height ?? shot.height;
    const w = pt(widthPx).toFixed(2);
    const h = pt(heightPx).toFixed(2);

    // Page object — resources differ by path: vector pages list their images,
    // ExtGStates and fonts; legacy pages carry the page JPEG.
    let resources: string;
    if (plan) {
      const xobjects = alloc.imageRefs.map((ref) => `/${ref.name} ${ref.id} 0 R`).join(" ");
      const states = alloc.stateIds.map((id, k) => `/GS${k} ${id} 0 R`).join(" ");
      resources =
        `/Resources << /ProcSet [ /PDF /Text /ImageB /ImageC /ImageI ]` +
        ` /XObject << ${xobjects} >> /ExtGState << ${states} >> /Font << ${fontEntries} >> >>`;
    } else {
      resources = `/Resources << /XObject << /Im0 ${alloc.imageId} 0 R >> /Font << ${fontEntries} >> >>`;
    }
    const contents = plan
      ? `[ ${alloc.sceneContentId} 0 R ${alloc.contentId} 0 R ]`
      : `${alloc.contentId} 0 R`;
    let pageDict =
      `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [ 0 0 ${w} ${h} ] ` +
      `${resources} /Contents ${contents}`;

    const allAnnots = [...alloc.annotIds, ...alloc.fieldAnnotIds];
    if (allAnnots.length > 0) {
      const annots = allAnnots.map((id) => `${id} 0 R`).join(" ");
      pageDict += ` /Annots [ ${annots} ]`;
    }
    if (isTagged) {
      pageDict += ` /StructParents ${i}`;
    }
    pageDict += ` >>\nendobj\n`;

    addObject(alloc.pageId, `${alloc.pageId} 0 obj\n${pageDict}`);

    // Scene content stream — Flate-encoded (glyph outlines are verbose); the
    // pure planner already emitted page-px operators, wrap adds the page
    // transform (px → pt, y flip) and the artifact marker.
    if (plan && scene && alloc.sceneContentId !== undefined) {
      const body = wrapSceneContent(plan.content, heightPx);
      const stream = await flateEncode(new TextEncoder().encode(body));
      addObject(
        alloc.sceneContentId,
        `${alloc.sceneContentId} 0 obj\n<< /Length ${stream.data.length}${stream.filter} >>\nstream\n`,
        stream.data,
        "\nendstream\nendobj\n",
      );
    }

    // Text (or legacy image) content stream.
    const spans = pageTextSpans[i] ?? [];
    // Embedded font measurement mode renders the text layer visibly; the
    // default outlines mode keeps it invisible (3 Tr) for search/selection.
    const visibleText = plan !== null && textMode === "embedded";
    let content = plan ? "" : `q ${w} 0 0 ${h} 0 0 cm /Im0 Do Q\n`;
    if (spans.length > 0) {
      if (isTagged && alloc.structElemIds.length > 0) {
        content += `/P << /MCID 0 >> BDC\n`;
      }
      content += visibleText ? `BT\n` : `BT\n3 Tr\n`;
      let currentFont = "";
      let currentSize = "";
      for (let sIdx = 0; sIdx < spans.length; sIdx++) {
        const span = spans[sIdx]!;
        if (!span.text) continue;
        const embedded = embeddedFontForSpan(span);
        const selected = embedded
          ? { fontName: embedded.resourceName, isUnicode: true }
          : fontForSpan(span);
        const { fontName, isUnicode } = selected;
        const sizeStr = span.fontSize.toFixed(2);
        if (currentFont !== fontName || currentSize !== sizeStr) {
          content += `/${fontName} ${sizeStr} Tf\n`;
          currentFont = fontName;
          currentSize = sizeStr;
        }
        if (visibleText) {
          const hex = (span.color ?? "1b1b1b").replace(/^#/, "");
          const r = parseInt(hex.slice(0, 2), 16) / 255;
          const g = parseInt(hex.slice(2, 4), 16) / 255;
          const b = parseInt(hex.slice(4, 6), 16) / 255;
          content += `${r.toFixed(3)} ${g.toFixed(3)} ${b.toFixed(3)} rg\n`;
          if (span.bold) {
            content += `${r.toFixed(3)} ${g.toFixed(3)} ${b.toFixed(3)} RG\n`;
            const strokeWidth = (span.fontSize * 0.03).toFixed(3);
            content += `${strokeWidth} w\n2 Tr\n`;
          } else {
            content += `0 Tr\n`;
          }
        }
        if (span.letterSpacing) {
          content += `${span.letterSpacing.toFixed(3)} Tc\n`;
        } else {
          content += `0 Tc\n`;
        }
        const skew = span.italic ? "0.2126" : "0";
        // When rendering visible text in embedded mode, compensate for the subpixel
        // FreeType rasterizer baseline alignment difference (~0.28 pt at 96 DPI) vs
        // Canvas layout geometry.
        const yCoord = visibleText ? span.y - 0.28 : span.y;
        content += `1 0 ${skew} 1 ${span.x.toFixed(2)} ${yCoord.toFixed(2)} Tm\n`;

        // Horizontal scaling Tz to match rendered word bounding box
        let estWidth = 0;
        const letterSpacing = span.letterSpacing ?? 0;
        if (embedded?.font.glyphAdvances && embedded.font.cidToGid) {
          for (let i = 0; i < span.text.length; i++) {
            const code = span.text.charCodeAt(i);
            const gid = embedded.font.cidToGid.get(code) ?? 0;
            const adv = embedded.font.glyphAdvances[gid] ?? 500;
            estWidth += (adv / 1000) * span.fontSize + letterSpacing;
          }
        } else {
          const isCjk = /[\u3000-\u9fff\uac00-\ud7af]/.test(span.text);
          const estCharWidth = isCjk ? span.fontSize : span.fontSize * 0.52;
          estWidth = (estCharWidth + letterSpacing) * span.text.length;
        }
        if (estWidth > 0 && span.width > 0) {
          const scale = (span.width / estWidth) * 100;
          if (scale >= 30 && scale <= 400 && Math.abs(scale - 100) > 0.05) {
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

    // /Length counts the UTF-8 bytes of the content stream, not its JS chars.
    const contentLength = new TextEncoder().encode(content).length;
    addObject(
      alloc.contentId,
      `${alloc.contentId} 0 obj\n<< /Length ${contentLength} >>\nstream\n${content}\nendstream\nendobj\n`,
    );

    // Legacy whole-page image XObject (never produced by the vector path).
    if (!plan) {
      const page = await jpegOf(shot);
      addObject(
        alloc.imageId!,
        `${alloc.imageId} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${page.width} /Height ${page.height} ` +
          `/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${page.jpeg.length} >>\nstream\n`,
        page.jpeg,
        "\nendstream\nendobj\n",
      );
    }

    // ExtGStates (one per distinct alpha/blend combination on this page).
    if (plan) {
      for (let sIdx = 0; sIdx < plan.extGStates.length; sIdx++) {
        const stateId = alloc.stateIds[sIdx]!;
        addObject(stateId, `${stateId} 0 obj\n${extGStateDict(plan.extGStates[sIdx]!)}\nendobj\n`);
      }
    }

    // Link Annotations
    for (let lIdx = 0; lIdx < (shot.links ?? []).length; lIdx++) {
      const link = shot.links![lIdx]!;
      const annotId = alloc.annotIds[lIdx]!;
      const [x1, y1, x2, y2] = link.rect;
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
      addObject(annotId, annotObj);
    }

    // Form Field Annotations (AcroForm Widgets)
    const pageFields = formFields.filter(
      (f) => Math.max(0, Math.min(shots.length - 1, f.pageIndex)) === i,
    );
    for (let fIdx = 0; fIdx < pageFields.length; fIdx++) {
      const field = pageFields[fIdx]!;
      const fieldId = alloc.fieldAnnotIds[fIdx]!;
      const [x1, y1, x2, y2] = field.rect;
      let fieldDict =
        `${fieldId} 0 obj\n<< /Type /Annot /Subtype /Widget ` +
        `/P ${alloc.pageId} 0 R ` +
        `/Rect [ ${x1.toFixed(2)} ${y1.toFixed(2)} ${x2.toFixed(2)} ${y2.toFixed(2)} ] ` +
        `/T (${escapePdfString(field.name)}) /F 4`;

      if (field.readOnly) {
        fieldDict += ` /Ff 1`;
      }

      if (field.type === "text") {
        const val = typeof field.value === "string" ? field.value : "";
        fieldDict += ` /FT /Tx /V (${escapePdfString(val)}) /DA (/F1 12 Tf 0 g)`;
      } else if (field.type === "checkbox") {
        const isChecked = field.value === true;
        const state = isChecked ? "/Yes" : "/Off";
        fieldDict += ` /FT /Btn /V ${state} /AS ${state}`;
      } else if (field.type === "dropdown") {
        const opts = field.options?.map((o) => `(${escapePdfString(o)})`).join(" ") ?? "";
        fieldDict += ` /FT /Ch /Opt [ ${opts} ] /V (${escapePdfString(String(field.value ?? ""))}) /DA (/F1 12 Tf 0 g)`;
      }
      fieldDict += ` >>\nendobj\n`;
      addObject(fieldId, fieldDict);
    }

    // Structure Element (if tagged)
    if (isTagged && alloc.structElemIds.length > 0 && structTreeRootId) {
      const structElemId = alloc.structElemIds[0]!;
      addObject(
        structElemId,
        `${structElemId} 0 obj\n<< /Type /StructElem /S /P /P ${structTreeRootId} 0 R /Pg ${alloc.pageId} 0 R /K 0 >>\nendobj\n`,
      );
    }
  }

  // Scene image XObjects — shared across pages by content key. JPEGs pass
  // through untouched (/DCTDecode); raw RGBA streams Flate-encode, with a
  // /DeviceGray SMask for the alpha channel when the bitmap has one.
  for (const imageAlloc of imageAllocs.values()) {
    const data = imageAlloc.data;
    if (data.jpeg) {
      addObject(
        imageAlloc.id,
        `${imageAlloc.id} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${data.width} /Height ${data.height} ` +
          `/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${data.jpeg.length} >>\nstream\n`,
        data.jpeg,
        "\nendstream\nendobj\n",
      );
      continue;
    }
    if (!data.rgba) continue;
    let rgb = data.rgba;
    if (imageAlloc.smaskId !== undefined || data.hasAlpha) {
      rgb = new Uint8Array(data.width * data.height * 3);
      for (let p = 0; p < data.width * data.height; p++) {
        rgb[p * 3] = data.rgba[p * 4]!;
        rgb[p * 3 + 1] = data.rgba[p * 4 + 1]!;
        rgb[p * 3 + 2] = data.rgba[p * 4 + 2]!;
      }
    }
    const stream = await flateEncode(rgb);
    const smask = imageAlloc.smaskId !== undefined ? ` /SMask ${imageAlloc.smaskId} 0 R` : "";
    addObject(
      imageAlloc.id,
      `${imageAlloc.id} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${data.width} /Height ${data.height} ` +
        `/ColorSpace /DeviceRGB /BitsPerComponent 8${stream.filter}${smask} /Length ${stream.data.length} >>\nstream\n`,
      stream.data,
      "\nendstream\nendobj\n",
    );
    if (imageAlloc.smaskId !== undefined) {
      const alpha = new Uint8Array(data.width * data.height);
      for (let p = 0; p < alpha.length; p++) alpha[p] = data.rgba[p * 4 + 3]!;
      const alphaStream = await flateEncode(alpha);
      addObject(
        imageAlloc.smaskId,
        `${imageAlloc.smaskId} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${data.width} /Height ${data.height} ` +
          `/ColorSpace /DeviceGray /BitsPerComponent 8${alphaStream.filter} /Length ${alphaStream.data.length} >>\nstream\n`,
        alphaStream.data,
        "\nendstream\nendobj\n",
      );
    }
  }

  // 7. Write every object in ascending ID order and record its byte offset.
  const ordered = [...objects].sort((a, b) => a.id - b.id);
  const offsets: number[] = Array.from({ length: nextId }, () => 0);
  for (const object of ordered) {
    if (offsets[object.id] !== 0) {
      throw new Error(`PDF object ${object.id} written twice`);
    }
    offsets[object.id] = offset();
    for (const part of object.parts) push(part);
  }
  for (let id = 1; id < nextId; id++) {
    if (!offsets[id]) {
      throw new Error(`PDF object ${id} was allocated but never written`);
    }
  }

  // 8. Xref Table & Trailer
  const xrefStart = offset();
  const size = nextId;
  let xref = `xref\n0 ${size}\n0000000000 65535 f \n`;
  for (let i = 1; i < size; i++) xref += `${xrefAt(offsets[i]!)} 00000 n \n`;
  xref += `trailer\n<< /Size ${size} /Root ${catalogId} 0 R /Info ${infoId} 0 R >>\nstartxref\n${xrefStart}\n%%EOF\n`;
  push(xref);

  return new Blob([bytes()], { type: "application/pdf" });
}

/** Concatenate the visible text a vector scene draws, one span per text node —
 *  the font-embedding input so a visible node's subset covers its glyphs.
 *  Coordinates are placeholders (the embedding pass reads text + style only). */
export function sceneTextSpans(page: PdfScenePage): PdfTextSpan[] {
  const spans: PdfTextSpan[] = [];
  const walk = (node: PdfSceneNode): void => {
    if (node.type === "group") {
      for (const child of node.children) walk(child);
      return;
    }
    if (node.type !== "text") return;
    const text = node.rows.map((row) => row.text).join("\n");
    if (!text) return;
    spans.push({
      text,
      x: 0,
      y: 0,
      width: 0,
      height: 0,
      fontSize: node.fontSize,
      ...(node.fontFamily ? { fontFamily: node.fontFamily } : {}),
      ...(node.bold ? { bold: true } : {}),
      ...(node.italic ? { italic: true } : {}),
      ...(node.fill ? { color: node.fill.replace(/^#/, "") } : {}),
    });
  };
  for (const node of page.nodes) walk(node);
  return spans;
}

/** One visible scene text row with its baseline in page px (y down). */
interface SceneTextPlacement {
  text: string;
  xPx: number;
  yPx: number;
}

/** Visible text the scene draws (page px) — used to keep the invisible text
 *  layer from re-extracting a run the scene already renders visibly. */
function sceneTextPlacements(page: PdfScenePage): SceneTextPlacement[] {
  const placements: SceneTextPlacement[] = [];
  const walk = (node: PdfSceneNode): void => {
    if (node.type === "group") {
      for (const child of node.children) walk(child);
      return;
    }
    if (node.type !== "text") return;
    const m = node.matrix;
    for (const row of node.rows) {
      if (!row.text) continue;
      placements.push({
        text: row.text,
        xPx: m.a * row.x + m.c * row.y + m.e,
        yPx: m.b * row.x + m.d * row.y + m.f,
      });
    }
  };
  for (const node of page.nodes) walk(node);
  return placements;
}

/** Match a text-layer span to a visible scene row and consume it (one-to-one),
 *  so the same run is never drawn AND extracted twice. */
function consumeVisiblePlacement(
  placements: SceneTextPlacement[],
  span: PdfTextSpan,
  pageHeightPx: number,
): boolean {
  const xPx = span.x / 0.75;
  const yPx = (pageHeightPx * 0.75 - span.y) / 0.75;
  for (let i = 0; i < placements.length; i++) {
    const placement = placements[i]!;
    if (placement.text !== span.text) continue;
    if (Math.abs(placement.xPx - xPx) > 24 || Math.abs(placement.yPx - yPx) > 4) continue;
    placements.splice(i, 1);
    return true;
  }
  return false;
}

/** Extract text spans, link annotations, and form fields from laid-out document pages for PDF export. */
export function extractPdfPageLayers(
  pages: readonly FlowPage[],
  sections: readonly CanvasStageSection[],
  sectionOfPage: readonly number[],
): { textSpans: PdfTextSpan[]; links: PdfLinkAnnotation[]; formFields: PdfFormField[] }[] {
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
    const formFields: PdfFormField[] = [];

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

              const letterSpacing = style?.letterSpacingPx
                ? toPt(style.letterSpacingPx)
                : undefined;

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
                ...(letterSpacing !== undefined ? { letterSpacing } : {}),
                ...(typeof style?.color === "string" ? { color: style.color } : {}),
                tag: "P",
              });

              // Form fields
              if (inline && "formField" in inline && inline.formField) {
                formFields.push({
                  name: inline.formField.name,
                  type: inline.formField.type,
                  pageIndex,
                  rect: [xPt, toPdfY(lineY + line.heightPx), xPt + widthPt, toPdfY(lineY)],
                  value: inline.formField.value,
                  readOnly: inline.formField.readOnly,
                  options: inline.formField.options,
                });
              }

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

    return { textSpans, links, formFields };
  });
}
