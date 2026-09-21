import {
  nodeKit,
  paintColumnSeparators,
  paintEndnotes,
  paintFootnotes,
  paintFurnitureStack,
  paintScene,
  withKit,
  type PaintContext,
} from "@docen/core";
import { compileDocument, type DocumentOptions, type JSONContent } from "@docen/docx";
import { projectDocumentOptions, type ProjectedSection } from "@docen/docx/layout";
import {
  browserFontMetrics,
  createMeasurer,
  layoutFlowSections,
  loadDefaultFonts,
  registerDefaultFonts,
  type FlowSection,
  type TextMeasurer,
} from "@docen/layout";

import {
  buildEmbeddedPdfFonts,
  extractPdfPageLayers,
  pagesToPdf,
  sceneTextSpans,
  type PdfEmbeddableFontSource,
  type PdfEmbeddedFont,
  type PdfExportOptions,
  type PdfPageShot,
  type PdfTextSpan,
} from "./export-pdf";
import { NodeImageCache, type PdfImageCache } from "./node-image";
import { serializeNodeScene } from "./node-scene";
import { buildPdfStructure, type PdfStructureView } from "./pdf-structure";
import { computePageInsets, layFurnitureSections, type PdfStageSection } from "./stage-section";

export interface RenderPdfOptions extends PdfExportOptions {
  /** Shorthand for metadata.title */
  title?: string;
  /** Optional custom text measurer. Default: createMeasurer(browserFontMetrics). */
  measurer?: TextMeasurer;
  /** Optional font sources to embed in the PDF. */
  fontSources?: readonly PdfEmbeddableFontSource[];
  /** Optional custom image cache. Default: new NodeImageCache(). */
  imageCache?: PdfImageCache;
  /** Whether to show paragraph marks / formatting markers. */
  showMarks?: boolean;
}

let cachedDefaultFontSources: readonly PdfEmbeddableFontSource[] | undefined;

async function getDefaultFontSources(): Promise<readonly PdfEmbeddableFontSource[]> {
  if (cachedDefaultFontSources) return cachedDefaultFontSources;
  try {
    const defaultFaces = await loadDefaultFonts();
    // One face per family — the regular slot — matching the browser editor's
    // export-embedding set (registerDefaultFonts keeps only regular faces in
    // its font map; the exporter synthesizes bold/italic). Passing all four
    // slots would let buildEmbeddedPdfFonts' family map keep the bold-italic
    // face for every family, so all embedded text would render bold-italic.
    cachedDefaultFontSources = defaultFaces
      .filter((f) => !f.bold && !f.italic)
      .map((f) => ({
        family: f.family,
        fontData: f.bytes,
      }));
    return cachedDefaultFontSources;
  } catch {
    return [];
  }
}

/** One-entry cache of the last render's embedded subsets. A server that
 *  re-renders the same document (or the memory-discipline loop) reuses the
 *  subsetter's output instead of re-running it per render; the entry is
 *  keyed by the rendered spans and the source array identity, so a different
 *  document or caller-provided fonts always re-subsets. One entry keeps the
 *  retained memory bounded. */
let cachedEmbeddedFonts:
  | {
      key: string;
      sources: readonly PdfEmbeddableFontSource[];
      fonts: readonly PdfEmbeddedFont[];
    }
  | undefined;

function embeddedFontsKey(spans: readonly PdfTextSpan[]): string {
  const parts: string[] = [];
  for (const span of spans) {
    if (!span.text) continue;
    parts.push(`${span.fontFamily ?? ""}\u0001${span.text}`);
  }
  return parts.join("\u0002");
}

function ensureNodeCanvas(): void {
  if (typeof globalThis.OffscreenCanvas !== "undefined") return;
  const fontShorthandEm = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? Number(m[1]) : 16;
  };
  const ctx = () => ({
    _font: "16px serif",
    set font(v: string) {
      this._font = v;
    },
    get font(): string {
      return this._font;
    },
    measureText(s: string): { width: number } {
      const em = fontShorthandEm(this._font);
      let w = 0;
      for (const ch of s) {
        if (ch === " " || ch === "\t") w += em / 4;
        else if (ch.charCodeAt(0) > 0x2e80) w += em;
        else w += em / 2;
      }
      return { width: w };
    },
  });
  (globalThis as any).OffscreenCanvas = class {
    getContext(): unknown {
      return ctx();
    }
  };
}

/** The bundled production faces are process-global (one shaping font manager),
 *  so the first render registers them and every later render reuses that
 *  promise. A failed attempt clears the cache so a later render can retry. */
let defaultFontsRegistration: Promise<void> | undefined;

/** Register docen's bundled faces before any layout/measuring so the default
 *  Word families shape with their real metrics in Node exactly as they do in
 *  the browser editor — without this the measurer silently falls back to the
 *  rough OffscreenCanvas shim and the server PDF drifts from the browser
 *  export. Failure is loud: the assets ship with @docen/layout, so their
 *  absence is a broken install, not a supported degradation. */
function ensureDefaultFonts(): Promise<void> {
  defaultFontsRegistration ??= registerDefaultFonts().then(
    () => undefined,
    (err) => {
      defaultFontsRegistration = undefined;
      throw new Error(
        "[@docen/pdf] failed to load the bundled production fonts — headless renders " +
          "would measure the default Word families with the canvas shim instead of " +
          "their real metrics and stop matching the browser export",
        { cause: err },
      );
    },
  );
  return defaultFontsRegistration;
}

function normalizeJsonContent(node: JSONContent): JSONContent {
  if (!node || typeof node !== "object") return node;
  if (node.type === "heading") {
    const level = Number(node.attrs?.level) || 1;
    const clamped = Math.min(9, Math.max(1, level));
    const content = Array.isArray(node.content) ? node.content.map(normalizeJsonContent) : [];
    return {
      type: "paragraph",
      attrs: {
        ...node.attrs,
        heading: `Heading${clamped}`,
      },
      content,
    };
  }

  if (Array.isArray(node.content)) {
    const newContent: JSONContent[] = [];
    for (const child of node.content) {
      if (node.type === "doc" && child?.type === "image") {
        newContent.push({
          type: "paragraph",
          content: [normalizeJsonContent(child)],
        });
      } else {
        newContent.push(normalizeJsonContent(child));
      }
    }
    return { ...node, content: newContent };
  }

  return node;
}

/**
 * Headless server-side PDF generation entry point — renders a document JSON
 * or DocumentOptions model directly to a PDF Uint8Array without DOM or Canvas.
 *
 * Implements strict per-page memory streaming discipline: each page scene
 * graph is allocated on nodeKit, serialized to plain vector JSON, and freed
 * immediately so transient memory stays bounded to ~5-20 MB.
 */
export async function renderPdf(
  input: JSONContent | DocumentOptions,
  options?: RenderPdfOptions,
): Promise<Uint8Array> {
  ensureNodeCanvas();
  // Bundled faces before any layout/measuring: the furniture pass, the body
  // measurer and the painter all resolve through the same process-global
  // shaping font manager, so one registration covers the whole render.
  await ensureDefaultFonts();
  const kit = nodeKit;

  const docOptions: DocumentOptions =
    typeof input === "object" && input !== null && (input as any).type === "doc"
      ? compileDocument(normalizeJsonContent(input as JSONContent))
      : (input as DocumentOptions);

  const { sections } = projectDocumentOptions(docOptions);
  const stageSections: PdfStageSection[] = sections.map((sec: ProjectedSection) => ({
    flow: sec.flow,
    blocks: sec.blocks,
    pageBorders: sec.pageBorders,
    lineNumbers: sec.lineNumbers,
    pageNumbering: sec.pageNumbering,
    columns: sec.columns,
    furniture: sec.furniture,
    footnoteDefinitions: sec.footnoteDefinitions,
    endnoteDefinitions: sec.endnoteDefinitions,
  }));

  const laidFurniture = layFurnitureSections(stageSections, browserFontMetrics);
  stageSections.forEach((sec, i) => {
    sec.furnitureLaid = laidFurniture[i];
  });

  const flowSections: FlowSection[] = stageSections.map((sec) => {
    const insets = computePageInsets(sec.flow, sec.furniture, sec.furnitureLaid);
    return {
      blocks: sec.blocks ?? [],
      opts: {
        ...sec.flow,
        columns: sec.columns,
        footnoteDefinitions: sec.footnoteDefinitions,
        endnoteDefinitions: sec.endnoteDefinitions,
        ...(insets ? { pageInsets: insets } : {}),
      },
    };
  });

  const measurer = options?.measurer ?? createMeasurer(browserFontMetrics);
  const { pages, sectionOfPage } = layoutFlowSections(flowSections, measurer);

  const layers = extractPdfPageLayers(pages, stageSections, sectionOfPage);

  const view: PdfStructureView = {
    doc: undefined,
    sections: stageSections,
    sectionOfPage,
    pages,
    boxOf: () => null,
  };
  const structure = buildPdfStructure(view);

  const imageCache = options?.imageCache ?? new NodeImageCache();
  const shots: PdfPageShot[] = [];

  for (let i = 0; i < pages.length; i++) {
    const page = pages[i]!;
    const secIdx = sectionOfPage[i] ?? 0;
    const sec = stageSections[secIdx] ?? stageSections[0]!;
    const flow = sec.flow;

    let slot = 0;
    if (sec.furniture?.titlePage && i === 0) slot = 1;
    else if (sec.furniture?.evenAndOddHeaders && (i + 1) % 2 === 0) slot = 2;

    const root = kit.createGroup();
    const ctx: PaintContext = {
      flow,
      columns: sec.columns,
      pageIndex: i,
      pageCount: pages.length,
      layer: "body",
      metrics: browserFontMetrics,
      showMarks: options?.showMarks,
      kit,
      rerender: () => {},
    };

    withKit(kit, () => {
      const behind = kit.createGroup();
      root.add(behind);
      const body = kit.createGroup();
      root.add(body);

      const header = sec.furnitureLaid?.header[slot] ?? sec.furnitureLaid?.header[0];
      const footer = sec.furnitureLaid?.footer[slot] ?? sec.furnitureLaid?.footer[0];

      if (header) {
        paintFurnitureStack(
          body,
          header.stack,
          flow.contentLeftPx,
          sec.furniture?.headerDistancePx ?? 48,
          ctx,
        );
      }
      if (footer) {
        const bottom = flow.pageHeightPx - (sec.furniture?.footerDistancePx ?? 48);
        paintFurnitureStack(body, footer.stack, flow.contentLeftPx, bottom - footer.heightPx, ctx);
      }

      paintScene(body, page.items, ctx);
      // The stage's flat path paints these after the body items; the server
      // render must too or footnotes/endnotes/column rules are missing from
      // the PDF (the invisible text layer would still carry their text).
      paintColumnSeparators(body, ctx);
      paintFootnotes(body, page.footnotes, ctx);
      paintEndnotes(body, page.endnotes, ctx);
    });

    const scene = await serializeNodeScene(root, flow.pageWidthPx, flow.pageHeightPx, imageCache);
    root.clear();

    const layer = layers[i] ?? { textSpans: [], links: [], formFields: [] };
    shots.push({
      width: flow.pageWidthPx,
      height: flow.pageHeightPx,
      scene,
      textSpans: layer.textSpans,
      links: layer.links,
      formFields: layer.formFields,
    });
  }

  const textMode = options?.textMode ?? (options?.pdfa || options?.pdfUa ? "embedded" : undefined);
  let embeddedFonts = options?.embeddedFonts;
  let fontSources = options?.fontSources;
  if (!embeddedFonts && (!fontSources || fontSources.length === 0)) {
    // The browser export always carries the registered faces. Outlines mode
    // still paints unshaped runs (letter-spacing runs, fallback scripts)
    // through the scene's text nodes, so without the embedded subset those
    // runs fall back to standard-14 Helvetica and stop matching the browser.
    fontSources = await getDefaultFontSources();
  }

  if (!embeddedFonts && fontSources && fontSources.length > 0) {
    // Visible scene text (chart/marks labels) uses the same subsets the
    // browser export builds its font set from; the invisible span layer
    // covers the body.
    const allSpans = [
      ...shots.flatMap((s) => s.textSpans ?? []),
      ...shots.flatMap((s) => (s.scene ? sceneTextSpans(s.scene) : [])),
    ];
    const key = embeddedFontsKey(allSpans);
    if (
      cachedEmbeddedFonts &&
      cachedEmbeddedFonts.key === key &&
      cachedEmbeddedFonts.sources === fontSources
    ) {
      embeddedFonts = cachedEmbeddedFonts.fonts;
    } else {
      embeddedFonts = await buildEmbeddedPdfFonts(allSpans, fontSources);
      cachedEmbeddedFonts = { key, sources: fontSources, fonts: embeddedFonts };
    }
  }

  const exportOpts: PdfExportOptions = {
    ...options,
    ...(textMode ? { textMode } : {}),
    metadata: {
      ...options?.metadata,
      ...(options?.title ? { title: options.title } : {}),
    },
    outline: options?.outline ?? structure.outline,
    pageLabels: options?.pageLabels ?? structure.pageLabels,
    destinations: options?.destinations ?? structure.destinations,
    structElements: options?.structElements ?? structure.structElements,
    embeddedFonts,
  };

  const blob = await pagesToPdf(shots, exportOpts);

  return new Uint8Array(await blob.arrayBuffer());
}
