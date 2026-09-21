/**
 * File I/O domain — split out of the host element (see document/index.ts):
 * open (docx family / Markdown / RTF / plain text), save (docx/odt/markdown/
 * rtf/html/txt/pdf), the JSON cache, the whole-doc load path, print/share/
 * close/inspect, templates, and the file-input change handlers. The host keeps
 * its public DocenHost surface as thin delegating members.
 */

import {
  EncryptedDocumentError,
  generateDOCX,
  generateHTML,
  generateMarkdown,
  generateODT,
  generatePlainText,
  generateRTF,
  normalizeDocument,
  parseDOCX,
  parseMarkdown,
  parsePlainText,
  parseRTF,
  prepareEmbeddedFonts,
  type DocxVariant,
  type FieldCacheOptions,
  type HtmlGenerateOptions,
  type JSONContent,
} from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import type { ProjectedFlowBox, ProjectedSection } from "@docen/docx/layout";
import { computePageNumberOffsets, type FlowPage } from "@docen/layout";
import { EditorState } from "@tiptap/pm/state";

import { t } from "../../ui";
import type { EditBridge } from "../canvas/edit-bridge";
import type { CanvasStage, CanvasStageSection } from "../canvas/stage";
import {
  buildEmbeddedPdfFonts,
  extractPdfPageLayers,
  pagesToPdf,
  sceneTextSpans,
  type PdfEmbeddedFont,
  type PdfExportOptions,
  type PdfPageShot,
} from "../export-pdf";
import { collectRevisions } from "../extensions/track-changes";
import {
  OpenFormatError,
  SAVE_FORMATS,
  detectOpenFormat,
  suggestedFileName,
  type SaveFormat,
} from "../file-formats";
import { findTemplate, templateLocale } from "../templates";
import { collectBookmarkPages, collectFieldPages, collectTocTargetPages } from "./field-pages";

/** The file-I/O domain's view of the host — only what its bodies touch. */
export interface IOHostView {
  element(): HTMLElement;
  root(): ShadowRoot | null;
  editor(): Editor | undefined;
  bridge(): EditBridge | undefined;
  stage(): CanvasStage | undefined;
  pages(): readonly FlowPage[];
  sectionOfPage(): readonly number[];
  flow(): ProjectedFlowBox | undefined;
  lastRun(): { sections: (ProjectedSection & CanvasStageSection)[] } | undefined;
  /** Host-registered font bytes by lowercased family — used by PDF/DOCX
   *  export embedding (registration lives on the element's `registerFont`). */
  fonts(): ReadonlyMap<string, { family: string; fontData: Uint8Array }>;
  viewMode(): "print" | "web" | "draft" | "read" | "outline";
  lang(): string;
  docxVariant(): DocxVariant;
  setDocxVariant(variant: DocxVariant): void;
  docProtected(): boolean;
  setDocProtected(value: boolean): void;
  protectionMode(): string | undefined;
  setProtectionMode(value: string | undefined): void;
  updateFieldsOnOpen(): boolean;
  setUpdateFieldsOnOpen(value: boolean): void;
  jsonDirty(): boolean;
  setJsonDirty(value: boolean): void;
  cachedJSON(): JSONContent | undefined;
  setCachedJSON(value: JSONContent | undefined): void;
  fileInput(): HTMLInputElement | undefined;
  getAttribute(name: string): string | null;
  setAttribute(name: string, value: string): void;
  emitCancelable(name: "docen:close", detail?: { format?: SaveFormat }): boolean;
  renderChrome(): void;
  renderDoc(doc: JSONContent): void;
  applyDocumentTheme(kind: string, value?: string, persist?: boolean): void;
  snapshotStyles(): void;
  syncEditable(): void;
  syncDocumentSettings?(settings: Record<string, unknown>): void;
}

/**
 * File I/O, split out of the host element. The host's public open and save
 * function surface (plus getJSON/setJSON) stays as thin delegations.
 */
export class IODomain {
  constructor(private readonly host: IOHostView) {}

  /** Open the OS file picker. The accept filter on the input element covers
   *  .docx/.md/.markdown; #onFileChange routes the chosen file by extension
   *  via open(). */
  pickFile(): void {
    this.host.fileInput()?.click();
  }

  /** The alert text for an open refusal: the two detection refusals resolve
   *  through the editor i18n table (en/zh), anything else surfaces its own
   *  message. */
  openRefusalMessage(err: unknown): string {
    if (err instanceof EncryptedDocumentError) {
      return t("open.encrypted", this.host.element());
    }
    if (err instanceof OpenFormatError) {
      if (err.code === "flat-opc") return t("open.flat-opc-unsupported", this.host.element());
      return t("open.unsupported", this.host.element()).replace("{name}", err.file ?? "(unknown)");
    }
    return err instanceof Error ? err.message : String(err);
  }

  /** Save the document in the given format via the native Save As dialog
   *  (showSaveFilePicker) when available so the user picks the location and name;
   *  falls back to a plain download otherwise. The header filename is updated to
   *  match the saved name. Defaults to the open document's own docx-family
   *  variant (docm/dotx/dotm save as themselves). */
  async saveAs(format: Exclude<SaveFormat, "pdf"> = this.host.docxVariant()): Promise<void> {
    const cfg = SAVE_FORMATS[format];
    let data: BlobPart;
    if (format === "markdown") {
      data = this.saveMarkdown();
    } else if (format === "rtf") {
      data = this.saveRTF();
    } else if (format === "html") {
      data = this.saveHTML();
    } else if (format === "txt") {
      data = this.savePlainText();
    } else if (format === "odt") {
      data = (await this.saveODT()) as unknown as BlobPart;
    } else {
      data = (await this.saveDOCX(format)) as unknown as BlobPart;
    }
    await this.saveBlob(data, cfg, true);
  }

  /** File menu → Save as Template: a `.dotx` download of the current document
   *  (template main-part content type via the packer variant). Export-shaped —
   *  the working document keeps its name and format. */
  async saveAsTemplate(): Promise<void> {
    const data = await this.saveDOCX("dotx");
    await this.saveBlob(data as BlobPart, SAVE_FORMATS.dotx, false);
  }

  /** Write a finished blob out through the File System Access picker (adopting
   *  the picked name as the filename when `adoptName`), falling back to a
   *  plain download where the picker doesn't exist. `adoptName` is false for
   *  format exports (PDF) — saving a copy doesn't rename the document. */
  async saveBlob(
    data: BlobPart,
    cfg: { description: string; mime: string; ext: string },
    adoptName: boolean,
  ): Promise<void> {
    const blob = new Blob([data], { type: cfg.mime });
    const suggestedName = suggestedFileName(
      this.host.getAttribute("filename")?.trim() || t("header.doc-name", this.host.element()),
      cfg,
    );
    const picker = (
      window as unknown as {
        showSaveFilePicker?: (opts: {
          suggestedName?: string;
          types?: Array<{ description?: string; accept: Record<string, string[]> }>;
        }) => Promise<{
          name: string;
          createWritable: () => Promise<{
            write: (data: Blob | BufferSource | string) => Promise<void>;
            close: () => Promise<void>;
          }>;
        }>;
      }
    ).showSaveFilePicker;
    if (picker) {
      try {
        const handle = await picker({
          suggestedName,
          types: [{ description: cfg.description, accept: { [cfg.mime]: [cfg.ext] } }],
        });
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        if (adoptName) {
          this.host.setAttribute("filename", handle.name);
          this.host.renderChrome();
        }
        return;
      } catch {
        // The user cancelled the picker (AbortError) or it was blocked — do NOT
        // fall back to a download, which would save despite the cancel. The
        // download fallback below only covers browsers without the picker.
        return;
      }
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = suggestedName;
    a.click();
    URL.revokeObjectURL(url);
  }

  /** Export as PDF (filename menu → Export as PDF): the paginated print
   *  snapshots flatten into a PDF blob. Same view round-trip as #print — a
   *  non-print view re-projects into print shape for the snapshot, then falls
   *  back. The export never renames the document. */
  async saveAsPdf(): Promise<void> {
    const { blob } = await this.buildPdf();
    await this.saveBlob(blob, SAVE_FORMATS.pdf, false);
  }

  /** Build the PDF for the document's print layout without saving it — the
   *  shared body of the menu action and the public export API. Every page's
   *  leafer scene serializes to a vector content stream; the returned pages
   *  carry both the scene and the preview PNG the caller can compare against.
   *  A non-print view re-projects for the export and falls back afterwards.
   *  `options.metadata` overrides the Info-dict defaults — a fixed
   *  `creationDate`/`modDate` makes the export byte-deterministic for
   *  server-side rendering and caching. */
  async buildPdf(options?: { metadata?: PdfExportOptions["metadata"] }): Promise<{
    blob: Blob;
    pages: readonly PdfPageShot[];
    embeddedFonts: readonly PdfEmbeddedFont[];
  }> {
    const mode = this.host.viewMode();
    if (mode !== "print") {
      this.host.stage()?.setViewMode("print");
      this.host.renderDoc(this.getJSON());
    }
    const shots = (await this.host.stage()?.sceneSnapshots()) ?? [];
    if (mode !== "print") {
      this.host.stage()?.setViewMode(mode);
      this.host.renderDoc(this.getJSON());
    }
    if (shots.length === 0) {
      return { blob: new Blob([], { type: "application/pdf" }), pages: [], embeddedFonts: [] };
    }
    const pageLayers = extractPdfPageLayers(
      this.host.pages(),
      this.host.lastRun()?.sections ?? [],
      this.host.sectionOfPage(),
    );
    const shotsWithLayers: PdfPageShot[] = shots.map((shot, i) => ({
      ...shot,
      textSpans: pageLayers[i]?.textSpans,
      links: pageLayers[i]?.links,
    }));
    const fonts = this.host.fonts();
    const embeddedFonts =
      fonts.size > 0
        ? await buildEmbeddedPdfFonts(
            // Visible scene text too: its subset must cover the glyphs it draws.
            [
              ...pageLayers.flatMap((layer) => layer.textSpans),
              ...shots.flatMap((shot) => (shot.scene ? sceneTextSpans(shot.scene) : [])),
            ],
            [...fonts.values()].map((entry) => ({
              family: entry.family,
              fontData: entry.fontData,
            })),
          )
        : [];
    const title = this.host.getAttribute("filename") ?? t("header.doc-name", this.host.element());
    const blob = await pagesToPdf(shotsWithLayers, {
      metadata: { title, author: "Docen", ...options?.metadata },
      tagged: true,
      ...(embeddedFonts.length > 0 ? { embeddedFonts } : {}),
    });
    return { blob, pages: shotsWithLayers, embeddedFonts };
  }

  /** Filename menu → Share: the Web Share sheet where the platform has one
   *  (title only — the document body is not uploaded); otherwise copy the
   *  document URL (Word for the web's share = share a link). */
  async share(): Promise<void> {
    const title = this.host.getAttribute("filename") ?? t("header.doc-name", this.host.element());
    const nav = navigator as Navigator & { share?: (data: ShareData) => Promise<void> };
    if (typeof nav.share === "function") {
      try {
        await nav.share({ title });
        return;
      } catch {
        return; // the user dismissed the sheet (AbortError)
      }
    }
    try {
      await navigator.clipboard.writeText(location.href);
    } catch {
      // Clipboard denied — nothing else to offer.
    }
  }

  /** Filename menu → Close: end the editing session. A host takes over via
   *  docen:close; otherwise the document resets to a blank slate (Word's Close
   *  closes the window — the browser element's equivalent). Unsaved work is
   *  confirmed away — there is no dirty-save model to offer. */
  closeDocument(): void {
    if (this.host.emitCancelable("docen:close")) return;
    if (this.host.jsonDirty() && !window.confirm(t("close.confirm", this.host.element()))) return;
    this.host.setAttribute("filename", t("header.doc-name", this.host.element()));
    this.host.renderChrome();
    this.host.setDocxVariant("docx");
    this.setJSON({ type: "doc", content: [{ type: "paragraph" }] });
  }

  /** The Document Inspector's scan: comment cards in documentExtras and the
   *  distinct revision records (w:ins/w:del/w:rPrChange ids, paragraph
   *  w:pPrChange records included) in the doc. */
  inspectFindings(): { comments: number; revisions: number } {
    const comments = (
      (this.host.editor()?.state.doc.attrs ?? {}) as {
        documentExtras?: { comments?: unknown[] };
      }
    ).documentExtras?.comments?.length;
    const ids = new Set<string>();
    if (this.host.editor()) {
      for (const revision of collectRevisions(this.host.editor()!.state.doc)) {
        ids.add(`${revision.type}:${String(revision.id)}`);
      }
    }
    return { comments: comments ?? 0, revisions: ids.size };
  }

  /** Filename menu → Inspect Document (Word's 检查问题): scan, then show the
   *  findings dialog; its removal buttons come back as events (#onInspect). */
  inspectDocument(): void {
    const dialog = this.host.root()?.querySelector("docen-inspect-dialog") as unknown as {
      setAttribute(name: string, value: string): void;
      show(): void;
    } | null;
    if (!dialog) return;
    dialog.setAttribute("findings", JSON.stringify(this.inspectFindings()));
    dialog.show();
  }

  /** Print only the document pages — never the ribbon/chrome. Each page
   *  canvas rasterizes into a hidden print-only iframe (one image per page at
   *  the page's true paper size, @page margin 0), so the browser's print
   *  dialog receives exactly the paginated document, like Word's print
   *  output. */
  async print(): Promise<void> {
    // Printing always outputs the paginated Print Layout pages (Word prints
    // the paper document whatever the view) — a continuous view re-projects
    // into print shape for the snapshot, then falls back.
    const mode = this.host.viewMode();
    if (mode !== "print") {
      this.host.stage()?.setViewMode("print");
      this.host.renderDoc(this.getJSON());
    }
    const shots = (await this.host.stage()?.printSnapshots()) ?? [];
    if (mode !== "print") {
      this.host.stage()?.setViewMode(mode);
      this.host.renderDoc(this.getJSON());
    }
    if (shots.length === 0) return;
    const first = shots[0]!;
    const frame = document.createElement("iframe");
    Object.assign(frame.style, {
      position: "fixed",
      right: "0",
      bottom: "0",
      width: "0",
      height: "0",
      border: "0",
    });
    document.body.append(frame);
    const doc = frame.contentDocument!;
    doc.open();
    doc.write(`<!doctype html><html><head><title>${this.host.getAttribute("filename") ?? "Document"}</title><style>
      @page { size: ${first.width / 96}in ${first.height / 96}in; margin: 0; }
      html, body { margin: 0; }
      img { display: block; width: 100%; }
      .pg { page-break-after: always; break-after: page; }
      .pg:last-child { page-break-after: auto; break-after: auto; }
    </style></head><body>`);
    for (const s of shots) doc.write(`<div class="pg"><img src="${s.url}"></div>`);
    doc.write("</body></html>");
    doc.close();
    frame.onload = () => {
      const win = frame.contentWindow;
      if (!win) return;
      const cleanup = (): void => frame.remove();
      win.addEventListener("afterprint", cleanup, { once: true });
      win.focus();
      win.print();
      // afterprint can lag behind the dialog closing — sweep after a grace.
      setTimeout(cleanup, 30_000);
    };
  }

  /** Common load path for openDOCX/openMarkdown: adopt a filename, replace the
   *  whole doc node. The #loadDoc wake-up transaction re-renders the canvas
   *  through the bridge. A doc without sectionProperties (parseMarkdown
   *  output, hand-built JSON) lacks the document-level defaults too — doc
   *  styles, page geometry, docGrid — so every heading renders as plain body
   *  text. Normalize on the way in, same gate as setJSON: a parseDOCX payload
   *  carries its own styles/section properties and is left untouched
   *  (normalizeDocument keeps existing attrs keys). */
  applyOpenedJSON(json: JSONContent, filename?: string): void {
    if (filename) this.host.setAttribute("filename", filename);
    if (!(json.attrs as { sectionProperties?: unknown } | undefined)?.sectionProperties) {
      json = normalizeDocument(json);
    }
    this.loadDoc(json);
  }

  /** Load a file into the editor, auto-detecting its format: the docx family
   *  (.docx/.docm/.dotx/.dotm) or Markdown (.md/.markdown). This is the single
   *  entry point the filename-menu "Open…" uses; openDOCX/openMarkdown remain
   *  for when the caller already knows the format (e.g. loading a server-fetched
   *  docx buffer that has no filename). Throws on a Flat OPC .xml and on an
   *  unrecognized extension. */
  async open(file: File): Promise<void> {
    const format = detectOpenFormat(file);
    if (format === "markdown") return this.openMarkdown(file);
    if (format === "rtf") return this.openRTF(file);
    if (format === "text") return this.openPlainText(file);
    return this.openDOCX(file, format);
  }

  /** Load a docx-family document (.docx/.docm/.dotx/.dotm) into the editor from
   *  a File or a buffer (ArrayBuffer / Uint8Array). A File also adopts its name
   *  as the filename; a bare buffer carries no name. `variant` names the package
   *  kind (the detected extension; docx by default) and becomes the document's
   *  save format — macro parts ride through parseDOCX either way. parseDOCX is
   *  async (office-open 0.14): a File is passed through whole and its bytes are
   *  read inside the parse. While loading, an "Opening <name>" veil covers the
   *  canvas (Office shows the same message for a slow open) and the scroller
   *  stays frozen until the document is ready. */
  async openDOCX(
    input: File | ArrayBuffer | Uint8Array,
    variant: DocxVariant = "docx",
  ): Promise<void> {
    const name = input instanceof File ? input.name : undefined;
    this.setProgress(t("status.opening", this.host.element()).replace("{name}", name ?? "DOCX"));
    try {
      // parseDOCX blocks the main thread (File read included) — yield two
      // frames so the veil paints before the freeze (the bar's sweep is
      // compositor-driven and keeps moving through it).
      await this.nextFrame();
      const json = await parseDOCX(input);
      // Adopt the variant only after a successful parse — a failed open must
      // not relabel the still-open document's save format.
      this.host.setDocxVariant(variant);
      this.applyOpenedJSON(json, name);
      await this.nextFrame();
      this.setProgress();
    } catch (err) {
      this.setProgress();
      throw err;
    }
  }

  /** New from Template → load the picked built-in template's model JSON as a
   *  fresh document. The template bodies are localized to the active UI locale;
   *  the new document takes the template's name (Word names a template-born
   *  document after the template) and the standard docx save format. */
  newFromTemplate(id: string): void {
    const template = findTemplate(id);
    if (!template) return;
    const locale = templateLocale(this.host.lang() || document.documentElement.lang);
    this.host.setDocxVariant("docx");
    this.applyOpenedJSON(
      template.build(locale),
      `${t(template.nameKey, this.host.element())}.docx`,
    );
  }

  /** Filename menu → New from Template: open the built-in template gallery. */
  openTemplateDialog(): void {
    const dialog = this.host.root()?.querySelector("docen-template-dialog") as unknown as {
      show(): void;
    } | null;
    dialog?.show();
  }

  /** Load a Markdown file/string into the editor. A File adopts its name as the
   *  filename; a bare string carries no name. */
  async openMarkdown(input: File | string): Promise<void> {
    const name = typeof input === "string" ? undefined : input.name;
    this.setProgress(
      t("status.opening", this.host.element()).replace("{name}", name ?? "Markdown"),
    );
    try {
      const text = typeof input === "string" ? input : await input.text();
      await this.nextFrame();
      // Markdown has no docx-family variant — a new document saves as .docx.
      this.host.setDocxVariant("docx");
      this.applyOpenedJSON(parseMarkdown(text), name);
      await this.nextFrame();
      this.setProgress();
    } catch (err) {
      this.setProgress();
      throw err;
    }
  }

  /** Load an RTF file/string into the editor. A File adopts its name as the
   *  filename; a bare string carries no name. */
  async openRTF(input: File | string): Promise<void> {
    const name = typeof input === "string" ? undefined : input.name;
    this.setProgress(t("status.opening", this.host.element()).replace("{name}", name ?? "RTF"));
    try {
      const text = typeof input === "string" ? input : await input.text();
      await this.nextFrame();
      this.host.setDocxVariant("docx");
      this.applyOpenedJSON(parseRTF(text), name);
      await this.nextFrame();
      this.setProgress();
    } catch (err) {
      this.setProgress();
      throw err;
    }
  }

  /** Load a Plain Text file/string into the editor. A File adopts its name as the
   *  filename; a bare string carries no name. */
  async openPlainText(input: File | string): Promise<void> {
    const name = typeof input === "string" ? undefined : input.name;
    this.setProgress(t("status.opening", this.host.element()).replace("{name}", name ?? "Text"));
    try {
      const text = typeof input === "string" ? input : await input.text();
      await this.nextFrame();
      this.host.setDocxVariant("docx");
      this.applyOpenedJSON(parsePlainText(text), name);
      await this.nextFrame();
      this.setProgress();
    } catch (err) {
      this.setProgress();
      throw err;
    }
  }

  /** Open progress on the canvas veil — a label + indeterminate Fluent
   *  progress bar centered over the document area (Word centers its opening
   *  spinner the same way). Byte reads are a sliver of the load and parse/
   *  layout report nothing, so the bar never fakes a percentage. Clearing
   *  hides the veil. */
  setProgress(label?: string): void {
    const root = this.host.root();
    const veil = root?.querySelector<HTMLElement>(".load-veil");
    if (!veil || !root) return;
    if (label == null) {
      veil.hidden = true;
      return;
    }
    const labelEl = root.querySelector<HTMLElement>(".load-veil .load-label");
    if (!labelEl) return;
    veil.hidden = false;
    labelEl.textContent = label;
  }

  /** Two rAFs — enough for the current progress state to paint before a
   *  synchronous block (parseDOCX) freezes the frame. */
  nextFrame(): Promise<void> {
    return new Promise((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
    );
  }

  /** Serialize the current document to a DOCX buffer. `variant` selects the
   *  package kind (default: the open document's own — a .docm saves as a .docm,
   *  a .dotx as a .dotx) and stamps the main-part content type; macro parts
   *  carried from the source stay in the package.
   *
   *  Host-registered fonts are embedded (word/fontTable.xml + obfuscated
   *  word/fonts/fontN.odttf parts) when their OS/2 fsType permits it — the
   *  engine writes the parts/relationships. */
  async saveDOCX(variant: DocxVariant = this.host.docxVariant()): Promise<Uint8Array> {
    const fonts = this.host.fonts();
    const embedded =
      fonts.size > 0
        ? prepareEmbeddedFonts(
            [...fonts.values()].map((entry) => ({
              family: entry.family,
              fontData: entry.fontData,
            })),
          )
        : [];
    // Generated-field caches: the builder re-derives SEQ/REF/TOC from the
    // model, and the live canvas pagination supplies the page context for
    // PAGE/NUMPAGES/PAGEREF/SECTION so Word never opens on a stale number.
    const fields = this.#fieldCacheOptions();
    const buffer = await generateDOCX(this.getJSON(), {
      variant,
      ...(embedded.length > 0 ? { document: { fonts: embedded } } : {}),
      ...(fields ? { fields } : {}),
    });
    return buffer as unknown as Uint8Array;
  }

  /** Page context for the generated-field pass, from the host's pagination
   *  (null when headless/not yet laid out). */
  #fieldCacheOptions(): FieldCacheOptions | undefined {
    const editor = this.host.editor();
    const pages = this.host.pages();
    if (!editor || pages.length === 0) return undefined;
    const sections = this.host.lastRun()?.sections ?? [];
    const sectionOfPage = this.host.sectionOfPage();
    const pageOffsets = computePageNumberOffsets(sections, sectionOfPage);
    const view = {
      sectionOfPage,
      pageOffsets,
      physicalPageOf: (pos: number) => this.host.bridge()?.pageOf(pos),
    };
    const fieldPages = collectFieldPages(editor.state.doc, view);
    const bookmarkPages = collectBookmarkPages(editor.state.doc, view);
    const tocPages = collectTocTargetPages(editor.state.doc, view);
    return {
      // PAGEREF resolves against the bookmark's page; everything else against
      // the field's own.
      ...(fieldPages.size > 0 || bookmarkPages.size > 0
        ? {
            pageOf: ({ index, bookmark }: { index: number; bookmark?: string }) =>
              bookmark != null ? bookmarkPages.get(bookmark) : fieldPages.get(index),
          }
        : {}),
      pageCount: pages.length,
      // A saved TOC whose cached entries were missing gets real page numbers
      // from the live canvas pagination instead of Word's empty slots.
      ...(tocPages.headingPages.size > 0 || tocPages.captionPages.size > 0
        ? {
            tocPageOf: ({ index, kind }: { index: number; kind: "heading" | "caption" }) =>
              kind === "heading"
                ? tocPages.headingPages.get(index)
                : tocPages.captionPages.get(index),
          }
        : {}),
    };
  }

  /** Serialize the current document to a Markdown string. */
  saveMarkdown(): string {
    return generateMarkdown(this.getJSON());
  }

  /** Serialize the current document to an RTF string. */
  saveRTF(): string {
    return generateRTF(this.getJSON());
  }

  /** Serialize the current document to an HTML string. */
  saveHTML(options?: HtmlGenerateOptions): string {
    return generateHTML(this.getJSON(), options);
  }

  /** Serialize the current document to plain text. */
  savePlainText(): string {
    return generatePlainText(this.getJSON());
  }

  /** Serialize the current document to an OpenDocument Text (.odt) zip buffer. */
  async saveODT(): Promise<Uint8Array> {
    return generateODT(this.getJSON());
  }

  /** Current document as Tiptap JSON. Cached — recomputed only after a doc
   *  change (see #onTransaction). */
  getJSON(): JSONContent {
    const editor = this.host.editor();
    if (!editor) return {} as JSONContent;
    if (this.host.jsonDirty() || this.host.cachedJSON() === undefined) {
      this.host.setCachedJSON(editor.getJSON());
      this.host.setJsonDirty(false);
    }
    return this.host.cachedJSON()!;
  }

  /** Replace the document with Tiptap JSON. */
  setJSON(json: JSONContent): void {
    // A hand-built JSON (not from parseDOCX) lacks office-open's document-level
    // schema defaults — doc.attrs.styles (docDefaults body font/size/spacing)
    // and doc.attrs.sectionProperties (page size/margins/docGrid linePitch).
    // Without them the document has no body font, no page geometry, and no grid
    // for snapToGrid to pitch against. Normalize once on the way in; a doc that
    // already carries sectionProperties (a parseDOCX/getJSON round-trip) is a
    // no-op (normalizeDocument shallow-merges user attrs over defaults).
    if (!(json.attrs as { sectionProperties?: unknown } | undefined)?.sectionProperties) {
      json = normalizeDocument(json);
    }
    this.loadDoc(json);
    this.host.renderChrome();
  }

  /** Replace the whole doc node (content + doc-level attrs) via a fresh
   *  EditorState. Tiptap's setContent only swaps content and drops doc-level
   *  attrs; this carries them (styles/core/sectionProperties). updateState
   *  bypasses appendTransaction/onTransaction, so extensions that react to doc
   *  changes wouldn't wake — dispatch a docChanged tr (re-stamp the first
   *  block's attrs, a no-op visually) to trigger them: Outline re-reports the
   *  anchor list, and the bridge's raf-merged onDoc re-renders the canvas. */
  loadDoc(doc: JSONContent): void {
    const editor = this.host.editor();
    if (!editor) return;
    // New document — invalidate the JSON cache.
    this.host.setJsonDirty(true);
    editor.view.updateState(
      EditorState.create({ doc: editor.schema.nodeFromJSON(doc), plugins: editor.state.plugins }),
    );
    // NOTE: no isDestroyed guard — the viewless editor's `isDestroyed` getter
    // defaults to true (it reads editorView, which element:null never sets).
    // updateState bypasses appendTransaction, so extensions that react to doc
    // changes wouldn't wake. Dispatch a docChanged tr to fire them. The tr
    // re-stamps the LAST leaf block's OWN attrs — a true no-op (same node,
    // same attrs) — so nothing is clobbered.
    const state = editor.state;
    // Last textblock/leaf block (deepest, rightmost) for the re-stamp — found
    // by descending the rightmost-child chain (O(depth)) instead of a full
    // nodesBetween scan (O(n)).
    const last = this.lastMarkupTarget(state.doc);
    if (last) {
      // addToHistory:false — this re-stamp is an intentional no-op (same node,
      // same attrs) whose sole purpose is to fire appendTransaction (updateState
      // bypasses it). Left in history, it plants a no-op undo entry at the stack
      // bottom (undo returns true but changes nothing); excluding it keeps the
      // undo stack clean after load.
      editor.view.dispatch(
        state.tr.setNodeMarkup(last.pos, undefined, last.attrs).setMeta("addToHistory", false),
      );
    } else {
      // An empty document has no markup target — render directly.
      this.host.renderDoc(editor.getJSON());
    }
    // The style-set gallery's "document default" restores the styles model the
    // document loaded with — captured at this load boundary (state settled),
    // never per layout, or the preset commands' own re-renders would overwrite
    // the snapshot and the restore would replay the current state.
    this.host.snapshotStyles();
    // Document settings ride documentExtras.settings — re-read the protection
    // at this load boundary (a previous document's state must not leak), then
    // re-derive editability. Word also forces revision tracking on when a
    // document opens under a tracked-changes restriction.
    const settings = this.documentSettings();
    this.host.syncDocumentSettings?.(settings);
    const docProtection = settings.documentProtection as
      | {
          edit?: string;
          hash?: string;
          formatting?: boolean;
        }
      | undefined;
    const protection = docProtection?.edit;
    this.host.setDocProtected(protection === "readOnly" || protection === "comments");
    this.host.setProtectionMode(protection);
    if (protection === "trackedChanges") {
      editor.commands["track-changes"](true);
    }
    const pane = this.host.root()?.querySelector("docen-restrict-editing-pane") as {
      setProtectionState?(state: any, hash?: string): void;
    } | null;
    if (pane && protection && protection !== "none") {
      pane.setProtectionState?.(
        {
          isEnforced: true,
          type: protection,
          formattingRestricted: Boolean(docProtection?.formatting),
        },
        docProtection?.hash,
      );
    }
    // w:updateFields — Word updates fields when the document opens. Arm the
    // flag here; the first completed render consumes it (fresh page map).
    if (settings.updateFields === true) this.host.setUpdateFieldsOnOpen(true);
    // Restore persisted document theme if present
    const theme = settings.theme as { id?: string; kind?: string } | undefined;
    if (theme?.id) {
      this.host.applyDocumentTheme(theme.kind || "theme", theme.id, false);
    }
    this.host.syncEditable();
  }

  /** The open document's settings.xml slice, as stored in
   *  doc.attrs.documentExtras.settings (the toggleSectionFlag channel). */
  documentSettings(): Record<string, unknown> {
    const attrs = (this.host.editor()?.state.doc.attrs ?? {}) as {
      documentExtras?: { settings?: Record<string, unknown> };
    };
    return attrs.documentExtras?.settings ?? {};
  }

  /** Last textblock/leaf block (deepest, rightmost) for the #loadDoc re-stamp
   *  hack — the re-stamp target that fires the extension wake-up. Runs only on
   *  load (setJSON/openDOCX), not per edit, so the walk cost is amortized over
   *  the load itself. */
  lastMarkupTarget(doc: import("@tiptap/pm/model").Node): {
    pos: number;
    attrs: Record<string, unknown>;
  } | null {
    let last: { pos: number; attrs: Record<string, unknown> } | null = null;
    doc.nodesBetween(0, doc.content.size, (node, pos) => {
      if (node.isText) return;
      if (node.isTextblock || node.isLeaf) {
        last = { pos, attrs: node.attrs as Record<string, unknown> };
      }
      // Don't descend into textblocks (their text isn't a markup target).
      return node.isTextblock ? false : undefined;
    });
    return last;
  }

  readonly onFileChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    // Reset so picking the same file twice still fires `change`.
    input.value = "";
    if (!file) return;
    // Surface the detection/parse refusal (unsupported type, Flat OPC XML)
    // instead of dropping it as an unhandled rejection.
    void this.open(file).catch((err: unknown) => {
      window.alert(this.openRefusalMessage(err));
    });
  };

  /** Insert the picked image as a data URL. Width/height are left unset — the
   *  canvas renders the natural size, and prepareImages fills them on DOCX
   *  export. */
  readonly onImageChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (): void => {
      // readAsDataURL always yields a string — the guard narrows the union.
      if (typeof reader.result !== "string") return;
      const src = reader.result;
      // Natural size → attrs, clamped to the content width (Word inserts at
      // natural size but never wider than the frame, keeping the aspect).
      // Without explicit dimensions renderDocx falls back to a flat 400×300,
      // which distorts every non-default-shaped picture.
      const img = new Image();
      img.onload = (): void => {
        this.host.bridge()?.focus();
        const contentW = this.host.flow()?.contentWidthPx ?? 620;
        const scale = Math.min(1, contentW / Math.max(1, img.naturalWidth));
        this.host.editor()?.commands.insertContent({
          type: "image",
          attrs: {
            src,
            width: Math.round(img.naturalWidth * scale),
            height: Math.round(img.naturalHeight * scale),
          },
        });
      };
      img.src = src;
    };
    reader.readAsDataURL(file);
  };

  /** Swap the selected image's source for the picked file (Change Picture):
   *  the frame keeps its size, the crop resets — the command side owns both. */
  readonly onPictureChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (): void => {
      if (typeof reader.result !== "string") return;
      this.host.bridge()?.focus();
      this.host.editor()?.commands["change-picture"](reader.result);
    };
    reader.readAsDataURL(file);
  };
}
