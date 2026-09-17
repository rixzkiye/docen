/**
 * Status bar / zoom / Navigation-pane domain — split out of the host element
 * (see document/index.ts). Owns the zoom level, the cached word count, and the
 * Navigation pane's thumbnail cache; the host keeps thin delegating members so
 * the element's own API (setZoom/getZoom) and call sites are unchanged.
 */

import type { JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import type { ProjectedFlowBox } from "@docen/docx/layout";
import type { FlowPage } from "@docen/layout";
import { redoDepth, undoDepth } from "@tiptap/pm/history";
import type { Node as PMNode } from "@tiptap/pm/model";

import { proofingLanguageName } from "../../ui/components/workspace/language-dialog";
import type { WordCountStats } from "../../ui/components/workspace/word-count-dialog";
import { textCounter, wordCounter } from "../addin";
import type { EditBridge } from "../canvas/edit-bridge";
import type { CanvasStage } from "../canvas/stage";

/** Idle pass that refreshes the cached word count after content changes. */
const STATUS_WORD_COUNT_IDLE_MS = 200;
/** Idle pass that re-rasterizes the Navigation pane's page thumbnails. */
const NAV_THUMB_IDLE_MS = 250;

/** The status / zoom / nav-pane view of the host — only what these bodies
 *  touch, so the domain never sees the whole element. */
export interface StatusHostView {
  root(): ShadowRoot | null;
  editor(): Editor | undefined;
  bridge(): EditBridge | undefined;
  stage(): CanvasStage | undefined;
  pages(): readonly FlowPage[];
  sectionOfPage(): readonly number[];
  /** The first section's flow box (page-size presets read its geometry). */
  flow(): ProjectedFlowBox | undefined;
  /** The active view, normalized. */
  viewMode(): "print" | "web" | "draft" | "read";
  /** The caret's proofing language. */
  caretLanguage(): { value: string; noProof: boolean };
  /** Whether the given task pane is currently open. */
  taskpaneOpen(id: "navigation" | "reveal"): boolean;
  /** Refresh the Reveal Formatting pane from the caret. */
  updateReveal(): void;
  /** Write the `view` attribute (status-bar view buttons). */
  setView(view: string): void;
  /** Emit `docen:zoom-change` after a real zoom change. */
  emitZoom(zoom: number): void;
}

/**
 * Status-bar mirror + zoom + Navigation-pane thumbnails, split out of the host
 * element. The host element keeps `setZoom`/`getZoom` as thin delegations.
 */
export class StatusDomain {
  #zoom = 100;
  /** Cached doc nodeSize + Office-style word count so caret-move transactions
   *  don't re-walk the whole document (recomputed only when content changes). */
  #lastDocSize = -1;
  #lastWords = 0;
  /** Idle pass that refreshes #lastWords after content changes. */
  #wordCountTimer?: ReturnType<typeof setTimeout>;
  /** Signature (page count : current page : thumbnail generation) of the last
   *  <docen-nav-pages> push — skips rebuilding the pane's list when nothing it
   *  shows actually changed. */
  #navPagesKey = "";
  #navThumbGen = 0;
  #navThumbs: (string | null)[] = [];
  /** Idle pass that re-rasterizes the Navigation pane's page thumbnails. */
  #navThumbTimer?: ReturnType<typeof setTimeout>;

  constructor(private readonly host: StatusHostView) {}

  /** Reconnect-safe teardown: drop the idle timers and force a recount on the
   *  next connect (even at the same docSize). */
  dispose(): void {
    clearTimeout(this.#wordCountTimer);
    this.#wordCountTimer = undefined;
    // A timed-out recount must re-run on reconnect, even at the same docSize.
    this.#lastDocSize = -1;
    clearTimeout(this.#navThumbTimer);
    this.#navThumbTimer = undefined;
  }

  /** Current zoom level (percent). */
  getZoom(): number {
    return this.#zoom;
  }

  #setZoom(pct: number): void {
    const next = Math.max(10, Math.min(500, Math.round(pct)));
    if (next === this.#zoom) return;
    this.#zoom = next;
    this.host.stage()?.setZoom(next);
    // The frames resized under the overlays — re-place them at the new scale.
    this.host.bridge()?.replaceOverlays();
    this.updateStatus();
    this.host.emitZoom(this.#zoom);
  }

  setZoom(pct: number): void {
    this.#setZoom(pct);
  }

  /** Resolve a zoom preset to a percent. Numeric presets map directly; the
   *  geometric ones read the stage viewport against the flow box (layout px
   *  at 100%) — page width fills the area width, text width fills it with the
   *  content column, one page fits the whole sheet into the visible height. */
  zoomPreset(preset: string): void {
    if (/^\d+$/.test(preset)) return this.#setZoom(Number(preset));
    const area = this.host.root()?.querySelector("docen-document-area");
    const flow = this.host.flow();
    if (!area || !flow) return;
    if (preset === "page-width") return this.#setZoom((area.clientWidth / flow.pageWidthPx) * 100);
    if (preset === "text-width")
      return this.#setZoom((area.clientWidth / flow.contentWidthPx) * 100);
    if (preset === "fit-page") {
      // Whole sheet visible: the net content-box height (clientHeight includes
      // the area's paddings, which would clip the page edges otherwise).
      const style = getComputedStyle(area);
      const visible =
        area.clientHeight -
        Number.parseFloat(style.paddingTop) -
        Number.parseFloat(style.paddingBottom);
      return this.#setZoom(
        Math.min(area.clientWidth / flow.pageWidthPx, visible / flow.pageHeightPx) * 100,
      );
    }
  }

  /** The Zoom dialog (View → Zoom, the status-bar percent click) — prefilled
   *  with the current zoom; the commit applies the preset or free percent. */
  showZoomDialog(): void {
    (
      this.host.root()?.querySelector("docen-zoom-dialog") as {
        show(zoom: number): void;
      } | null
    )?.show(this.#zoom);
  }

  readonly onZoomOk = (event: CustomEvent<string | number>): void => {
    if (typeof event.detail === "number") this.#setZoom(event.detail);
    else this.zoomPreset(event.detail);
  };

  readonly onZoomOpen = (): void => {
    this.showZoomDialog();
  };

  readonly onWordCountOpen = (): void => {
    this.showWordCount();
  };

  /** Status-bar zoom slider → apply the new zoom level. Named (not inline) so
   *  it can be removed on disconnect. */
  readonly onZoomChange = (event: CustomEvent<{ zoom: number }>): void => {
    this.#setZoom(event.detail.zoom);
  };

  /** A status-bar view button (the detail names the status-bar's view:
   *  "reading" | "print" | "web") → the `view` attribute. */
  readonly onViewSelect = (event: CustomEvent<{ view?: string }>): void => {
    const v = event.detail?.view;
    this.host.setView(v === "reading" ? "read" : v === "web" ? "web" : "print");
  };

  /** Refresh the status bar to mirror Word's bottom row: the left cluster is
   *  the caret's section, then "Page X of Y", then the word count; the right
   *  cluster is the zoom slider value + percent. Runs from the coalesced UI
   *  sync (caret moves, a re-render changes the page count) and on zoom /
   *  locale change. The word count is cached by doc nodeSize and refreshed on
   *  an idle pass, so typing never re-walks the full document. */
  updateStatus(): void {
    const root = this.host.root();
    if (!root) return;
    const bar = root.querySelector<HTMLElement>("docen-status-bar");
    const editor = this.host.editor();
    const page = editor ? (this.host.bridge()?.pageOf(editor.state.selection.from) ?? -1) + 1 : 0;
    const total = this.host.pages().length;
    // The caret's section: the section its page belongs to (1-based).
    const section = page > 0 ? (this.host.sectionOfPage()[page - 1] ?? 0) + 1 : 1;
    // Word count is cached by doc nodeSize so caret moves skip re-walking the
    // full document (CharacterCount.words() regexes all text). A content
    // change only schedules the idle recount; the bar keeps the last finished
    // count until it lands.
    const docSize = editor?.state.doc.nodeSize ?? 0;
    if (docSize !== this.#lastDocSize) {
      this.#lastDocSize = docSize;
      if (docSize === 0) this.#lastWords = 0;
      else this.#scheduleWordCount();
    }
    let wordsVal = String(this.#lastWords);
    if (editor && !editor.state.selection.empty) {
      const selText = editor.state.doc.textBetween(
        editor.state.selection.from,
        editor.state.selection.to,
        " ",
      );
      const selWords = (selText.trim().match(/\S+/g) || []).length;
      wordsVal = `${selWords} / ${this.#lastWords}`;
    }
    // Push the numeric state to <docen-status-bar>; it localizes + renders.
    // Guard each write — re-stamping an unchanged attribute re-renders the bar
    // for nothing (the pass runs per frame while typing).
    if (bar) {
      const attrs: Record<string, string> = {
        section: String(section),
        page: String(page || 1),
        total: String(total || 1),
        words: wordsVal,
        zoom: String(this.#zoom),
        view: this.host.viewMode(),
      };
      for (const [name, value] of Object.entries(attrs)) {
        if (bar.getAttribute(name) !== value) bar.setAttribute(name, value);
      }
    }
    // The QAT history carets follow the undo/redo depths live (the header
    // only rebuilds on chrome renders — Word hides the flyout on an empty
    // stack; documentStyles' [data-history-empty] rule drops the caret).
    const liveEditor = this.host.bridge()?.activeEditor() ?? editor;
    for (const [kind, depth] of [
      ["undo", liveEditor ? undoDepth(liveEditor.state) : 0],
      ["redo", liveEditor ? redoDepth(liveEditor.state) : 0],
    ] as const) {
      root
        .querySelector(`docen-ribbon-split-button[data-history="${kind}"]`)
        ?.toggleAttribute("data-history-empty", depth === 0);
    }
    // The pane's page list follows the page count / current page immediately;
    // its thumbnails (one PNG encode per page) refresh on an idle pass.
    this.pushNavPages(total, page || 1);
    if (this.host.taskpaneOpen("navigation")) this.#scheduleNavThumbnails();
    if (this.host.taskpaneOpen("reveal")) {
      this.host.updateReveal();
    }
  }

  /** Recount the Office-style word count once typing pauses (debounced). The
   *  recount re-runs #updateStatus so the bar picks the number up. */
  #scheduleWordCount(): void {
    if (this.#wordCountTimer !== undefined) return;
    this.#wordCountTimer = setTimeout(() => {
      this.#wordCountTimer = undefined;
      const cc = this.host.editor()?.storage.characterCount as { words?: () => number } | undefined;
      this.#lastWords = cc?.words?.() ?? 0;
      this.updateStatus();
    }, STATUS_WORD_COUNT_IDLE_MS);
  }

  /** Push page count / current page / cached thumbnails to the Navigation
   *  pane — skipped entirely when none of them changed. */
  pushNavPages(total: number, current: number): void {
    const navPages = this.host.root()?.querySelector("docen-nav-pages") as
      | (HTMLElement & {
          setPageCount(count: number, current?: number, thumbnails?: (string | null)[]): void;
        })
      | null;
    if (!navPages || total <= 0) return;
    const key = `${total}:${current}:${this.#navThumbGen}`;
    if (key === this.#navPagesKey) return;
    this.#navPagesKey = key;
    navPages.setPageCount(total, current, this.#navThumbs);
  }

  /** Re-rasterize the Navigation pane's thumbnails on an idle pass (only
   *  while the pane is open — the pages' canvas PNGs are expensive). */
  #scheduleNavThumbnails(): void {
    if (this.#navThumbTimer !== undefined) return;
    this.#navThumbTimer = setTimeout(() => {
      this.#navThumbTimer = undefined;
      const total = this.host.pages().length;
      if (total === 0 || !this.host.taskpaneOpen("navigation")) return;
      const thumbs: (string | null)[] = [];
      for (let i = 0; i < total; i++) {
        thumbs.push(this.host.stage()?.pageThumbnail(i) ?? null);
      }
      this.cacheNavThumbs(thumbs);
      this.pushNavPages(
        total,
        (this.host.bridge()?.pageOf(this.host.editor()?.state.selection.from ?? 0) ?? 0) + 1,
      );
    }, NAV_THUMB_IDLE_MS);
  }

  /** Adopt a freshly rasterized thumbnail set as the Navigation pane's cache
   *  (the generation bump makes the next #pushNavPages re-render the list). */
  cacheNavThumbs(thumbs: (string | null)[]): void {
    this.#navThumbs = thumbs;
    this.#navThumbGen += 1;
  }

  /** Rasterize every page (forcing off-screen slots through one render pass)
   *  and feed the Navigation pane real thumbnails for the whole document. The
   *  sync #updateStatus path only fills pages whose canvas already exists;
   *  this is the pane-open completion that covers the rest. */
  async refreshNavThumbnails(): Promise<void> {
    const stage = this.host.stage();
    if (!this.host.root()?.querySelector("docen-nav-pages") || !stage) return;
    if (this.host.pages().length === 0) return;
    const thumbs = await stage.pageThumbnails();
    if (thumbs.length === 0) return;
    this.cacheNavThumbs(thumbs);
    this.pushNavPages(
      this.host.pages().length,
      (this.host.bridge()?.pageOf(this.host.editor()?.state.selection.from ?? 0) ?? 0) + 1,
    );
  }

  /** Word Count (Review tab) — compute the document statistics twice (Word's
   *  dialog shape): the body alone, and with textboxes + footnotes/endnotes
   *  folded back in — the dialog's "include" toggle (default ON) switches
   *  between the two readouts. Textboxes are wpsShape subtrees and textbox
   *  nodes in the body; the notes live in the documentExtras channels. */
  showWordCount(): void {
    const editor = this.host.editor();
    const dialog = this.host.root()?.querySelector("docen-word-count-dialog") as
      | (HTMLElement & { stats?: string; statsExtra?: string; show(): void })
      | undefined;
    if (!editor || !dialog) return;
    // Walk the doc once: paragraphs/text under a wpsShape subtree or inside a
    // textbox node are the textbox bucket, everything else the body bucket.
    let bodyText = "";
    let bodyParas = 0;
    let shapeText = "";
    let shapeParas = 0;
    const walk = (node: PMNode, inShape: boolean): void => {
      const shape = inShape || node.type.name === "wpsShape" || node.type.name === "textbox";
      if (node.type.name === "paragraph") {
        if (shape) {
          shapeParas++;
          shapeText += `${node.textContent}\n`;
        } else {
          bodyParas++;
          bodyText += `${node.textContent}\n`;
        }
        return;
      }
      node.forEach((child) => walk(child, shape));
    };
    walk(editor.state.doc, false);
    // Footnotes/endnotes — documentExtras note bodies are paragraph JSON.
    const extras =
      (
        editor.state.doc.attrs as {
          documentExtras?: {
            footnotes?: Array<{ children?: JSONContent[] }>;
            endnotes?: Array<{ children?: JSONContent[] }>;
          };
        }
      ).documentExtras ?? {};
    let notesText = "";
    let notesParas = 0;
    const jsonText = (node: JSONContent): string =>
      (typeof node.text === "string" ? node.text : "") +
      (node.content ?? []).map(jsonText).join("");
    for (const channel of [extras.footnotes, extras.endnotes]) {
      for (const note of channel ?? []) {
        for (const para of note.children ?? []) {
          notesParas++;
          notesText += `${jsonText(para)}\n`;
        }
      }
    }
    const counted = (text: string, paras: number): WordCountStats => ({
      pages: this.host.pages().length,
      words: wordCounter(text),
      charsWithSpaces: textCounter(text),
      charsNoSpaces: textCounter(text.replace(/\s+/g, "")),
      paragraphs: paras,
      lines: this.layoutLines(),
    });
    dialog.stats = JSON.stringify(counted(bodyText, bodyParas));
    const merged = {
      text: bodyText + shapeText + notesText,
      paras: bodyParas + shapeParas + notesParas,
    };
    dialog.statsExtra = JSON.stringify(counted(merged.text, merged.paras));
    dialog.show();
  }

  /** The laid-out line total (paragraph blocks across every page). */
  layoutLines(): number {
    let lines = 0;
    for (const page of this.host.pages()) {
      for (const item of page.items) {
        if (item.block.kind === "paragraph") lines += item.block.lines.length;
      }
    }
    return lines;
  }

  /** Mirror the caret's proofing language into the status bar (Word shows the
   *  selection's language there). */
  syncStatusLanguage(): void {
    this.host
      .root()
      ?.querySelector("docen-status-bar")
      ?.setAttribute("language", proofingLanguageName(this.host.caretLanguage().value));
  }
}
