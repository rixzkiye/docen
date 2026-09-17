/**
 * Canvas render/pagination domain — split out of the host element (see
 * document/index.ts): the projection → layout → paint pipeline, the
 * pagination-feedback field resolve, the per-page diff, and the a11y mirror
 * schedule. The host keeps thin delegates (and #armStage, the stage wiring).
 */

import { compileDocument, type JSONContent } from "@docen/docx";
import {
  projectDocumentOptions,
  type ProjectedFlowBox,
  type ProjectedPageBackground,
  type ProjectedSection,
} from "@docen/docx/layout";
import {
  browserFontMetrics,
  computePageNumberOffsets,
  layoutFlowSections,
  layoutSectionsIncremental,
  TextMeasurer,
  type FlowPage,
  type FlowSection,
} from "@docen/layout";

import type { A11yMirror } from "../canvas/a11y-mirror";
import type { EditBridge } from "../canvas/edit-bridge";
import { deepEq, dirtyPagesOf } from "../canvas/page-eq";
import {
  layFurnitureSections,
  type LaidFurnitureSection,
  type CanvasStage,
  type CanvasStageSection,
} from "../canvas/stage";
import type { CommentsCommands } from "../commands/comments";
import type { DialogCommands } from "../commands/dialogs";
import type { RevisionsCommands } from "../commands/revisions";
import type { SpellingCommands } from "../commands/spelling";
import { THEMES } from "../commands/themes";
import { liveFieldResolver, resolvePageFieldsBounded } from "../field-resolve";
import { customPropertiesOf, finiteNumber, type FieldContext, type FieldFrame } from "../fields";
import { getSettings } from "../settings";

/** Per-slice layout budget (ms) — the incremental walk yields between slices. */
const LAYOUT_SLICE_MS = 12;
/** Pagination-feedback passes before the layout settles (see resolveFields). */
const FIELD_RESOLVE_PASSES = 3;

/** The projection half's output — the flow inputs both the synchronous drain
 *  and the incremental walk lay. */
export interface ProjectedFlowInputs {
  sections: (ProjectedSection & CanvasStageSection)[];
  background?: ProjectedPageBackground;
  flowSections: FlowSection[];
  viewMode: "print" | "web" | "draft" | "read";
  continuous: boolean;
}

type RenderRun = {
  pages: FlowPage[];
  sectionOfPage: number[];
  sections: (ProjectedSection & CanvasStageSection)[];
  background?: ProjectedPageBackground;
  viewMode?: "print" | "web" | "draft" | "read";
};

/** The render domain's view of the host — only what its bodies touch. */
export interface RenderHostView {
  root(): ShadowRoot | null;
  editor(): import("@docen/docx/core").Editor | undefined;
  bridge(): EditBridge | undefined;
  stage(): CanvasStage | undefined;
  stageHost(): HTMLElement | undefined;
  /** Wire the stage (kept on the host — the handler table lives there). */
  armStage(projected: {
    sections: (ProjectedSection & CanvasStageSection)[];
    background?: ProjectedPageBackground;
    viewMode: "print" | "web" | "draft" | "read";
  }): CanvasStage;
  measurer(): TextMeasurer;
  a11yMirror(): A11yMirror;
  a11yTimer(): number | undefined;
  setA11yTimer(value: number | undefined): void;
  comments(): CommentsCommands;
  revisions(): RevisionsCommands;
  spelling(): SpellingCommands;
  dialogs(): DialogCommands;
  getJSON(): JSONContent;
  getTaskpaneState(id: "a11y" | "navigation" | "reveal" | "properties"): boolean;
  hasAttribute(name: string): boolean;
  isConnected(): boolean;
  filename(): string | undefined;
  debug(): string | undefined;
  hyphenation(): {
    auto?: boolean;
    doNotHyphenateCaps?: boolean;
    zoneTw?: number;
    limit?: number;
  };
  markupView(): "simple" | "all" | "none" | "original";
  markupAuthors(): string[] | null;
  markupColors(): "author" | "changeType";
  balloons(): "all" | "comments" | "revisions" | "none";
  hiddenTextShown(): boolean;
  setHiddenTextShown(value: boolean): void;
  fieldCodes(): boolean;
  setFieldCodes(value: boolean): void;
  updateFieldsOnOpen(): boolean;
  setUpdateFieldsOnOpen(value: boolean): void;
  documentSettings(): Record<string, unknown>;
  pages(): readonly FlowPage[];
  setPages(pages: FlowPage[]): void;
  sectionOfPage(): readonly number[];
  setSectionOfPage(sectionOfPage: number[]): void;
  flow(): ProjectedFlowBox | undefined;
  setFlow(flow: ProjectedFlowBox | undefined): void;
  lastRun(): RenderRun | undefined;
  setLastRun(run: RenderRun | undefined): void;
  renderSeq(): number;
  setRenderSeq(seq: number): void;
  mergedView(doc: JSONContent): JSONContent;
  pageOriginOf(
    sections: readonly (ProjectedSection & CanvasStageSection)[],
    sectionOfPage: readonly number[],
  ): (page: number) => { contentLeftPx: number; contentTopPx: number };
  pageInsets(
    flow: ProjectedFlowBox,
    furniture: import("@docen/docx/layout").ProjectedPageFurniture | undefined,
    laid: LaidFurnitureSection | undefined,
  ): import("@docen/layout").FlowPageInsets | undefined;
  setProgress(label?: string): void;
  updateStatus(): void;
  syncStatusLanguage(): void;
  syncActiveTabStops(): void;
  viewMode(): "print" | "web" | "draft" | "read";
}

/**
 * The render pipeline, split out of the host element. Owns no state — the
 * pages/section map/flow/last-run caches stay on the host and are read/written
 * through this view.
 */
export class RenderDomain {
  constructor(private readonly host: RenderHostView) {}

  /** The canvas pipeline's projection half, shared by every layout path:
   *  compile → project (one section per document section) → lay each
   *  section's furniture ONCE (the insets and the painter's bands share the
   *  pass) → assemble the flow inputs. Pure preparation — no pagination. */
  projectFlowSections(doc: JSONContent): ProjectedFlowInputs {
    // Word's Options → Display "Show hidden text" (w:vanish): hidden runs
    // project suppressed (no advance, no ink) unless the shared settings
    // store asks for their display — read per projection, so a store change
    // reaches the next render.
    const showHiddenText = getSettings().writing.showHiddenText;
    this.host.setHiddenTextShown(showHiddenText);
    // The active document theme's font pair feeds the projection's fallback
    // for text with no explicit font (body → minor, headings → major).
    const themeId = (this.host.documentSettings().theme as { id?: string } | undefined)?.id;
    const themeFonts = themeId ? THEMES[themeId]?.fonts : undefined;
    // Balloons are print-layout chrome: Draft/Web/Read project inline markup
    // only (Word hides the markup area outside Print Layout / Web Layout has
    // no margin at all).
    const balloonsOn = this.host.balloons() !== "none" && this.host.viewMode() === "print";
    const { sections, background } = projectDocumentOptions(
      compileDocument(this.host.mergedView(doc)),
      // Word's Display for Review: "simple" is also the all-marks projection
      // minus the review chrome Word draws outside the flow, so only an
      // actual filter (or the change-type palette, or balloons) needs the
      // non-default pass. "simple" maps to "all" inside that pass: the canvas
      // has no simple-markup chrome of its own, and simple must never hide
      // the marks it is supposed to summarize.
      this.host.markupView() !== "simple" ||
        this.host.markupAuthors() ||
        this.host.markupColors() !== "author" ||
        balloonsOn
        ? {
            view: this.host.markupView() === "simple" ? "all" : this.host.markupView(),
            authors: this.host.markupAuthors() ?? undefined,
            colors: this.host.markupColors(),
            ...(balloonsOn ? { balloons: this.host.balloons() } : {}),
          }
        : undefined,
      // Alt+F9: every field projects its instruction instead of the result.
      this.host.fieldCodes(),
      // Options → Display: hidden runs render with their dotted marker
      // instead of being suppressed.
      showHiddenText,
      this.host.hyphenation(),
      themeFonts,
    );
    const stageSections: (ProjectedSection & CanvasStageSection)[] = sections.map((section) => ({
      ...section,
    }));
    // The continuous views (Web Layout / Read Mode) re-box every section to
    // the viewport width and lay it as ONE unbounded page — Word's web view
    // has no page breaks, and its text width follows the window (page margins
    // kept as the gutters). Furniture is a print concept: no insets. Columns
    // stay a print-layout feature in this pass.
    const mode = this.host.viewMode();
    const continuous = mode === "web" || mode === "read";
    if (continuous) {
      // The scroll surface's width (the document area — the stage host is
      // width:fit-content and only reports the pages' own width) minus the
      // page gutter on each side.
      const area = this.host.root()?.querySelector<HTMLElement>("docen-document-area");
      const availW = Math.max(320, (area?.clientWidth ?? 794) - 48);
      for (const section of stageSections) {
        // Columns stay a print-layout feature: drop them at the source so
        // every downstream consumer (flow opts, separator painting) lays a
        // single-column stream.
        section.columns = undefined;
        const marginL = section.flow.contentLeftPx;
        const marginR =
          section.flow.pageWidthPx - section.flow.contentLeftPx - section.flow.contentWidthPx;
        section.flow = {
          ...section.flow,
          pageWidthPx: availW,
          contentWidthPx: Math.max(200, availW - marginL - marginR),
        };
      }
    }
    const laidFurniture = layFurnitureSections(stageSections, browserFontMetrics);
    stageSections.forEach((section, i) => {
      section.furnitureLaid = laidFurniture[i];
    });
    const flowSections = stageSections.map((section) => {
      const pageInsets = continuous
        ? undefined
        : this.host.pageInsets(section.flow, section.furniture, section.furnitureLaid);
      return {
        blocks: section.blocks,
        ...(section.type ? { type: section.type } : {}),
        opts: {
          ...section.flow,
          columns: section.columns,
          footnoteDefinitions: section.footnoteDefinitions,
          endnoteDefinitions: section.endnoteDefinitions,
          endnotePlacement: (section as { endnotePlacement?: "sectEnd" | "docEnd" })
            .endnotePlacement,
          ...(continuous ? { unbounded: true, contentHeightPx: 1_000_000 } : {}),
          ...(pageInsets ? { pageInsets } : {}),
        },
      };
    });
    return { sections: stageSections, background, flowSections, viewMode: mode, continuous };
  }

  /** The layout half in one synchronous drain — the pagination walk over the
   *  projected flow inputs, plus the continuous views' page-height correction
   *  (the unbounded layout reports where the content ends; the host sizes the
   *  page from it). */
  laySections(projected: ProjectedFlowInputs): {
    pages: FlowPage[];
    sectionOfPage: number[];
  } {
    const { pages, sectionOfPage } = layoutFlowSections(
      projected.flowSections,
      this.host.measurer(),
    );
    if (projected.continuous) {
      // Size each continuous page to where its content actually ends plus the
      // bottom margin.
      pages.forEach((page, i) => {
        const section = projected.sections[sectionOfPage[i] ?? 0];
        if (!section || page.contentBottomPx == null) return;
        const flow = section.flow;
        const bottomMargin = flow.pageHeightPx - flow.contentTopPx - flow.contentHeightPx;
        flow.pageHeightPx = Math.max(
          flow.pageHeightPx,
          Math.ceil(page.contentBottomPx) + flow.contentTopPx + bottomMargin,
        );
      });
    }
    return { pages, sectionOfPage };
  }

  /** The canvas pipeline's projection + layout half, shared by the full
   *  render and the story's live re-render. The layout half runs through the
   *  pagination-feedback pass, so fields paint from live numbers. */
  projectAndLayout(doc: JSONContent): {
    pages: FlowPage[];
    sectionOfPage: number[];
    sections: (ProjectedSection & CanvasStageSection)[];
    background?: ProjectedPageBackground;
  } {
    const projected = this.projectFlowSections(doc);
    const resolved = this.resolveFields(projected, this.laySections(projected), doc);
    return {
      pages: resolved.pages,
      sectionOfPage: resolved.sectionOfPage,
      sections: projected.sections,
      background: projected.background,
    };
  }

  /** The per-render field context base: one fixed clock for the whole render
   *  (DATE/TIME stay stable while a single layout settles) plus the document
   *  state the evaluators read — core properties, filename, revision, custom
   *  properties. Word/char counts are edit-time values (NUMWORDS/NUMCHARS are
   *  not live fields), so the render base skips their text walk. */
  renderFieldBase(doc: JSONContent): Omit<FieldContext, "frame" | "sequences"> {
    const attrs = (doc.attrs ?? {}) as {
      core?: Record<string, unknown>;
      documentExtras?: Record<string, unknown>;
    };
    const core = attrs.core ?? {};
    const revision = finiteNumber(core.revision);
    const custom = customPropertiesOf(attrs.documentExtras);
    return {
      now: new Date(),
      core,
      ...(this.host.filename() != null && this.host.filename() !== ""
        ? { filename: this.host.filename() }
        : {}),
      ...(revision != null ? { revision } : {}),
      ...(custom ? { customProperties: custom } : {}),
    };
  }

  /** The pagination-feedback loop: resolve the live numbering fields against
   *  the pages just laid, and re-lay when a resolution rewrote measured text
   *  (the resolved value changes where a field breaks). Bounded by
   *  {@link FIELD_RESOLVE_PASSES}; returns the final pages plus the pages
   *  whose painted field values changed (the incremental path repaints just
   *  those). In field-code view (Alt+F9) the resolver resolves nothing, so
   *  code atoms stay untouched and no re-layout rides the toggle. */
  resolveFields(
    projected: ProjectedFlowInputs,
    laid: { pages: FlowPage[]; sectionOfPage: number[] },
    doc: JSONContent,
  ): { pages: FlowPage[]; sectionOfPage: number[]; dirty?: number[] } {
    return resolvePageFieldsBounded(
      laid.pages,
      projected.sections,
      laid.sectionOfPage,
      liveFieldResolver(this.renderFieldBase(doc), this.host.fieldCodes()),
      () => this.laySections(projected),
      FIELD_RESOLVE_PASSES,
    );
  }

  /** The pagination's view of a document position — the FieldFrame the field
   *  evaluators read (shown page number, page count, section number/span, the
   *  section's numFmt). Undefined before the first layout or when the caret
   *  map has no page for the position. */
  fieldFrame(pos: number): FieldFrame | undefined {
    const pageIndex = this.host.bridge()?.pageOf(pos);
    if (pageIndex == null || pageIndex < 0 || pageIndex >= this.host.pages().length)
      return undefined;
    const sections = this.host.lastRun()?.sections ?? [];
    const section = this.host.sectionOfPage()[pageIndex] ?? 0;
    const offsets = computePageNumberOffsets(sections, this.host.sectionOfPage());
    const sectionPages = this.host.sectionOfPage().reduce((n, s) => (s === section ? n + 1 : n), 0);
    const format = sections[section]?.pageNumbering?.format;
    return {
      page: pageIndex + 1 + (offsets[section] ?? 0),
      pageCount: this.host.pages().length,
      section: section + 1,
      sectionPages,
      ...(format ? { pageFormat: format } : {}),
    };
  }

  /** The canvas pipeline — the single render entry the bridge's transactions
   *  and the loaders share: compile → project → layout → paint, then re-arm
   *  the caret map against the fresh geometry. Page-level diff: only pages
   *  whose laid-out content changed repaint (the rest keep their canvas), so
   *  a keystroke costs one page, not one per scrolled-into-view page. */
  renderDoc(doc: JSONContent): void {
    if (!this.host.stageHost()) return;
    const seq = this.host.renderSeq() + 1;
    this.host.setRenderSeq(seq);
    const projected = this.projectFlowSections(doc);
    // A first render in a paged view (the open path) paints page-by-page:
    // the layout walk yields sealed pages, the stage appends their slots per
    // time slice, and the veil lifts over the first slice — an open shows its
    // first screens in seconds instead of blocking until the last page is
    // laid. Any render landing mid-walk (a transaction, a view switch) bumps
    // the sequence; the walk drops and this entry re-runs from the top.
    if (!this.host.lastRun() && !projected.continuous) {
      void this.renderDocIncremental(projected, seq, doc);
      return;
    }
    const laid = this.resolveFields(projected, this.laySections(projected), doc);
    const run = {
      pages: laid.pages,
      sectionOfPage: laid.sectionOfPage,
      sections: projected.sections,
      background: projected.background,
      viewMode: projected.viewMode,
    };
    this.scheduleA11yMirror(projected.sections);
    const prev = this.host.lastRun();
    this.host.setLastRun(run);
    this.host.setPages(run.pages);
    this.host.setSectionOfPage(run.sectionOfPage);
    this.host.setFlow(run.sections[0]?.flow);
    const stage = this.host.armStage(run);
    // Anything structural (section geometry, page background, section count)
    // repaints everything — the per-page diff only skips pages whose
    // placement, section, AND background are all unchanged. A page-count
    // shift is deliberately NOT structural: deleting across a pagination
    // boundary bounces the count between renders, and a full repaint per
    // bounce is the visible flicker of holding Backspace. dirtyPagesOf
    // covers count changes positionally (pages past either end stay dirty;
    // the trailing slots' lifecycles are handled in sync). The overlapping
    // page range still compares its section map: a deletion may pull a later
    // section onto an existing page slot, which changes its flow/furniture.
    // Furniture compares on the projected options, not the laid stacks (the
    // stacks derive from them plus the already-compared flow width); the
    // frame CSS (background/borders) re-stamps on every sync and needs no
    // diff.
    const structural =
      !prev ||
      prev.viewMode !== run.viewMode ||
      prev.sectionOfPage.some(
        (section, index) =>
          index < run.sectionOfPage.length && section !== run.sectionOfPage[index],
      ) ||
      prev.background?.color !== run.background?.color ||
      prev.background?.image !== run.background?.image ||
      prev.sections.length !== run.sections.length ||
      prev.sections.some((s, i) => !deepEq(s.flow, run.sections[i]!.flow)) ||
      prev.sections.some((s, i) => !deepEq(s.furniture, run.sections[i]!.furniture)) ||
      prev.sections.some((s, i) => !deepEq(s.lineNumbers, run.sections[i]!.lineNumbers)) ||
      prev.sections.some((s, i) => !deepEq(s.pageNumbering, run.sections[i]!.pageNumbering)) ||
      prev.sections.some((s, i) => !deepEq(s.columns, run.sections[i]!.columns));
    const dirty = structural ? undefined : dirtyPagesOf(prev.pages, run.pages);
    stage!.sync(run.pages, run.sections, run.sectionOfPage, run.background, dirty);
    this.host
      .bridge()
      ?.updatePages(run.pages, this.host.pageOriginOf(run.sections, run.sectionOfPage));
    this.afterLayout();
  }

  /** The paged first render: consume {@link layoutSectionsIncremental} in
   *  ~12ms slices, syncing the stage's growing page list each slice (painted
   *  pages keep their canvas — only the slice's tail is dirty). The first
   *  slice lifts the opening veil and hands the painted pages a caret map;
   *  the finished walk records the run and arms the panes without a second
   *  sync. A render starting mid-walk bumps the sequence and this walk just
   *  drops — that render re-projects and takes over. */
  async renderDocIncremental(
    projected: ProjectedFlowInputs,
    seq: number,
    doc: JSONContent,
  ): Promise<void> {
    const stage = this.host.armStage(projected);
    const pages: FlowPage[] = [];
    const sectionOfPage: number[] = [];
    const origin = this.host.pageOriginOf(projected.sections, sectionOfPage);
    const iterator = layoutSectionsIncremental(projected.flowSections, this.host.measurer());
    let done = false;
    while (!done) {
      const painted = pages.length;
      const deadline = performance.now() + LAYOUT_SLICE_MS;
      let step = iterator.next();
      while (!step.done) {
        pages.push(step.value.page);
        sectionOfPage.push(step.value.section);
        if (performance.now() >= deadline) break;
        step = iterator.next();
      }
      if (step.done) done = true;
      // Another render started, or the element went away — the walk is dead.
      if (seq !== this.host.renderSeq() || !this.host.isConnected()) return;
      stage!.sync(
        pages,
        projected.sections,
        sectionOfPage,
        projected.background,
        Array.from({ length: painted }, () => false),
      );
      if (painted === 0) {
        // First slice: real pages are on screen — lift the veil and give the
        // painted range a caret map so clicks and typing already work.
        this.host.setProgress();
        this.host.bridge()?.updatePages(pages, origin);
      }
      if (done) break;
      await new Promise((resolve) => requestAnimationFrame(() => resolve(null)));
    }
    // The whole page list is known now: resolve the live numbering fields
    // against it and repaint just the pages whose painted values changed. A
    // resolution that rewrote measured text (SECTION/SECTIONPAGES width)
    // re-lays synchronously — the pages array is then a fresh one.
    const resolved = this.resolveFields(projected, { pages, sectionOfPage }, doc);
    if (seq !== this.host.renderSeq() || !this.host.isConnected()) return;
    const finalPages = resolved.pages;
    const finalSectionOfPage = resolved.sectionOfPage;
    const finalOrigin = this.host.pageOriginOf(projected.sections, finalSectionOfPage);
    if (finalPages !== pages) {
      stage!.sync(finalPages, projected.sections, finalSectionOfPage, projected.background);
    } else if (resolved.dirty && resolved.dirty.length > 0) {
      const dirty = finalPages.map((_, index) => resolved.dirty!.includes(index));
      stage!.sync(finalPages, projected.sections, finalSectionOfPage, projected.background, dirty);
    }
    this.host.setLastRun({
      pages: finalPages,
      sectionOfPage: finalSectionOfPage,
      sections: projected.sections,
      background: projected.background,
      viewMode: projected.viewMode,
    });
    this.host.setPages(finalPages);
    this.host.setSectionOfPage(finalSectionOfPage);
    this.host.bridge()?.updatePages(finalPages, finalOrigin);
    this.afterLayout();
  }

  /** The hidden semantic mirror, off the render's critical path: a full
   *  rebuild is O(document) DOM, so a keystroke must not pay it — it lands on
   *  a short debounce instead (screen readers read settled text). */
  scheduleA11yMirror(sections: readonly { blocks: unknown }[]): void {
    clearTimeout(this.host.a11yTimer());
    this.host.setA11yTimer(
      window.setTimeout(() => {
        this.host.setA11yTimer(undefined);
        try {
          this.host.a11yMirror().update({
            sections: sections.map((s) => ({ blocks: s.blocks })),
          } as any);
        } catch {
          // Ignore a11y mirror update errors in non-browser environments
        }
      }, 250),
    );
  }

  /** The panes-and-status tail both render paths run after their final sync. */
  afterLayout(): void {
    // w:updateFields (Options → Update fields on open): one Update All Fields
    // against the freshly pinned pagination. Deferred to this tail because the
    // command's REF/PAGEREF lookups read the bridge's page map, which the
    // just-finished render updated. It dispatches only when a cache changed —
    // no render loop.
    if (this.host.updateFieldsOnOpen()) {
      this.host.setUpdateFieldsOnOpen(false);
      this.host.dialogs().updateAllFields();
    }
    this.host.updateStatus();
    if (this.host.getTaskpaneState("a11y")) {
      (this.host.root()?.querySelector("docen-a11y-checker-pane") as any)?.check(
        this.host.getJSON(),
      );
    }
    this.host.comments().syncCommentsPane();
    this.host.revisions().syncRevisionsPane();
    this.host.spelling().schedule();
    this.host.syncStatusLanguage();
    this.host.syncActiveTabStops();
  }
}
