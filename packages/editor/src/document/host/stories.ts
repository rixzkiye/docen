/**
 * Header/footer story domain — split out of the host element (see
 * document/index.ts). Owns the story being edited (`kind` + anchor `page`) and
 * every slots read/write the bridge's story callbacks drive; the host keeps
 * thin delegates for the ribbon commands and the render path.
 */

import type { JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import type {
  ProjectedFlowBox,
  ProjectedPageBackground,
  ProjectedPageFurniture,
  ProjectedSection,
} from "@docen/docx/layout";
import type { FlowPage, FlowPageInsets } from "@docen/layout";

import { t } from "../../ui";
import type { EditBridge, StoryKind, StorySlot } from "../canvas/edit-bridge";
import type { CanvasStage, CanvasStageSection, LaidFurnitureSection } from "../canvas/stage";

export type ProjectedRun = {
  pages: FlowPage[];
  sectionOfPage: number[];
  sections: (ProjectedSection & CanvasStageSection)[];
  background?: ProjectedPageBackground;
};

/** The stories domain's view of the host — only what its bodies touch. */
export interface StoriesHostView {
  /** The host element (locale resolution for `t`). */
  element(): HTMLElement;
  editor(): Editor | undefined;
  bridge(): EditBridge | undefined;
  stage(): CanvasStage | undefined;
  sectionOfPage(): readonly number[];
  projectAndLayout(doc: JSONContent): ProjectedRun;
  /** Adopt a freshly projected/lain run (pages + section map + first flow). */
  setLayout(run: ProjectedRun): void;
  /** Rebuild the editor state from doc JSON (slots on the doc node). */
  loadDoc(doc: JSONContent): void;
  /** Position of the current section's sectPr paragraph (null = doc node). */
  sectPrPosInCurrentSection(): number | null;
  hideContextTab(): void;
  spellingSchedule(): void;
  spellingRun(): void;
  /** Is this inline passthrough atom a PAGE (not NUMPAGES/PAGEREF) field? */
  isPageField(child: JSONContent): boolean;
}

/** Write a slots group through a transaction: the group lives on the
 *  current section's sectPr paragraph when there is one, else on the doc
 *  node (the #loadDoc state-rebuild path — the doc node is not step
 *  addressable; see #persistStory). */
function writeSlots(
  host: StoriesHostView,
  key: "sectionHeaders" | "sectionFooters",
  group: Record<string, unknown>,
): void {
  const bridge = host.bridge();
  const editor = host.editor();
  if (!bridge || !editor) return;
  const { doc, tr } = editor.state;
  const targetPos = host.sectPrPosInCurrentSection();
  if (targetPos != null) {
    const node = doc.nodeAt(targetPos);
    if (node) {
      tr.setNodeMarkup(targetPos, undefined, { ...node.attrs, [key]: group });
      editor.view.dispatch(tr);
      return;
    }
  }
  const raw = bridge.editor.getJSON();
  host.loadDoc({
    ...raw,
    attrs: { ...(raw.attrs as Record<string, unknown>), [key]: group },
  } as JSONContent);
}

/** The slots group as writeSlots addresses it (same container semantics:
 *  the current section's sectPr paragraph, else the doc node). */
function readSlotsGroup(
  host: StoriesHostView,
  key: "sectionHeaders" | "sectionFooters",
): Record<string, unknown> {
  const editor = host.editor();
  if (!editor) return {};
  const targetPos = host.sectPrPosInCurrentSection();
  const group = (attrs: Record<string, unknown> | undefined): Record<string, unknown> =>
    (attrs?.[key] as Record<string, unknown> | undefined) ?? {};
  if (targetPos != null) {
    const node = editor.state.doc.nodeAt(targetPos);
    if (node) return group(node.attrs as Record<string, unknown>);
  }
  return group(editor.state.doc.attrs as Record<string, unknown>);
}

/** Offsets a flow run's slots by one section: clone the walk down to the
 *  patched section so the doc's live attrs never mutate. */
function patchSectionSlots(
  root: JSONContent,
  sectionIndex: number,
  key: "sectionHeaders" | "sectionFooters",
  slot: StorySlot,
  content: JSONContent[],
): JSONContent | undefined {
  let remaining = sectionIndex;
  const walk = (node: JSONContent): JSONContent | undefined => {
    if (node.type === "paragraph") {
      const attrs = (node.attrs ?? {}) as Record<string, unknown>;
      if (attrs.sectionProperties != null) {
        if (remaining === 0) {
          const group = (attrs[key] ?? {}) as Record<string, unknown>;
          return {
            ...node,
            attrs: { ...attrs, [key]: { ...group, [slot]: content } },
          };
        }
        remaining--;
      }
    }
    if (!Array.isArray(node.content)) return undefined;
    for (let i = 0; i < node.content.length; i++) {
      const patched = walk(node.content[i]!);
      if (patched) {
        const children = node.content.slice();
        children[i] = patched;
        return { ...node, content: children };
      }
    }
    return undefined;
  };
  return walk(root);
}

function pageOriginOf(
  sections: readonly (ProjectedSection & CanvasStageSection)[],
  sectionOfPage: readonly number[],
): (page: number) => { contentLeftPx: number; contentTopPx: number } {
  return (page) => {
    const flow = sections[sectionOfPage[page] ?? 0]?.flow;
    if (!flow) return { contentLeftPx: 0, contentTopPx: 0 };
    if (flow.mirrorMargins && page % 2 === 1) {
      return {
        contentLeftPx: flow.pageWidthPx - flow.contentLeftPx - flow.contentWidthPx,
        contentTopPx: flow.contentTopPx,
      };
    }
    return flow;
  };
}

/** Word's furniture overflow rule: a header taller than the top margin
 *  pushes the body down, a taller footer pushes it up — each page by its
 *  own slot's LAID stack (the first page by the first slot when titlePage
 *  asks for one, even pages by the even slot). Slots without their own
 *  content fall back to the default stack (OOXML reference semantics),
 *  matching the stage's paint fallback — the heights are the same layout
 *  pass the painter's bands come from. */
export function pageInsets(
  flow: ProjectedFlowBox,
  furniture: ProjectedPageFurniture | undefined,
  laid: LaidFurnitureSection | undefined,
): FlowPageInsets | undefined {
  if (!furniture) return undefined;
  const topMargin = flow.contentTopPx;
  const bottomMargin = flow.pageHeightPx - flow.contentTopPx - flow.contentHeightPx;
  const headerDistance = furniture.headerDistancePx ?? 48;
  const footerDistance = furniture.footerDistancePx ?? 48;
  const height = (kind: "header" | "footer", slot: 0 | 1 | 2): number | undefined =>
    laid?.[kind][slot]?.heightPx;
  const inset = (headerPx: number | undefined, footerPx: number | undefined) => {
    const top = Math.max(0, headerDistance + (headerPx ?? 0) - topMargin);
    const bottom = Math.max(0, footerDistance + (footerPx ?? 0) - bottomMargin);
    return top > 0 || bottom > 0
      ? { topPx: Math.round(top), bottomPx: Math.round(bottom) }
      : undefined;
  };
  const def = inset(height("header", 0), height("footer", 0));
  if (!def) return undefined;
  const out: FlowPageInsets = { default: def };
  if (furniture.titlePage) {
    out.first =
      inset(
        height("header", 1) ?? height("header", 0),
        height("footer", 1) ?? height("footer", 0),
      ) ?? undefined;
  }
  if (furniture.evenAndOddHeaders) {
    out.even =
      inset(
        height("header", 2) ?? height("header", 0),
        height("footer", 2) ?? height("footer", 0),
      ) ?? undefined;
  }
  return out;
}

/**
 * Header/footer story editing, split out of the host element. The host keeps
 * thin delegates for the ribbon commands and the bridge callbacks.
 */
export class StoriesDomain {
  /** The header/footer story under edit (null = none); #page is the anchor
   *  page the story edits in place on. */
  #kind: StoryKind | null = null;
  #page = -1;

  constructor(private readonly host: StoriesHostView) {}

  /** The story currently under edit, or null. */
  kind(): StoryKind | null {
    return this.#kind;
  }

  /** The story's anchor page (-1 when none). */
  page(): number {
    return this.#page;
  }

  /** The bridge's `story.entered` callback: mark the story under edit, label
   *  its band, show the contextual tab, and let the proofing pass re-check
   *  against the story being edited. */
  enter(kind: StoryKind, _slot: StorySlot, page: number): void {
    this.#kind = kind;
    this.#page = page;
    this.host.stage()?.setStoryEdit({
      kind,
      label: t(kind === "header" ? "story.header" : "story.footer", this.host.element()),
    });
    this.host.hideContextTab();
    // Word re-checks against the story being edited. entered() runs
    // before the bridge registers the story (the strut's onDoc needs
    // the story kind set), so the check must wait out the synchronous
    // entry — by then the active editor is the story, or (rolled back
    // entry) the body again.
    queueMicrotask(() => this.host.spellingRun());
  }

  #slotsKeyOf(kind: StoryKind): "sectionHeaders" | "sectionFooters" {
    return kind === "header" ? "sectionHeaders" : "sectionFooters";
  }

  /** The story's source JSON — the section owning `page` holds the slots
   *  (Word: the band double-clicked edits that page's section, regardless of
   *  where the caret sits); a section without stamped slots is linked to the
   *  previous one, so the walk merges the first content found per slot, and
   *  an absent slot falls back to the default slot's (what the page displays
   *  until the edit breaks the tie). */
  readSource(kind: StoryKind, slot: StorySlot, page: number): JSONContent[] {
    const editor = this.host.editor();
    if (!editor) return [];
    const merged: { default?: JSONContent[]; first?: JSONContent[]; even?: JSONContent[] } = {};
    for (let i = this.host.sectionOfPage()[page] ?? 0; i >= 0; i--) {
      const pos = this.#sectPrPosOfSection(i);
      const attrs =
        pos >= 0
          ? (editor.state.doc.nodeAt(pos)?.attrs as Record<string, unknown> | undefined)
          : (editor.state.doc.attrs as Record<string, unknown> | undefined);
      const group = attrs?.[this.#slotsKeyOf(kind)] as
        | { default?: JSONContent[]; first?: JSONContent[]; even?: JSONContent[] }
        | undefined;
      if (!group) continue;
      merged.default ??= group.default;
      merged.first ??= group.first;
      merged.even ??= group.even;
    }
    return merged[slot] ?? merged.default ?? [];
  }

  /** A story keystroke's render path: patch the slot into a copy of the doc
   *  JSON and re-run the full pipeline — the body re-flows because the
   *  header/footer it edits pushes on it (Word: typing in a header moves the
   *  body live). The fresh furniture stack goes back to the story's map. */
  renderFurniture(kind: StoryKind, slot: StorySlot, json: JSONContent[]): void {
    const bridge = this.host.bridge();
    const stage = this.host.stage();
    if (!bridge || !stage || this.#kind !== kind) return;
    const raw = bridge.editor.getJSON();
    const key = this.#slotsKeyOf(kind);
    // The story edits the section its anchor page belongs to (Word: the band
    // double-clicked) — patch that section's slots, not the doc's: a doc-level
    // write compiles into the LAST section, and a mid-document story would
    // never reach the stack the band reads. getJSON()'s objects carry live PM
    // attrs by reference — patchSectionSlots copies along the walked path, so
    // the editor state never mutates without a transaction (no render, no
    // undo, no docen:change).
    const sectionIndex = this.host.sectionOfPage()[this.#page] ?? 0;
    const patched = patchSectionSlots(raw, sectionIndex, key, slot, json);
    const doc = patched ?? {
      ...raw,
      attrs: {
        ...(raw.attrs as Record<string, unknown>),
        [key]: {
          ...((raw.attrs as Record<string, unknown>)[key] as object | undefined),
          [slot]: json,
        },
      },
    };
    const run = this.host.projectAndLayout(doc);
    this.host.setLayout(run);
    stage.sync(run.pages, run.sections, run.sectionOfPage, run.background);
    bridge.updatePages(run.pages, this.pageOriginOf(run.sections, run.sectionOfPage));
    const band = stage.furnitureBand(kind, this.#page);
    bridge.updateStoryMap(
      band ? stage.furnitureStack(kind, this.#page) : null,
      band ?? { top: 0, bottom: 0, paintY: 0 },
    );
    // The story's transactions never cross the main editor, so the render
    // tail that schedules the body's re-check never runs for them — schedule
    // here (debounced; the check reads the active story).
    this.host.spellingSchedule();
  }

  /** The doc position of the paragraph closing the given section (0-based —
   *  the Nth sectionProperties paragraph in document order), or -1 when that
   *  section closes at the body end (its sectPr lives on the doc node). */
  #sectPrPosOfSection(sectionIndex: number): number {
    const editor = this.host.editor();
    if (!editor) return -1;
    let seen = -1;
    let target = -1;
    editor.state.doc.descendants((node, pos) => {
      if (target >= 0) return false;
      if (
        node.type.name === "paragraph" &&
        (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
      ) {
        seen++;
        if (seen === sectionIndex) {
          target = pos;
          return false;
        }
      }
      return true;
    });
    return target;
  }

  /** Write a finished story's JSON back. The story edits the section its
   *  anchor page belongs to — the caret is no address here (exiting by
   *  clicking another page's body moves it). An earlier section's slots live
   *  on its closing sectPr paragraph and go through a plain setNodeMarkup
   *  transaction — one undo step. The final section closes at the body end
   *  and its slots live on the doc node, which no step can address
   *  (nodeAt(0) is the first child) — they land through #loadDoc's state
   *  rebuild, the same path setJSON takes (history resets with it, like any
   *  document load). */
  #persist(kind: StoryKind, slot: StorySlot, json: JSONContent[], anchorPage: number): void {
    const bridge = this.host.bridge();
    if (!bridge) return;
    const key = this.#slotsKeyOf(kind);
    const slots = (attrs: Record<string, unknown>): Record<string, unknown> => ({
      ...(attrs[key] as object | undefined),
      [slot]: json,
    });
    const sectionIndex = this.host.sectionOfPage()[anchorPage] ?? 0;
    const target = this.#sectPrPosOfSection(sectionIndex);
    if (target < 0) {
      const raw = bridge.editor.getJSON();
      this.host.loadDoc({
        ...raw,
        attrs: {
          ...(raw.attrs as Record<string, unknown>),
          [key]: slots(raw.attrs as Record<string, unknown>),
        },
      } as JSONContent);
      return;
    }
    bridge.editor.commands.command(({ state: s, dispatch }) => {
      const node = s.doc.nodeAt(target)!;
      dispatch?.(
        s.tr.setNodeMarkup(target, undefined, { ...node.attrs, [key]: slots(node.attrs) }),
      );
      return true;
    });
  }

  exit(kind: StoryKind, slot: StorySlot, json: JSONContent[], dirty: boolean): void {
    this.host.hideContextTab();
    this.host.stage()?.setStoryEdit(null);
    this.#kind = null;
    if (dirty) this.#persist(kind, slot, json, this.#page);
    this.#page = -1;
    // Back to the body's issues (the persist above re-laid the doc first).
    this.host.spellingRun();
  }

  /** Remove Header / Remove Footer — drop the story's whole slots group from
   *  the current section (Word removes the content; the slot stops
   *  rendering on every page). */
  remove(kind: StoryKind): void {
    writeSlots(this.host, this.#slotsKeyOf(kind), {});
  }

  /** Remove Page Numbers — strip the PAGE field atoms from every slot of
   *  both stories (Word deletes the fields, leaving their paragraphs). */
  removePageNumbers(): void {
    const strip = (blocks: unknown): unknown => {
      const json = blocks as JSONContent[] | undefined;
      if (!Array.isArray(json)) return blocks;
      return json.map((block) =>
        block.type === "paragraph"
          ? {
              ...block,
              content: (block.content ?? []).filter((c) => !this.host.isPageField(c)),
            }
          : block,
      );
    };
    for (const key of ["sectionHeaders", "sectionFooters"] as const) {
      const group = readSlotsGroup(this.host, key);
      const next: Record<string, unknown> = {};
      for (const slot of ["default", "first", "even"] as const) {
        if (group[slot] !== undefined) next[slot] = strip(group[slot]);
      }
      writeSlots(this.host, key, next);
    }
  }

  /** The page→section origin resolver the bridge's caret maps need (each
   *  page's own section's content-box origin). */
  pageOriginOf(
    sections: readonly (ProjectedSection & CanvasStageSection)[],
    sectionOfPage: readonly number[],
  ): (page: number) => { contentLeftPx: number; contentTopPx: number } {
    return pageOriginOf(sections, sectionOfPage);
  }
}
