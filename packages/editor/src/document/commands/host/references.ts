import type { Editor } from "@docen/docx/core";
import { twipToPx } from "@docen/layout";

import { t } from "../../../ui";
import type { CrossReferenceTarget } from "../../cross-reference";
import type { HostCommandDomain } from "./registry";

/** The references domain's view of the host — only what its command bodies
 *  touch. */
export interface ReferencesHostView {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — insert/update route into the caret's story, and page
   *  numbers come from the caret map. */
  bridge():
    | {
        activeEditor(): Editor;
        pageOf(pos: number): number | null;
      }
    | undefined;
  /** The current section's flow box (content width for TOC tab stops). */
  flow(): { contentWidthPx: number } | undefined;
  /** The host element — the shadow-DOM root for the reference dialogs. */
  element(): HTMLElement;
  /** Open the Footnote and Endnote settings dialog. */
  openNoteSettings(): void;
  /** Prompt for an index entry and seed an XE field at the selection. */
  markIndexEntry(target: Editor): void;
  insertBibliography(): void;
  bibliographySources(): unknown[];
  crossReferenceTargets(): CrossReferenceTarget[];
  noteInsert(kind: "footnote" | "endnote"): void;
  noteEditAtSelection(): void;
  noteDeleteAtSelection(): void;
  jumpNextNote(): void;
  jumpPreviousNote(): void;
  /** Prompt for a name and wrap the selection with a bookmark pair. */
  insertBookmark(): void;
}

/**
 * References commands split out of the host element: the Table of Contents /
 * Table of Figures / Index insert-update passes, the footnote and endnote
 * dialogs, bookmarks, captions, cross-references, and citations/bibliography.
 */
export class ReferencesHostCommands implements HostCommandDomain {
  constructor(private readonly host: ReferencesHostView) {}

  readonly chrome: readonly string[] = ["note-settings-dialog"];

  readonly editor: readonly string[] = [
    "toc",
    "update-toc",
    "remove-toc",
    "update-toc-page",
    "table-of-figures",
    "update-figures",
    "mark-entry",
    "insert-index",
    "update-index",
    "bookmark",
    "insert-caption",
    "cross-reference",
    "manage-sources",
    "insert-citation",
    "bibliography",
    "insert-footnote",
    "edit-note",
    "delete-note",
    "toc-dialog",
  ];

  run(event: string, value?: string): boolean {
    // References → footnotes group launcher: the Word Footnote and Endnote
    // dialog (document-level numbering settings).
    if (event === "note-settings-dialog") {
      this.host.openNoteSettings();
      return true;
    }
    const editor = this.host.editor();
    if (!editor) return false;
    // TOC insert/update — commands take the bridge's pageOf (entry page
    // numbers come from the canvas caret map; 0-based → Word's 1-based) and
    // the content-width tab stop. Inserting repaginates, so insert re-runs
    // the update once the fresh layout lands (Word's insert-then-update-
    // fields behavior). remove-toc drops the block; update-toc-page is
    // Word's "update page numbers only".
    if (
      event === "toc" ||
      event === "update-toc" ||
      event === "remove-toc" ||
      event === "update-toc-page"
    ) {
      // Insert/update in the story the caret lives in (a header/footer story
      // opening must not send the TOC into the stale main-doc selection).
      const target = this.host.bridge()?.activeEditor() ?? editor;
      const pageOf = (pos: number): number | null => {
        const page = this.host.bridge()?.pageOf(pos);
        return typeof page === "number" ? page + 1 : null;
      };
      const flow = this.host.flow();
      const tabPositionTw = flow ? Math.round(flow.contentWidthPx / twipToPx(1)) : undefined;
      const ran =
        event === "remove-toc"
          ? target.commands["remove-toc"]()
          : event === "update-toc-page"
            ? target.commands["update-toc-page"](pageOf)
            : target.commands[event](pageOf, tabPositionTw);
      if (event === "toc" && ran) {
        // Frame N re-flows (the bridge's raf-merged onDoc), frame N+1 the
        // caret map carries the post-insert pagination.
        requestAnimationFrame(() =>
          requestAnimationFrame(() => target.commands["update-toc"](pageOf, tabPositionTw)),
        );
      }
      return true;
    }
    // Table of Figures — insert/update the caption directory (the TOC field's
    // \c switch): same story routing and post-insert update pass as the TOC.
    if (event === "table-of-figures" || event === "update-figures") {
      const target = this.host.bridge()?.activeEditor() ?? editor;
      const pageOf = (pos: number): number | null => {
        const page = this.host.bridge()?.pageOf(pos);
        return typeof page === "number" ? page + 1 : null;
      };
      const flow = this.host.flow();
      const tabPositionTw = flow ? Math.round(flow.contentWidthPx / twipToPx(1)) : undefined;
      const ran = target.commands[event](pageOf, tabPositionTw);
      if (event === "table-of-figures" && ran) {
        requestAnimationFrame(() =>
          requestAnimationFrame(() => target.commands["update-figures"](pageOf, tabPositionTw)),
        );
      }
      return true;
    }
    // Index — Mark Entry prompts for the entry text and seeds an XE field at
    // the selection (the invisible marker Word hides from the page); insert
    // and update collect the XE fields into the Index-styled entry block.
    if (event === "mark-entry" || event === "insert-index" || event === "update-index") {
      const target = this.host.bridge()?.activeEditor() ?? editor;
      if (event === "mark-entry") {
        this.host.markIndexEntry(target);
        return true;
      }
      const pageOf = (pos: number): number | null => {
        const page = this.host.bridge()?.pageOf(pos);
        return typeof page === "number" ? page + 1 : null;
      };
      const flow = this.host.flow();
      const tabPositionTw = flow ? Math.round(flow.contentWidthPx / twipToPx(1)) : undefined;
      const ran = target.commands[event](pageOf, tabPositionTw);
      if (!ran) window.alert(t("index.empty", this.host.element()));
      return true;
    }
    // Bookmark — prompt for a name and wrap the selection with a
    // bookmarkStart/bookmarkEnd pair (Word's Insert → Bookmark).
    if (event === "bookmark") {
      this.host.insertBookmark();
      return true;
    }
    // Caption — open the dialog; the commit arrives via caption:ok
    // (#dialogs.onCaptionOk, References → Captions → Insert Caption).
    if (event === "insert-caption") {
      (
        this.host.element().shadowRoot?.querySelector("docen-caption-dialog") as {
          show(): void;
        } | null
      )?.show();
      return true;
    }
    // Cross-reference — open the dialog over the document's candidates; the
    // commit arrives via cross-ref:ok (#dialogs.onCrossRefOk). The caret
    // position rides along for the dialog's above/below preview.
    if (event === "cross-reference") {
      const target = this.host.bridge()?.activeEditor() ?? editor;
      (
        this.host.element().shadowRoot?.querySelector("docen-cross-reference-dialog") as {
          show(targets: CrossReferenceTarget[], caretPos?: number): void;
        } | null
      )?.show(this.host.crossReferenceTargets(), target.state.selection.from);
      return true;
    }
    // Source Manager / Insert Citation — the same dialog in two modes (Word's
    // References → Citations & Bibliography group); commits arrive via
    // sources:ok / citation:ok (#references.onSourcesOk / onCitationOk).
    if (event === "manage-sources" || event === "insert-citation") {
      (
        this.host.element().shadowRoot?.querySelector("docen-sources-dialog") as {
          show(mode: "manage" | "cite", sources: unknown[]): void;
        } | null
      )?.show(event === "insert-citation" ? "cite" : "manage", this.host.bibliographySources());
      return true;
    }
    // Bibliography — insert (or rebuild) the Bibliography-styled block beside
    // the caret from the document's sources (#references.insertBibliography).
    if (event === "bibliography") {
      this.host.insertBibliography();
      return true;
    }
    // Footnote / Endnote — open the note dialog; the commit arrives via
    // note:ok (#dialogs.onNoteOk, References → Insert Footnote; the split's
    // endnote item shares the event, and Next Footnote steps references).
    if (event === "insert-footnote") {
      if (value === "endnote") this.host.noteInsert("endnote");
      else if (value === "next") this.host.jumpNextNote();
      else if (value === "prev") this.host.jumpPreviousNote();
      else this.host.noteInsert("footnote");
      return true;
    }
    if (event === "edit-note") {
      this.host.noteEditAtSelection();
      return true;
    }
    if (event === "delete-note") {
      this.host.noteDeleteAtSelection();
      return true;
    }
    // Custom Table of Contents — open the dialog; the commit arrives via
    // toc:ok (#insertCustomToc).
    if (event === "toc-dialog") {
      (
        this.host.element().shadowRoot?.querySelector("docen-toc-dialog") as {
          show(): void;
        } | null
      )?.show();
      return true;
    }
    return false;
  }
}
