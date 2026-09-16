import type { JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";

import type { HostCommandDomain } from "./registry";

/** The revisions domain's view of the host — only what its command bodies touch. */
export interface RevisionsHostView {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** Toggle a task pane (Word's reviewing pane). */
  togglePane(id: "revisions"): void;
  /** Word's Display for Review projection state. */
  setMarkupView(view: "simple" | "all" | "none" | "original"): void;
  /** The Specific People author filter (null = every author). */
  getMarkupAuthors(): string[] | null;
  setMarkupAuthors(authors: string[] | null): void;
  /** Re-render the document projection (display-only markup changes). */
  renderDoc(doc: JSONContent): void;
  /** Re-stamp the markup menus to the live state. */
  syncMarkupMenus(): void;
  getJSON(): JSONContent;
}

/**
 * Track-changes display commands split out of the host element: the reviewing
 * pane, Word's Display for Review projection, the Specific People filter, and
 * the "…All Changes Shown" sweeps.
 */
export class RevisionsHostCommands implements HostCommandDomain {
  constructor(private readonly host: RevisionsHostView) {}

  readonly chrome: readonly string[] = [];

  readonly editor: readonly string[] = [
    "reviewing-pane",
    "display-for-review",
    "review-specific-people",
    "accept-all-changes-shown",
    "reject-all-changes-shown",
  ];

  run(event: string, value?: string): boolean {
    const editor = this.host.editor();
    if (!editor) return false;
    // Review → Reviewing Pane: toggle the revisions pane (Word's vertical
    // reviewing pane listing every tracked change).
    if (event === "reviewing-pane") {
      this.host.togglePane("revisions");
      return true;
    }
    // Word's Display for Review: switch the tracked-changes projection and
    // re-render (the marks in the document are untouched — display only).
    if (event === "display-for-review") {
      if (value === "simple" || value === "all" || value === "none" || value === "original") {
        this.host.setMarkupView(value);
        this.host.renderDoc(this.host.getJSON());
        this.host.syncMarkupMenus();
      }
      return true;
    }
    // Word's Specific People: scope the display to one reviewer ("all" clears
    // the filter). An author outside the filter renders as accepted (Word).
    if (event === "review-specific-people" && value) {
      this.host.setMarkupAuthors(value === "all" ? null : [value]);
      this.host.renderDoc(this.host.getJSON());
      this.host.syncMarkupMenus();
      return true;
    }
    // The "…All Changes Shown" sweeps accept/reject exactly what the display
    // filter shows (every revision when no filter is set).
    if (event === "accept-all-changes-shown" || event === "reject-all-changes-shown") {
      editor.commands[event]?.(this.host.getMarkupAuthors() ?? undefined);
      return true;
    }
    return false;
  }
}
