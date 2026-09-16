import type { Editor } from "@docen/docx/core";

import { pageNumberInlinePreset, pageNumberStoryPreset } from "../../page-number";
import type { HostCommandDomain } from "./registry";

/** The header/footer domain's view of the host — only what its command bodies
 *  touch. */
export interface HeaderFooterHostView {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — story entry/exit, page lookup, and the active editor. */
  bridge():
    | {
        activeEditor(): Editor;
        pageOf(pos: number): number | null;
        enterStory(kind: "header" | "footer", page: number): boolean;
        exitStory(): unknown;
      }
    | undefined;
  /** The editor input commands must target (the open story, else the main
   *  editor). */
  activeEditor(): Editor | null | undefined;
  /** The page the open story belongs to (-1 = the main document). */
  storyPage(): number;
  /** first-page-different / odd-even-different section flags. */
  toggleSectionFlag(flag: "titlePage" | "evenAndOddHeaders"): void;
  /** Drop a header/footer story from the current section. */
  removeStory(kind: "header" | "footer"): void;
  /** Strip the section's page numbering. */
  removePageNumbers(): void;
  /** Open the Page Number Format dialog. */
  openPageNumberFormat(): void;
}

/**
 * Header/Footer and Page Number commands split out of the host element: the
 * story splits (edit/remove/slot flags), the Header & Footer context tab's
 * story switches, and the Page Number menu (placements, current-position
 * presets, removal, format).
 */
export class HeaderFooterHostCommands implements HostCommandDomain {
  constructor(private readonly host: HeaderFooterHostView) {}

  readonly chrome: readonly string[] = [];

  readonly editor: readonly string[] = [
    "header",
    "footer",
    "goto-header",
    "goto-footer",
    "close-header-footer",
    "header-option",
    "page-number",
  ];

  run(event: string, value?: string): boolean {
    const editor = this.host.editor();
    if (!editor) return false;
    // Header/Footer — the split's main action opens the story on the caret's
    // page; the drop-down carries remove + the slot-visibility flags.
    if (event === "header" || event === "footer") {
      if (value === "title-page" || value === "odd-even") {
        this.host.toggleSectionFlag(value === "title-page" ? "titlePage" : "evenAndOddHeaders");
        return true;
      }
      if (value === "remove-header" || value === "remove-footer") {
        this.host.removeStory(event);
        return true;
      }
      const page = this.host.bridge()?.pageOf(editor.state.selection.from);
      if (page != null) this.host.bridge()?.enterStory(event, page);
      return true;
    }
    // The Header & Footer context tab — switch stories (the dirty close rides
    // the normal exit path), flip the same slot flags, and close.
    if (event === "goto-header" || event === "goto-footer") {
      const page =
        this.host.storyPage() >= 0
          ? this.host.storyPage()
          : this.host.bridge()?.pageOf(editor.state.selection.from);
      this.host.bridge()?.exitStory();
      if (page != null)
        this.host.bridge()?.enterStory(event === "goto-header" ? "header" : "footer", page);
      return true;
    }
    if (event === "close-header-footer") {
      this.host.bridge()?.exitStory();
      return true;
    }
    if (event === "header-option") {
      if (value === "title-page" || value === "odd-even")
        this.host.toggleSectionFlag(value === "title-page" ? "titlePage" : "evenAndOddHeaders");
      return true;
    }
    // Page Number — the split's main button is Word's default (bottom of
    // page, centered). Top/bottom placements open the story and REPLACE its
    // content with the preset paragraph (picking a placement states the
    // intent outright; undo keeps the previous furniture reachable), and the
    // normal exit persists the slots. Current-position presets splice into
    // the caret's paragraph — the active story when one is open, else the
    // body.
    if (event === "page-number") {
      const placement = value ?? "page-bottom-center";
      if (placement === "remove-numbers") {
        this.host.removePageNumbers();
        return true;
      }
      if (placement === "format") {
        this.host.openPageNumberFormat();
        return true;
      }
      if (placement.startsWith("cur-")) {
        (this.host.activeEditor() ?? editor).commands.insertContent(
          pageNumberInlinePreset(placement),
        );
        return true;
      }
      const page = this.host.bridge()?.pageOf(editor.state.selection.from);
      if (
        page != null &&
        this.host.bridge()?.enterStory(placement.startsWith("page-top") ? "header" : "footer", page)
      ) {
        (this.host.activeEditor() ?? editor).commands.setContent(pageNumberStoryPreset(placement));
      }
      return true;
    }
    return false;
  }
}
