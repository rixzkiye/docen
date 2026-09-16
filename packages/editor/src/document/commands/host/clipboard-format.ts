import type { Editor } from "@docen/docx/core";

import type { HostCommandDomain } from "./registry";

/** The clipboard/formatting domain's view of the host — only what its command
 *  bodies touch. */
export interface ClipboardFormatHostView {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The editor input commands must target (the open story, else the main
   *  editor). */
  activeEditor(): Editor | null | undefined;
  /** The host element — the shadow-DOM root for the Paste Special dialog. */
  element(): HTMLElement;
  /** Copy/cut the active story's selection through the bridge's slice lane. */
  copySelection(cut: boolean): Promise<void> | undefined;
  /** Paste through the clipboard lanes (keep-text-only drops the payload). */
  paste(textOnly: boolean): Promise<void> | undefined;
  /** Toggle a task pane (the Office Clipboard pane). */
  togglePane(id: "clipboard"): void;
  /** Open the Styles task pane and rebuild its content. */
  showTaskpane(id: "styles"): void;
  renderStylesPane(): void;
  /** Toggle the Markdown input mode (the bridge reads it per keystroke). */
  toggleMarkdownInput(): void;
  /** Re-stamp the format buttons to the live state. */
  syncFormatButtons(): void;
  /** Prompt for a link address and mark the selection. */
  insertLink(): void;
  /** The href of the link under the caret (null when none). */
  hrefAtCaret(): string | null;
  /** Select the bookmark (the `#name` open-link lane). */
  jumpToBookmark(name: string): void;
  /** Editing → Select (all/objects/similar). */
  select(value?: string): void;
  /** Toggle the format painter's capture/apply mode. */
  toggleFormatPainter(): void;
}

/**
 * Clipboard, editing, hyperlink, and format commands split out of the host
 * element: the QAT history + Repeat, Markdown input, link/unset/open/copy,
 * copy/cut/paste, the Clipboard and Styles panes, Select, and Format Painter.
 */
export class ClipboardFormatHostCommands implements HostCommandDomain {
  constructor(private readonly host: ClipboardFormatHostView) {}

  readonly chrome: readonly string[] = ["undo", "redo", "repeat"];

  readonly editor: readonly string[] = [
    "markdown-input",
    "link",
    "unset-link",
    "open-link",
    "copy-link",
    "copy",
    "cut",
    "paste",
    "clipboard-dialog",
    "styles-pane",
    "select",
    "format-painter",
  ];

  run(event: string, value?: string): boolean {
    // The QAT history flyout's entries arrive as undo/redo carrying their step
    // count — batch-run that many commands (each its own transaction, stop at
    // the first refusal); the value-less primary click falls through to the
    // wired single-step command.
    if ((event === "undo" || event === "redo") && value) {
      const editor = this.host.activeEditor();
      const steps = Number(value);
      if (editor && Number.isInteger(steps) && steps > 0) {
        for (let i = 0; i < steps; i++) {
          if (!(event === "redo" ? editor.commands.redo() : editor.commands.undo())) break;
        }
      }
      return true;
    }
    // Repeat (QAT; F4 lives in the bridge) — retype the last plain-text
    // insertion at the caret. The bridge records it on the editor's storage;
    // both entry points read that single source.
    if (event === "repeat") {
      const editor = this.host.activeEditor();
      const repeat = (editor?.storage as { repeat?: string } | undefined)?.repeat;
      if (editor && repeat) {
        const { from, to } = editor.state.selection;
        editor.view.dispatch(editor.state.tr.insertText(repeat, from, to));
      }
      return true;
    }
    const editor = this.host.editor();
    if (!editor) return false;
    // Markdown input mode toggle — the same flag the Options dialog writes;
    // the bridge reads it per keystroke.
    if (event === "markdown-input") {
      this.host.toggleMarkdownInput();
      // The click may land outside any transaction — re-stamp the lit state.
      this.host.syncFormatButtons();
      return true;
    }
    // Link — prompt for an address and mark the selection (or insert fresh
    // display text when the selection is empty).
    if (event === "link") {
      this.host.insertLink();
      return true;
    }
    // Context menu → Remove Hyperlink: unset the link mark across the
    // right-clicked link (extendMarkRange reaches past the caret's spot).
    if (event === "unset-link") {
      // The link mark spans the right-clicked range in whichever editor the
      // caret lives in (a furniture story has its own links).
      (this.host.activeEditor() ?? editor)?.chain().extendMarkRange("link").unsetLink().run();
      return true;
    }
    // Context menu → Open Hyperlink: `#name` jumps to its bookmark, anything
    // else opens in a new window.
    if (event === "open-link") {
      const href = this.host.hrefAtCaret();
      if (!href) return true;
      if (href.startsWith("#")) this.host.jumpToBookmark(href.slice(1));
      else window.open(href, "_blank", "noopener,noreferrer");
      return true;
    }
    // Context menu → Copy Hyperlink: the address to the system clipboard.
    if (event === "copy-link") {
      const href = this.host.hrefAtCaret();
      if (href) void navigator.clipboard.writeText(href);
      return true;
    }
    // Clipboard — the selection is canvas-rendered (no DOM editor selection),
    // so copy/cut route through the bridge's lane (it pins the slice payload
    // exactly like a keyboard copy, keeping every paste entry lossless).
    if (event === "copy" || event === "cut") {
      void this.host.copySelection(event === "cut");
      return true;
    }
    if (event === "paste-special" || (event === "paste" && value === "paste-special")) {
      (
        this.host.element().shadowRoot?.querySelector("docen-paste-special-dialog") as unknown as {
          show(): void;
        } | null
      )?.show();
      return true;
    }
    if (event === "paste") {
      void this.host.paste(value === "keep-text-only");
      return true;
    }
    // Home → Clipboard group launcher — the Office Clipboard pane.
    if (event === "clipboard-dialog") {
      this.host.togglePane("clipboard");
      return true;
    }
    // Styles group launcher (Home → Styles): the Styles task pane.
    if (event === "styles-pane") {
      this.host.showTaskpane("styles");
      this.host.renderStylesPane();
      return true;
    }
    // Editing → Select: selectAll() spans the whole document.
    if (event === "select") {
      this.host.select(value);
      return true;
    }
    // Format Painter — toggle capture/apply of the current run's marks.
    if (event === "format-painter") {
      this.host.toggleFormatPainter();
      return true;
    }
    return false;
  }
}
