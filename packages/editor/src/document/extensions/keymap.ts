import { Extension } from "@docen/docx/core";

/**
 * Centralized Tiptap keymap for docen EDITING shortcuts (MS Office-aligned).
 *
 * Each shortcut dispatches its mapped command directly on `editor.commands`
 * — every command name IS a native Tiptap command (see ./commands), so a
 * shortcut and its ribbon button share ONE definition with no bridge. Add
 * entries to {@link KEYBOARD_SHORTCUTS} to bind more.
 *
 * ── Scope — two layers, by necessity ──
 *
 * This extension binds only EDITING-layer shortcuts: keystrokes that mutate the
 * ProseMirror document via a command. Chrome-layer shortcuts operate on the UI
 * shell (not the document) and live as a host `keydown` listener in index.ts
 * (`#onZoomKey`), because they read host state the editor cannot reach and must
 * be ignored inside inputs/comboboxes:
 *   • Ctrl/Cmd + = / +   → zoom in        (canvas CSS zoom)
 *   • Ctrl/Cmd + - / _   → zoom out
 *   • Ctrl/Cmd + 0       → zoom reset 100%
 *   • Ctrl/Cmd + F       → open Find
 *   • Ctrl/Cmd + H       → open Find & Replace
 *
 * Per-extension defaults (bold=Mod-B, italic=Mod-I, HardBreak Mod/Shift-Enter,
 * the ListKeymap) stay with their owning extensions by Tiptap convention —
 * but ONLY a real EditorView dispatches those. The canvas route is viewless,
 * so the canvas input bridge (edit-bridge.ts) matches the SAME table in its
 * own keydown handler; keep every editing shortcut here so the two
 * consumers cannot drift.
 */
export const KEYBOARD_SHORTCUTS: Readonly<Record<string, string>> = {
  // Ctrl+Enter → page break, Ctrl+Shift+Enter → column break (Word). Shift+Enter
  // (soft line break) stays on @tiptap/extension-hard-break's default. High
  // priority: HardBreak also maps Mod-Enter (to a soft break), and these must win.
  "Mod-Enter": "page-break",
  "Mod-Shift-Enter": "column-break",

  // Character formatting (W4.1)
  "Mod-B": "bold",
  "Mod-I": "italic",
  "Mod-U": "underline",
  "Mod-Shift-X": "strike",
  "Mod-Shift-W": "underline-words",
  "Mod-Shift-K": "small-caps",
  "Mod-Shift-D": "underline-double",
  "Mod-Shift->": "grow-font",
  "Mod-Shift-<": "shrink-font",
  "Mod-Space": "clear-format",
  "Mod-D": "font-dialog",

  // Format painter (W4.1)
  "Mod-Shift-C": "copy-format",
  "Mod-Shift-V": "paste-format",

  // Paragraph formatting (W4.1, W4.2)
  "Mod-Q": "clear-paragraph-format",
  "Mod-M": "indent-increase",
  "Mod-Shift-M": "indent-decrease",
  "Mod-T": "hanging-indent-increase",
  "Mod-Shift-T": "hanging-indent-decrease",
  "Mod-E": "align-center",
  "Mod-L": "align-left",
  "Mod-R": "align-right",
  "Mod-J": "justify",
  "Mod-Shift-8": "show-marks",
  "Mod-Shift-*": "show-marks",

  // Line spacing & paragraph space (W4.2)
  "Mod-1": "line-spacing:1",
  "Mod-2": "line-spacing:2",
  "Mod-5": "line-spacing:1.5",
  "Mod-0": "line-spacing:toggle-before",

  // Styles (W4.1, W4.2)
  "Mod-Shift-N": "style:Normal",
  "Mod-Alt-1": "style:Heading1",
  "Mod-Alt-2": "style:Heading2",
  "Mod-Alt-3": "style:Heading3",

  // Lists (W4.2)
  "Mod-Shift-L": "bullet-list",
  "Mod-Shift-7": "ordered-list",
  "Mod-Shift-&": "ordered-list",

  // Review & notes (W4.4)
  "Mod-Shift-E": "track-changes",
  "Mod-Alt-M": "new-comment",
  "Mod-Alt-F": "insert-footnote",
  "Mod-Alt-D": "insert-footnote:endnote",

  // Outline / Action (W4.5)
  "Alt-Shift-ArrowLeft": "promote-heading",
  "Alt-Shift-Left": "promote-heading",
  "Alt-Shift-ArrowRight": "demote-heading",
  "Alt-Shift-Right": "demote-heading",
  "Alt-Shift-ArrowUp": "move-row-up",
  "Alt-Shift-Up": "move-row-up",
  "Alt-Shift-ArrowDown": "move-row-down",
  "Alt-Shift-Down": "move-row-down",

  // Function keys (W4.3)
  F3: "autotext-f3",
  "Shift-F3": "change-case:cycle",
};

export const DocenKeymap = Extension.create({
  name: "docenKeymap",
  priority: 1000,
  addKeyboardShortcuts() {
    return Object.fromEntries(
      Object.entries(KEYBOARD_SHORTCUTS).map(([key, event]) => [
        key,
        () => {
          const [name, arg] = event.split(":");
          const cmd = (
            this.editor.commands as unknown as Record<
              string,
              ((arg?: string) => boolean) | undefined
            >
          )[name];
          return typeof cmd === "function" ? cmd(arg) : false;
        },
      ]),
    );
  },
});
