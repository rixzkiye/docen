import type { JSONContent } from "@docen/docx";
import { DOCEN_CLIP_MIME, parseHTMLBody, parseRTF } from "@docen/docx";
import type { Editor } from "@docen/docx/core";

import { t } from "../../ui";
import type { PasteSpecialFormat } from "../../ui/components/workspace/paste-special-dialog";
import type { EditBridge } from "../canvas/edit-bridge";

/** Word's Match Destination Formatting in paste-options form: the block
 *  structure survives (lists, tables, images), the source's run formatting
 *  (bold/italic/color/…) is dropped so the destination's takes over. */
function stripRunMarks(nodes: JSONContent[]): JSONContent[] {
  return nodes.map((n) => {
    const content = n.content ? stripRunMarks(n.content) : undefined;
    return n.type === "text" ? { ...n, marks: undefined } : content ? { ...n, content } : n;
  });
}

function insertPlainText(editor: Editor, text: string): void {
  if (text.includes("\n")) {
    const lines = text.split(/\r?\n/);
    const paragraphs = lines.map((line) => ({
      type: "paragraph",
      content: line ? [{ type: "text", text: line }] : [],
    }));
    editor.commands.insertContent(paragraphs);
  } else {
    editor.commands.insertContent(text);
  }
}

/** Where the paste options bar hangs — the pasted content's last line, as
 *  frame-relative screen px (the bridge's paste anchor). */
export type PasteSource = { kind: "slice" | "html" | "rtf"; raw: string; text: string };

/** The clipboard domain's view of the host — resolved per call so the
 *  controller can be built before a document opens. */
export interface ClipboardHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The edit bridge (the paste lanes and the active-story routing). */
  bridge(): EditBridge | undefined;
  /** The host element — the i18n scope and the shadow root the clipboard
   *  pane lives in. */
  element(): HTMLElement;
}

/**
 * The clipboard domain, split out of the host element: the system-clipboard
 * paste lanes (docen slice → styled HTML → RTF → plain text), Word's paste-options
 * bar (the picks replay the same content in the picked form), and the
 * Office Clipboard pane's session collection.
 */
export class ClipboardCommands {
  constructor(private readonly host: ClipboardHost) {}

  /** Paste from the system clipboard. Delegated to {@link pasteSpecial}. */
  async paste(textOnly = false): Promise<void> {
    return this.pasteSpecial(textOnly ? "text" : undefined);
  }

  /** Paste using a specific format or automatic ladder when omitted. */
  async pasteSpecial(format?: PasteSpecialFormat): Promise<void> {
    const bridge = this.host.bridge();
    const editor = bridge?.activeEditor() ?? this.host.editor();
    if (!editor) return;
    bridge?.focus();

    const allowSlice = !format || format === "slice";
    const allowHtml = !format || format === "html";
    const allowRtf = !format || format === "rtf";
    const allowText = !format || format === "text";

    const docenType = allowSlice ? `web ${DOCEN_CLIP_MIME}` : null;
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (docenType && item.types.includes(docenType)) {
          const raw = await (await item.getType(docenType)).text();
          if (raw && bridge?.insertSlicePayload(raw)) {
            const plain =
              item.types.includes("text/plain") &&
              (await (await item.getType("text/plain")).text());
            if (plain) this.showPasteOptions({ kind: "slice", raw, text: plain });
            return;
          }
        }
        if (allowHtml && item.types.includes("text/html")) {
          const text = await (await item.getType("text/html")).text();
          if (text) {
            const body = new DOMParser().parseFromString(text, "text/html").body;
            const content = parseHTMLBody(body, editor.state.schema).content ?? [];
            if (content.length) {
              editor.commands.insertContent(content);
              const plain = item.types.includes("text/plain")
                ? await (await item.getType("text/plain")).text()
                : "";
              this.showPasteOptions({ kind: "html", raw: text, text: plain });
              return;
            }
          }
        }
        const rtfMime = allowRtf
          ? item.types.find((t) => t === "text/rtf" || t === "application/rtf")
          : null;
        if (rtfMime) {
          const text = await (await item.getType(rtfMime)).text();
          if (text) {
            const json = parseRTF(text);
            const content = (json.content ?? []).filter((n) => n.type !== "text" || n.text);
            if (content.length) {
              editor.commands.insertContent(content);
              const plain = item.types.includes("text/plain")
                ? await (await item.getType("text/plain")).text()
                : "";
              this.showPasteOptions({ kind: "rtf", raw: text, text: plain });
              return;
            }
          }
        }
        if (allowText && item.types.includes("text/plain")) {
          const text = await (await item.getType("text/plain")).text();
          if (!text) continue;
          const pinned = allowSlice ? bridge?.copiedSlice() : null;
          if (pinned && pinned.text === text && bridge?.insertSlicePayload(pinned.payload)) {
            return;
          }
          insertPlainText(editor, text);
          return;
        }
      }
    } catch {
      // read() may be denied (permission policy) — fall through to readText.
    }
    if (allowText || allowSlice) {
      try {
        const text = await navigator.clipboard.readText();
        if (text) {
          const pinned = allowSlice ? bridge?.copiedSlice() : null;
          if (pinned && pinned.text === text && bridge?.insertSlicePayload(pinned.payload)) {
            this.showPasteOptions({ kind: "slice", raw: pinned.payload, text });
            return;
          }
          if (allowText) {
            insertPlainText(editor, text);
          }
        }
      } catch {
        /* clipboard unavailable — nothing to paste */
      }
    }
  }

  /** The clipboard content behind the live paste-options bar — the source
   *  the three picks replay in their picked form (Word's paste options). */
  #pasteBar?: HTMLElement;
  #pasteSource?: PasteSource;

  /** Word's paste-options bar — after a rich paste it hangs below the pasted
   *  content; the three picks undo the insertion and replay the same
   *  clipboard content in the picked form (source / destination-matched /
   *  text only). Also the bridge's onRichPaste landing (the bridge ran the
   *  insertion; the bar only offers the replay picks). */
  showPasteOptions(source: PasteSource): void {
    this.hidePasteOptions();
    const bridge = this.host.bridge();
    const editor = bridge?.activeEditor() ?? this.host.editor();
    const anchor = editor ? bridge?.pasteAnchorRect(editor.state.selection.from) : null;
    if (!anchor) return;
    this.#pasteSource = source;
    const bar = document.createElement("div");
    bar.setAttribute("data-docen-overlay", "");
    Object.assign(bar.style, {
      position: "absolute",
      zIndex: "7",
      display: "flex",
      gap: "2px",
      padding: "3px",
      background: "var(--docen-color-bg, #ffffff)",
      border: "1px solid var(--docen-color-divider, #e2e2e2)",
      borderRadius: "var(--borderRadiusMedium, 6px)",
      boxShadow: "var(--shadow4, 0 4px 8px rgba(0,0,0,.14))",
    } satisfies Partial<CSSStyleDeclaration>);
    const picks: Array<{ mode: "source" | "match" | "text"; key: string }> = [
      { mode: "source", key: "ribbon.opt.keep-source" },
      { mode: "match", key: "ribbon.opt.match-dest" },
      { mode: "text", key: "ribbon.opt.keep-text-only" },
    ];
    for (const pick of picks) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = t(pick.key, this.host.element());
      Object.assign(btn.style, {
        padding: "3px 8px",
        border: "none",
        background: "transparent",
        borderRadius: "4px",
        cursor: "pointer",
        whiteSpace: "nowrap",
        fontSize: "12px",
        lineHeight: "1.4",
        color: "inherit",
        fontFamily: "inherit",
      } satisfies Partial<CSSStyleDeclaration>);
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        this.hidePasteOptions();
        this.replayPaste(pick.mode);
      });
      bar.append(btn);
    }
    anchor.frame.append(bar);
    // Below the pasted content's last line, clamped into the frame.
    Object.assign(bar.style, {
      left: `${Math.max(0, anchor.left)}px`,
      top: `${anchor.top + anchor.height + 4}px`,
    });
    this.#pasteBar = bar;
    document.addEventListener("mousedown", this.#dismiss, true);
    document.addEventListener("keydown", this.#dismiss, true);
  }

  /** Close the bar and drop the pinned source (also the controller's
   *  teardown — the document-level dismiss listeners die with it). */
  hidePasteOptions(): void {
    document.removeEventListener("mousedown", this.#dismiss, true);
    document.removeEventListener("keydown", this.#dismiss, true);
    this.#pasteBar?.remove();
    this.#pasteBar = undefined;
    this.#pasteSource = undefined;
  }

  /** A click outside the bar or Escape dismisses it (Word keeps the paste
   *  itself — only the options bar goes away). Clicks inside the bar fall
   *  through to the pick buttons. */
  readonly #dismiss = (event: Event): void => {
    if (event instanceof KeyboardEvent && event.key !== "Escape") return;
    if (event instanceof MouseEvent && this.#pasteBar?.contains(event.target as Node)) return;
    this.hidePasteOptions();
  };

  /** The Office Clipboard's session collection (newest first, Word's 24-item
   *  cap) — fed by the bridge's onClipboardCollect, rendered by the pane. */
  #items: { text: string; payload: string | null }[] = [];

  collect(item: { text: string; payload: string | null }): void {
    // A re-copy of the newest item keeps the pane's order stable.
    if (this.#items[0]?.text === item.text) return;
    this.#items.unshift({ ...item });
    if (this.#items.length > 24) this.#items.length = 24;
    this.#syncPane();
  }

  #syncPane(): void {
    const pane = this.host.element().shadowRoot?.querySelector("docen-clipboard-pane");
    if (pane) (pane as unknown as { entries: unknown[] }).entries = [...this.#items];
  }

  #pasteEntry(entry: { text: string; payload: string | null } | null): void {
    const bridge = this.host.bridge();
    const editor = bridge?.activeEditor() ?? this.host.editor();
    if (!editor || !entry) return;
    bridge?.focus();
    if (entry.payload && bridge?.insertSlicePayload(entry.payload)) return;
    editor.commands.insertContent(entry.text);
  }

  readonly onPanePaste = (event: Event): void => {
    this.#pasteEntry((event as CustomEvent<{ text: string; payload: string | null }>).detail);
  };

  /** Word's Paste All — items land in collection order (oldest first). */
  readonly onPanePasteAll = (): void => {
    for (const entry of [...this.#items].reverse()) this.#pasteEntry(entry);
  };

  readonly onPaneClear = (): void => {
    this.#items = [];
  };

  /** A paste-options pick: undo the insertion, then replay the same clipboard
   *  content in the picked form. "match" keeps the block structure (lists,
   *  tables, links) but drops the source's run formatting so the destination's
   *  takes over; "text" inserts the plain text. */
  replayPaste(mode: "source" | "match" | "text"): void {
    const source = this.#pasteSource;
    const bridge = this.host.bridge();
    const editor = bridge?.activeEditor() ?? this.host.editor();
    if (!source || !editor) return;
    editor.commands.undo?.();
    if (mode === "text") {
      editor.commands.insertContent(source.text);
      return;
    }
    if (source.kind === "slice") {
      if (mode === "source" && bridge?.insertSlicePayload(source.raw)) return;
      // A slice replayed destination-matched: strip its run marks.
      try {
        const parsed = JSON.parse(source.raw) as { content?: JSONContent[] };
        if (parsed.content) {
          editor.commands.insertContent(stripRunMarks(parsed.content));
        }
      } catch {
        /* unparsable payload — the undo already restored the prior state */
      }
      return;
    }
    if (source.kind === "rtf") {
      const json = parseRTF(source.raw);
      const content = (json.content ?? []).filter((n) => n.type !== "text" || n.text);
      if (content.length)
        editor.commands.insertContent(mode === "match" ? stripRunMarks(content) : content);
      return;
    }
    const body = new DOMParser().parseFromString(source.raw, "text/html").body;
    const json = parseHTMLBody(body, editor.state.schema);
    const content = (json.content ?? []).filter((n) => n.type !== "text" || n.text);
    if (content.length)
      editor.commands.insertContent(mode === "match" ? stripRunMarks(content) : content);
  }
}
