/**
 * Insert / jump commands — split out of the host element (see
 * document/index.ts): Symbols, Shapes, WordArt, blank/cover pages, date-time,
 * TOC, text from file, equations, tab stops, bookmarks, and the note jumps.
 * Pure command bodies over the editor; the host keeps thin delegates.
 */

import type { JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import type { ProjectedFlowBox } from "@docen/docx/layout";
import type { FlowPage } from "@docen/layout";
import { EMU_PER_PX, twipToPx } from "@docen/layout";
import { EditorState } from "@tiptap/pm/state";
import { NodeSelection } from "@tiptap/pm/state";

import { t } from "../../ui";
import type { FontDialogPatch } from "../../ui/components/workspace/font-dialog";
import type { DocenHyphenationDialog } from "../../ui/components/workspace/hyphenation-dialog";
import type { HyphenationDialogOptions } from "../../ui/components/workspace/hyphenation-dialog";
import type { LinkValues } from "../../ui/components/workspace/link-dialog";
import type { DocenTabsDialog } from "../../ui/components/workspace/tabs-dialog";
import type { EditBridge } from "../canvas/edit-bridge";
import { equationSeed } from "../commands/equation";
import type { MailMergeCommands } from "../commands/mail-merge";
import { applyRecipientsRow } from "../commands/mail-merge";
import { mathAtomAt } from "../math-atom";

/** The insert domain's view of the host — only what its bodies touch. */
export interface InsertHostView {
  element(): HTMLElement;
  root(): ShadowRoot | null;
  editor(): Editor | undefined;
  bridge(): EditBridge | undefined;
  flow(): ProjectedFlowBox | undefined;
  pages(): readonly FlowPage[];
  hyphenation(): {
    auto?: boolean;
    doNotHyphenateCaps?: boolean;
    zoneTw?: number;
    limit?: number;
  };
  setHyphenation(value: {
    auto?: boolean;
    doNotHyphenateCaps?: boolean;
    zoneTw?: number;
    limit?: number;
  }): void;
  merge(): MailMergeCommands;
  getJSON(): JSONContent;
  setJSON(json: JSONContent): void;
  renderDoc(doc: JSONContent): void;
  setTextSelection(from: number, to?: number): void;
  openBookmarkDialog(): void;
}

/**
 * Insert/format command bodies, split out of the host element. The host keeps
 * its private command members as thin delegations.
 */
export class InsertDomain {
  constructor(private readonly host: InsertHostView) {}

  /** Insert → Text Box / Shapes: a standalone wps shape run, floating
   *  wrap-none. Without a rect (Text Box, and the drawer's landing spot for
   *  a bare click it cannot resolve) the shape centers on the page at Word's
   *  2" × 1.2" default; a draw rect (page-local px) fixes both. The text box
   *  carries Word's plain look — white fill, accent-1 hairline — and an
   *  editable empty body (the PM `content`); a gallery shape carries its
   *  preset geometry with the accent fill instead. */
  insertShapeAt(
    preset: string | undefined,
    rect?: {
      page: number;
      x: number;
      y: number;
      w: number;
      h: number;
      flipH?: boolean;
      flipV?: boolean;
    },
  ): void {
    // Insert into the story the caret lives in — a header/footer story must
    // receive the shape, not the stale main-doc selection behind it.
    const editor = this.host.bridge()?.activeEditor() ?? this.host.editor();
    if (!editor) return;
    const geometry: Record<string, unknown> = rect
      ? {
          transformation: {
            width: Math.max(1, Math.round(rect.w * EMU_PER_PX)),
            height: Math.max(1, Math.round(rect.h * EMU_PER_PX)),
            // A line drawn right-to-left / bottom-to-top mirrors its diagonal
            // (the preset path always runs corner to corner, top-left first).
            ...(rect.flipH ? { flipHorizontal: true } : {}),
            ...(rect.flipV ? { flipVertical: true } : {}),
          },
          floating: {
            horizontalPosition: { relative: "page", offset: Math.round(rect.x * EMU_PER_PX) },
            verticalPosition: { relative: "page", offset: Math.round(rect.y * EMU_PER_PX) },
            wrap: { type: "none" },
          },
        }
      : {
          // Word's plain text box default: 2" × 1.2".
          transformation: { width: 1828800, height: 1097280 },
          floating: {
            horizontalPosition: { relative: "page", align: "center" },
            verticalPosition: { relative: "page", align: "center" },
            wrap: { type: "none" },
          },
        };
    if (preset) {
      geometry.geometry = preset;
      // The theme's accent-1 pair (fill + its darkened outline) — the same
      // look Word gives a fresh shape; the projection paints flat hex.
      geometry.fill = { type: "solid", color: "4472C4" };
      geometry.outline = { color: "2F528F", width: 12700 };
    } else {
      geometry.fill = { type: "solid", color: "FFFFFF" };
      geometry.outline = { color: "4472C4", width: 12700 };
    }
    editor.commands.insertContentAt(editor.state.selection.from, {
      type: "wpsShape",
      attrs: { wpsShape: geometry },
      content: [{ type: "paragraph" }],
    } as JSONContent);
  }

  /** WordArt — a centered text box whose single run carries the preset look
   *  (large, bold, theme accent); Word 2013+ models WordArt the same way. */
  insertWordArt(): void {
    const editor = this.host.bridge()?.activeEditor() ?? this.host.editor();
    if (!editor) return;
    editor.commands.insertContentAt(editor.state.selection.from, {
      type: "wpsShape",
      attrs: {
        wpsShape: {
          transformation: { width: 3657600, height: 914400 },
          fill: { type: "solid", color: "FFFFFF" },
          outline: { color: "4472C4", width: 12700 },
          floating: {
            horizontalPosition: { relative: "page", align: "center" },
            verticalPosition: { relative: "page", align: "center" },
            wrap: { type: "none" },
          },
        },
      },
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "text",
              text: t("wordArt.placeholder", this.host.element()),
              marks: [{ type: "textStyle", attrs: { size: 48, bold: true, color: "4472C4" } }],
            },
          ],
        },
      ],
    } as JSONContent);
  }

  /** Blank Page — two page breaks at the caret: the rest of the current page
   *  stays empty and a full empty page follows (Word's Blank Page). */
  insertBlankPage(): void {
    const editor = this.host.bridge()?.activeEditor() ?? this.host.editor();
    if (!editor) return;
    editor
      .chain()
      .insertContentAt(editor.state.selection.from, { type: "pageBreak" } as JSONContent)
      .insertContentAt(editor.state.selection.from, { type: "pageBreak" } as JSONContent)
      .run();
  }

  /** Cover Page — a title block at the document start (title/subtitle/author/
   *  company/date, centered and oversized) followed by a page break. */
  insertCoverPage(): void {
    const editor = this.host.bridge()?.activeEditor() ?? this.host.editor();
    if (!editor) return;
    const centered = (text: string, attrs: Record<string, unknown>): JSONContent =>
      ({
        type: "paragraph",
        attrs: { alignment: "center" },
        content: text ? [{ type: "text", text, marks: [{ type: "textStyle", attrs }] }] : undefined,
      }) as JSONContent;
    editor.commands.insertContentAt(1, [
      { type: "paragraph" } as JSONContent,
      centered(t("coverPage.title", this.host.element()), {
        size: 56,
        bold: true,
        color: "2E74B5",
      }),
      centered(t("coverPage.subtitle", this.host.element()), { size: 28, color: "595959" }),
      { type: "paragraph" } as JSONContent,
      centered(t("coverPage.author", this.host.element()), { size: 24 }),
      centered(t("coverPage.company", this.host.element()), { size: 22, color: "595959" }),
      centered(
        new Intl.DateTimeFormat(undefined, {
          year: "numeric",
          month: "long",
          day: "numeric",
        }).format(new Date()),
        { size: 22 },
      ),
      { type: "pageBreak" } as JSONContent,
    ]);
  }

  /** Date and Time — static formatted text, or a DATE field ("update
   *  automatically") whose cached result renders on canvas and refreshes
   *  when Word updates fields. */
  insertDateTime(detail: { text: string; instruction?: string }): void {
    const editor = this.host.bridge()?.activeEditor() ?? this.host.editor();
    if (!editor || !detail?.text) return;
    const node: JSONContent = detail.instruction
      ? ({
          type: "inlinePassthrough",
          attrs: {
            data: JSON.stringify({
              simpleField: { instruction: detail.instruction, cachedValue: detail.text },
            }),
          },
        } as JSONContent)
      : ({ type: "text", text: detail.text } as JSONContent);
    editor.commands.insertContentAt(editor.state.selection.from, node);
  }

  /** Custom Table of Contents — run the toc command with the dialog's picks,
   *  then the same repaginate-and-update pass the plain toc uses. */
  insertCustomToc(detail: {
    headingRange: string;
    leader: string;
    showPageNumbers: boolean;
    alignPageNumbers: boolean;
    hyperlink?: boolean;
    styles?: string;
  }): void {
    const editor = this.host.bridge()?.activeEditor() ?? this.host.editor();
    if (!editor) return;
    const pageOf = (pos: number): number | null => {
      const page = this.host.bridge()?.pageOf(pos);
      return typeof page === "number" ? page + 1 : null;
    };
    const tabPositionTw = this.host.flow()
      ? Math.round(this.host.flow()!.contentWidthPx / twipToPx(1))
      : undefined;
    if (editor.commands.toc(pageOf, tabPositionTw, detail)) {
      requestAnimationFrame(() =>
        requestAnimationFrame(() => editor.commands["update-toc"](pageOf, tabPositionTw)),
      );
    }
  }

  /** Object → Text from File — read a plain-text file in at the caret, one
   *  paragraph per line (Word's Insert File). */
  insertFileText(): void {
    const input = this.host.root()?.querySelector<HTMLInputElement>("#text-input");
    if (!input) return;
    input.onchange = () => {
      const file = input.files?.[0];
      input.value = "";
      if (!file) return;
      void file.text().then((text) => {
        const editor = this.host.bridge()?.activeEditor() ?? this.host.editor();
        if (!editor || !text) return;
        const paragraphs = text.split(/\r\n|\n|\r/).map(
          (line) =>
            ({
              type: "paragraph",
              content: line ? [{ type: "text", text: line }] : undefined,
            }) as JSONContent,
        );
        editor.commands.insertContentAt(editor.state.selection.from, paragraphs);
      });
    };
    input.click();
  }

  /** Insert → Online Pictures: insert the dialog's embedded picture at the
   *  caret as a normal image node. The natural size clamps to the content
   *  width (Word inserts at natural size but never wider than the frame,
   *  keeping the aspect); a missing size leaves the attrs unset — the image
   *  painter falls back to its default box and the DOCX prepare step fills
   *  the dimensions from the image header on export. */
  insertOnlinePicture(picture: {
    src: string;
    width?: number;
    height?: number;
    alt?: string;
  }): void {
    const editor = this.host.bridge()?.activeEditor() ?? this.host.editor();
    if (!editor || !picture.src) return;
    const contentW = this.host.flow()?.contentWidthPx ?? 620;
    const naturalW = picture.width && picture.width > 0 ? picture.width : undefined;
    const naturalH = picture.height && picture.height > 0 ? picture.height : undefined;
    const scale = naturalW ? Math.min(1, contentW / naturalW) : 1;
    editor.commands.insertContent({
      type: "image",
      attrs: {
        src: picture.src,
        ...(picture.alt ? { alt: picture.alt } : {}),
        ...(naturalW != null ? { width: Math.round(naturalW * scale) } : {}),
        ...(naturalH != null ? { height: Math.round(naturalH * scale) } : {}),
      },
    });
    this.host.bridge()?.focus();
  }

  /** Insert → Equation — drop one placeholder template (fraction / script /
   *  radical / sum / integral) at the caret as a mathInline atom (Word's
   *  Insert → Symbols → Equation gallery). Each argument is an empty run —
   *  the □ slot; the radical's absent degree reads as the square root
   *  (degHide follows). Round-trips through DOCX via the node's
   *  parseDocxInline/compile pair; the projection paints the structured
   *  elements (or the placeholder box when a shape has none). */
  insertEquation(template: string): void {
    const editor = this.host.editor();
    if (!editor) return;
    const seed = equationSeed(template);
    if (!seed) return;
    const node = editor.schema.nodeFromJSON(seed);
    // One transaction: insert the atom and wrap it in a NodeSelection — the
    // selection rides the new atom, keeping the equation context tab alive
    // (a bare insert leaves the caret behind it and the tab would blink off).
    const pos = editor.state.selection.from;
    const tr = editor.state.tr.insert(pos, node);
    tr.setSelection(NodeSelection.create(tr.doc, pos));
    editor.view.dispatch(tr.scrollIntoView());
  }

  /** The equation context tab's symbol grid: drop the glyph at the caret.
   *  When the selection rides the math atom, hand the selection back to it
   *  afterwards — the atom shifts by the inserted glyph — so the tab stays
   *  up and symbols can be typed in a run. (insertContent would replace the
   *  NodeSelection's whole atom, so this goes through the transaction.) */
  insertEquationSymbol(char: string): void {
    const editor = this.host.editor();
    if (!editor || !char) return;
    const touching = mathAtomAt(editor.state);
    if (!touching) {
      editor.commands.insertContent(char);
      return;
    }
    // A caret inserts at its edge; a NodeSelection appends after the atom
    // (the formula stays put). The atom shifts by the glyph when the insert
    // lands before it — aim the selection back accordingly.
    const onAtom = editor.state.selection instanceof NodeSelection;
    const insertAt = onAtom ? touching.pos + touching.node.nodeSize : editor.state.selection.from;
    const backAt = insertAt <= touching.pos ? touching.pos + char.length : touching.pos;
    const tr = editor.state.tr.insertText(
      char,
      insertAt,
      onAtom ? insertAt : editor.state.selection.to,
    );
    const $back = tr.doc.resolve(backAt);
    if ($back.nodeAfter?.type.name === "inlinePassthrough") {
      tr.setSelection(NodeSelection.create(tr.doc, backAt));
    }
    editor.view.dispatch(tr.scrollIntoView());
  }

  insertSoftHyphen(): void {
    this.host.bridge()?.focus();
    this.host.editor()?.commands.insertContent("\u00AD");
  }

  /** Insert Bookmark: open the Bookmark dialog. */
  insertBookmark(): void {
    this.host.openBookmarkDialog();
  }

  /** Ctrl+Click / Open Hyperlink on a `#name` link — place the caret past the
   *  matching bookmarkStart atom and scroll it into view (Word scrolls to the
   *  bookmark). No matching bookmark is a no-op. */
  jumpToBookmark(name: string): void {
    const editor = this.host.editor();
    if (!editor) return;
    let target: number | null = null;
    editor.state.doc.descendants((child, pos) => {
      if (target != null || child.type.name !== "inlinePassthrough") return;
      try {
        const data = JSON.parse(String(child.attrs?.data ?? "{}")) as {
          bookmarkStart?: { name?: string };
        };
        if (data.bookmarkStart?.name === name) target = pos + child.nodeSize;
      } catch {
        // opaque verbatim blob — not a bookmark
      }
    });
    if (target == null) return;
    this.host.setTextSelection(target);
    this.host.bridge()?.scrollIntoView(target);
  }

  /** References → Next Footnote: place the caret on the next
   *  footnote/endnote reference after the selection (document order); none is
   *  a no-op (Word steps through its notes without wrapping). */
  jumpNextNote(): void {
    const editor = this.host.editor();
    if (!editor) return;
    const { from } = editor.state.selection;
    let target: number | null = null;
    editor.state.doc.descendants((child, pos) => {
      if (target != null) return false;
      if (pos <= from || child.type.name !== "inlinePassthrough") return;
      try {
        const data = JSON.parse(String(child.attrs?.data ?? "{}")) as Record<string, unknown>;
        if ("footnoteReference" in data || "endnoteReference" in data) target = pos;
      } catch {
        // opaque verbatim blob — not a note reference
      }
    });
    // After the leaf atom (pos is its left edge) so the caret sits past it.
    if (target != null) {
      this.host.setTextSelection(target + 1);
      this.host.bridge()?.scrollIntoView(target + 1);
    }
  }

  /** References → Previous Footnote: place the caret on the previous
   *  footnote/endnote reference before the selection (document order). */
  jumpPreviousNote(): void {
    const editor = this.host.editor();
    if (!editor) return;
    const { from } = editor.state.selection;
    let target: number | null = null;
    editor.state.doc.descendants((child, pos) => {
      if (pos >= from || child.type.name !== "inlinePassthrough") return;
      try {
        const data = JSON.parse(String(child.attrs?.data ?? "{}")) as Record<string, unknown>;
        if ("footnoteReference" in data || "endnoteReference" in data) target = pos;
      } catch {
        // opaque verbatim blob — not a note reference
      }
    });
    if (target != null) {
      this.host.setTextSelection(target + 1);
      this.host.bridge()?.scrollIntoView(target + 1);
    }
  }

  openHyphenationOptions(): void {
    const dialog = this.host
      .root()
      ?.querySelector<DocenHyphenationDialog>("docen-hyphenation-dialog");
    dialog?.show(this.host.hyphenation());
  }

  setHyphenation(mode: "none" | "auto" | "manual"): void {
    if (mode === "none") {
      this.host.setHyphenation({ ...this.host.hyphenation(), auto: false });
      this.host.renderDoc(this.host.getJSON());
    } else if (mode === "auto") {
      this.host.setHyphenation({ ...this.host.hyphenation(), auto: true });
      this.host.renderDoc(this.host.getJSON());
    } else if (mode === "manual") {
      this.openHyphenationOptions();
    }
  }

  /** Insert → Link / Ctrl+K / right-click Edit Link: open the hyperlink dialog
   *  prefilled from the selection — its text and the link mark riding it (a
   *  caret inside a link edits the whole one via extendMarkRange at commit). */
  insertLink(): void {
    const editor = this.host.bridge()?.activeEditor() ?? this.host.editor();
    if (!editor) return;
    const { empty, from, to } = editor.state.selection;
    (
      this.host.root()?.querySelector("docen-link-dialog") as {
        show(values?: Partial<LinkValues>): void;
      } | null
    )?.show({
      text: empty ? "" : editor.state.doc.textBetween(from, to, " "),
      href: editor.getAttributes("link").href as string | undefined,
    });
  }

  /** The render's preview view — the document JSON with every merge field's
   *  chevron swapped to the previewed recipient's value (identity when the
   *  preview is off). Runs before compile so measure and paint agree. */
  mergedView(doc: JSONContent): JSONContent {
    const row = this.host.merge().previewRow();
    if (row === null) return doc;
    const recipients = this.host.merge().recipients();
    return recipients ? applyRecipientsRow(doc, recipients, row) : doc;
  }

  runStateOf(state: EditorState): FontDialogPatch {
    const seen = new Map<string, Record<string, unknown>>();
    state.doc.nodesBetween(state.selection.from, state.selection.to, (node) => {
      if (seen.size > 0) return false;
      if (!node.isText) return true;
      for (const m of node.marks)
        if (!seen.has(m.type.name)) seen.set(m.type.name, m.attrs as Record<string, unknown>);
      return false;
    });
    const ts = seen.get("textStyle") ?? {};
    const um = seen.get("underline") as
      | { style?: string | null; color?: string | null }
      | undefined;
    const tsU = ts.underline as { type?: string; color?: string } | undefined;
    const str = (v: unknown): string | null => (typeof v === "string" && v ? v : null);
    return {
      font: str(ts.font),
      size: typeof ts.size === "number" || typeof ts.size === "string" ? String(ts.size) : null,
      bold: seen.has("bold") || ts.bold === true,
      italic: seen.has("italic") || ts.italic === true,
      underlineStyle: um
        ? (str(um.style) ?? "single")
        : tsU && tsU.type && tsU.type !== "none"
          ? tsU.type
          : null,
      underlineColor: um ? str(um.color) : str(tsU?.color),
      strike: seen.has("strike") || ts.strike === true,
      doubleStrike: ts.doubleStrike === true,
      superscript: seen.has("superscript"),
      subscript: seen.has("subscript"),
      smallCaps: ts.smallCaps === true,
      allCaps: ts.allCaps === true,
      hidden: ts.vanish === true,
      shadow: ts.shadow === true || Boolean(ts.shadow),
      outline: ts.outline === true || Boolean(ts.outline),
      emboss: ts.emboss === true,
      imprint: ts.imprint === true,
    };
  }

  openTabsDialog = (): void => {
    const dialog = this.host.root()?.querySelector<DocenTabsDialog>("docen-tabs-dialog");
    if (!dialog) return;
    const state = this.host.editor()?.state;
    let tabStops: Array<{
      position: number;
      type: "left" | "center" | "right" | "decimal" | "bar";
      leader?: "dot" | "heavy" | "hyphen" | "middleDot" | "underscore";
    }> = [];
    if (state) {
      state.doc.nodesBetween(state.selection.from, state.selection.to, (node) => {
        if (node.type.name === "paragraph" && Array.isArray(node.attrs.tabStops)) {
          tabStops = node.attrs.tabStops;
          return false;
        }
      });
    }
    const doc = this.host.getJSON() as { settings?: { defaultTabStop?: number } } | null;
    const defaultTabStop = doc?.settings?.defaultTabStop ?? 720;
    dialog.show({ tabStops, defaultTabStop });
  };

  onTabsOk = (
    event: CustomEvent<{
      tabStops?: Array<{ position: number; type?: string; leader?: string }>;
      defaultTabStop?: number;
    }>,
  ): void => {
    const { tabStops, defaultTabStop } = event.detail ?? {};
    if (tabStops !== undefined) {
      this.host.editor()?.commands["set-paragraph-tabs"](tabStops);
    }
    if (defaultTabStop !== undefined) {
      const doc = this.host.getJSON() as { settings?: Record<string, unknown> } | null;
      if (doc) {
        doc.settings = { ...doc.settings, defaultTabStop };
        this.host.setJSON(doc);
      }
    }
    this.host.bridge()?.focus();
  };

  onHyphenationOk = (event: CustomEvent<HyphenationDialogOptions>): void => {
    if (!event.detail) return;
    this.host.setHyphenation({
      auto: event.detail.auto,
      doNotHyphenateCaps: event.detail.doNotHyphenateCaps,
      zoneTw: event.detail.zoneTw,
      limit: event.detail.limit,
    });
    this.host.renderDoc(this.host.getJSON());
  };
}
