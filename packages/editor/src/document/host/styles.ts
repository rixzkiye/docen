/**
 * Styles domain — split out of the host element (see document/index.ts): the
 * styles-model reads (gallery/pane/inspector), the Modify/New Style dialog
 * plumbing, and the loaded-styles snapshot. The host keeps thin delegating
 * members so its call sites are unchanged.
 */

import { indexParagraphStyles, type StyleEntry, type StylesOptions } from "@docen/docx";
import { defaultParagraphStyleId, mergeStyleChain, resolveFontName } from "@docen/docx";
import type { Editor } from "@docen/docx/core";

import { t } from "../../ui";
import type { FontDialogPatch } from "../../ui/components/workspace/font-dialog";
import type { ModifyStyleState } from "../../ui/components/workspace/modify-style-dialog";
import type { StyleChoice } from "../../ui/components/workspace/modify-style-dialog";
import type { NewStyleState } from "../../ui/components/workspace/new-style-dialog";
import type { StylesPaneState } from "../components/styles-pane";
import type { StylesInspectorData } from "../components/styles-pane";
import { styleGalleryItems } from "../ribbon";

/** Built-in paragraph/character/table style ids → their localized name keys
 *  (Word's built-in style names follow the UI language). */
export const BUILT_IN_STYLE_KEYS: Readonly<Record<string, string>> = {
  normal: "styleName.normal",
  heading1: "styleName.heading1",
  heading2: "styleName.heading2",
  heading3: "styleName.heading3",
  heading4: "styleName.heading4",
  heading5: "styleName.heading5",
  heading6: "styleName.heading6",
  heading7: "styleName.heading7",
  heading8: "styleName.heading8",
  heading9: "styleName.heading9",
  title: "styleName.title",
  subtitle: "styleName.subtitle",
  quote: "styleName.quote",
  intensequote: "styleName.intenseQuote",
  listparagraph: "styleName.listParagraph",
};

/** The OOXML default display names — a built-in whose explicit name matches
 *  this shows the localized key instead of the raw name. */
export const BUILT_IN_DEFAULT_NAMES: Readonly<Record<string, string>> = {
  normal: "Normal",
  heading1: "heading 1",
  heading2: "heading 2",
  heading3: "heading 3",
  heading4: "heading 4",
  heading5: "heading 5",
  heading6: "heading 6",
  heading7: "heading 7",
  heading8: "heading 8",
  heading9: "heading 9",
  title: "Title",
  subtitle: "Subtitle",
  quote: "Quote",
  intensequote: "Intense Quote",
  listparagraph: "List Paragraph",
};

/** The styles domain's view of the host — only what its bodies touch. */
export interface StylesHostView {
  element(): HTMLElement;
  root(): ShadowRoot | null;
  editor(): Editor | undefined;
  renderChrome(): void;
  fontPatchOfRun(run: Record<string, unknown>): FontDialogPatch;
  fontRunPropsOf(patch: FontDialogPatch): Record<string, unknown>;
}

/**
 * Styles pane / gallery / Modify Style, split out of the host element. Owns
 * the rendered-model caches (`galleryStyles`, `stylesSnapshot`).
 */
export class StylesDomain {
  /** The styles model identity currently rendered into the Styles gallery —
   *  skips the rebuild until the model object is replaced. */
  #galleryStyles?: StylesOptions;
  /** The styles model as the document opened with — the Design tab's style-set
   *  gallery restores it from here ("default" entry). */
  #stylesSnapshot: string | null = null;

  /** Run marks the Style Inspector lists as direct formatting (i18n keys — the
   *  ribbon's command labels double as the formatting names). */
  static readonly MARK_LABELS: Readonly<Record<string, string>> = {
    bold: "ribbon.cmd.bold",
    italic: "ribbon.cmd.italic",
    underline: "ribbon.cmd.underline",
    strike: "ribbon.cmd.strike",
    subscript: "ribbon.cmd.subscript",
    superscript: "ribbon.cmd.superscript",
  };

  constructor(private readonly host: StylesHostView) {}

  /** The loaded document's styles model (doc.attrs.styles), or null. */
  docStyles(editor: Editor): StylesOptions | null {
    return (editor.state.doc.attrs?.styles as StylesOptions | undefined) ?? null;
  }

  /** A style's display name: built-in styles show Word's localized name
   *  (BUILT_IN_STYLE_KEYS), everything else shows the document's own name
   *  (the id as the fallback). */
  styleDisplayName(id: string, name: unknown): string {
    const key = id.toLowerCase();
    const builtin = BUILT_IN_STYLE_KEYS[key];
    // A renamed built-in (explicit name off the OOXML default) shows as-is;
    // otherwise the built-in key localizes.
    const def = BUILT_IN_DEFAULT_NAMES[key];
    if (
      def &&
      typeof name === "string" &&
      name.trim() &&
      name.trim().toLowerCase() !== def.toLowerCase()
    )
      return name.trim();
    if (builtin) return t(builtin, this.host.element());
    return typeof name === "string" && name ? name : id;
  }

  /** The paragraph-style id at the caret (the HeadingLevel literal carried on
   *  `heading` for heading paragraphs, the pStyle id on `style` otherwise). */
  currentStyleId(editor: Editor): string | null {
    const attrs = editor.getAttributes("paragraph") as {
      heading?: unknown;
      style?: unknown;
    };
    if (typeof attrs.heading === "string" && attrs.heading) return attrs.heading;
    if (typeof attrs.style === "string" && attrs.style) return attrs.style;
    return null;
  }

  /** Mirror the paragraph style at the caret into the Styles gallery — its
   *  value is the current paragraph's style id (the HeadingLevel literal
   *  carried on `heading` for heading paragraphs, the pStyle id on `style`
   *  otherwise, or "Normal" when the paragraph carries none). The gallery
   *  outlines the entry matching the value (Word's applied-style card). */
  syncStyleControl(): void {
    const editor = this.host.editor();
    if (!editor) return;
    const value = this.currentStyleId(editor) || "Normal";
    const cb = this.host.root()?.querySelector<HTMLElement>('docen-ribbon-gallery[event="style"]');
    if (cb && cb.getAttribute("value") !== value) cb.setAttribute("value", value);
    // The cards carry each style's effective formatting, so the items must
    // track the document's styles model — the ribbon template bakes one
    // snapshot at build time (usually before the document loads). Rebuild only
    // when the model object is replaced (load / style-set switch / modify
    // style): the identity guard keeps caret-only transactions from
    // recomputing the basedOn merges.
    const styles = this.docStyles(editor);
    if (cb && styles && styles !== this.#galleryStyles) {
      this.#galleryStyles = styles;
      const items = styleGalleryItems(styles).map((item) => {
        const text = this.styleDisplayName(item.value ?? "", item.text);
        // The card renders preview.text (the label inside the card), item.text
        // is the menu/tooltip name — both follow the same display naming.
        return { ...item, text, preview: { ...item.preview, text } };
      });
      cb.setAttribute("items", JSON.stringify(items));
    }
    // The Styles pane's highlight follows the caret's paragraph style.
    const pane = this.host.root()?.querySelector("docen-styles-pane") as
      | (HTMLElement & { setCurrent(id: string): void })
      | null;
    pane?.setCurrent(value);
  }

  /** Build the Styles pane's list from the document's styles model: every
   *  paragraph style (custom + built-in named), each row previewed with the
   *  formatting its basedOn chain merges to. */
  renderStylesPane(): void {
    const pane = this.host.root()?.querySelector("docen-styles-pane") as
      | (HTMLElement & { renderStyles(state: StylesPaneState): void })
      | null;
    const editor = this.host.editor();
    if (!pane || !editor) return;
    const styles = this.docStyles(editor);
    const byId = styles ? indexParagraphStyles(styles) : new Map();
    const entries = [...byId.entries()].map(([id, style]) => {
      const run = mergeStyleChain(byId, id).run as Record<string, unknown>;
      return {
        id,
        name: this.styleDisplayName(id, style.name),
        preview: {
          font: typeof run.font === "string" ? run.font : undefined,
          size: typeof run.size === "number" ? run.size : undefined,
          bold: run.bold === true,
          italic: run.italic === true,
          color: typeof run.color === "string" ? run.color : undefined,
          underline: !!run.underline,
        },
      };
    });
    // A paragraph without heading/pStyle attrs carries the document's default
    // paragraph style (Word highlights "Normal" in that case).
    const currentId =
      this.currentStyleId(editor) ??
      defaultParagraphStyleId(styles) ??
      (byId.has("Normal") ? "Normal" : "");
    pane.renderStyles({ entries, currentId });
  }

  /** Push the selection's style stack to the pane's inspector view: the
   *  paragraph style, the hyperlink character style (the docx editing model
   *  has no other character-style carrier), and the run marks at the
   *  selection as the direct-formatting list. */
  renderStylesInspector(): void {
    const pane = this.host.root()?.querySelector("docen-styles-pane") as
      | (HTMLElement & { renderInspector(data: StylesInspectorData): void })
      | null;
    const editor = this.host.editor();
    if (!pane || !editor) return;
    const styles = this.docStyles(editor);
    const byId = styles ? indexParagraphStyles(styles) : new Map();
    const styleId = this.currentStyleId(editor);
    const marks = editor.state.selection.$from.marks();
    const characterStyle = marks.some((m) => m.type.name === "link")
      ? t("styleName.hyperlink", this.host.element())
      : null;
    const direct: string[] = [];
    for (const mark of marks) {
      const attrs = mark.attrs as Record<string, unknown>;
      if (mark.type.name === "textStyle") {
        const { font, size, color } = attrs;
        if (typeof font === "string")
          direct.push(`${t("fontDialog.font", this.host.element())}: ${font}`);
        if (typeof size === "number")
          direct.push(`${t("fontDialog.size", this.host.element())}: ${size} pt`);
        if (typeof color === "string")
          direct.push(`${t("modifyStyleDialog.color", this.host.element())}: #${color}`);
        continue;
      }
      const label = StylesDomain.MARK_LABELS[mark.type.name];
      if (label) direct.push(t(label, this.host.element()));
    }
    pane.renderInspector({
      paragraphStyle: styleId
        ? this.styleDisplayName(styleId, byId.get(styleId)?.name)
        : t("styleName.normal", this.host.element()),
      characterStyle,
      direct,
    });
  }

  /** Prefill and open the Modify Style dialog for one style. The fields read
   *  the style's OWN definition (not the merged chain) — Word shows what the
   *  style itself says, leaving inherited values blank. */
  openModifyStyle(id: string): void {
    const dialog = this.host.root()?.querySelector("docen-modify-style-dialog") as
      | (HTMLElement & { show(state: ModifyStyleState): void })
      | null;
    const editor = this.host.editor();
    if (!dialog || !editor || !id) return;
    const styles = this.docStyles(editor);
    const byId = styles ? indexParagraphStyles(styles) : new Map();
    const style = byId.get(id);
    const run = (style?.run ?? {}) as Record<string, unknown>;
    const underline = run.underline as { type?: unknown } | undefined;
    // The w:pPr block only exists on the paragraph side of the StyleEntry union.
    const paragraph = (style as { paragraph?: Record<string, unknown> } | undefined)?.paragraph;
    const spacing = (paragraph?.spacing ?? {}) as Record<string, unknown>;
    const indent = (paragraph?.indent ?? {}) as Record<string, unknown>;
    const choices: StyleChoice[] = [...byId.entries()]
      .map(([cid, cs]) => ({ id: cid, name: this.styleDisplayName(cid, cs.name) }))
      .sort((a, b) => a.name.localeCompare(b.name));
    const color = run.color as string | { val?: unknown } | undefined;
    // The line buttons are multiples (240ths of a line); the atLeast/exact
    // rules have no button and stay blank (the Paragraph dialog edits them).
    const lineMultiple = spacing.lineRule == null || spacing.lineRule === "auto";
    const state: ModifyStyleState = {
      id,
      name: this.styleDisplayName(id, style?.name),
      choices,
      basedOn: (style?.basedOn as string | undefined) ?? null,
      next: (style?.next as string | undefined) ?? null,
      font: resolveFontName(run.font),
      size: typeof run.size === "number" ? run.size : null,
      bold: run.bold === true,
      italic: run.italic === true,
      underline: underline?.type != null,
      color: typeof color === "string" ? color : typeof color?.val === "string" ? color.val : null,
      alignment: (paragraph?.alignment as string | undefined) ?? null,
      lineSpacing: lineMultiple && typeof spacing.line === "number" ? spacing.line : null,
      indentLeft: typeof indent.left === "number" ? indent.left : null,
      indentRight: typeof indent.right === "number" ? indent.right : null,
      spacingBefore: typeof spacing.before === "number" ? spacing.before : null,
      spacingAfter: typeof spacing.after === "number" ? spacing.after : null,
      quickFormat: (style as { quickFormat?: boolean } | undefined)?.quickFormat === true,
      autoRedefine: (style as { autoRedefine?: boolean } | undefined)?.autoRedefine === true,
      // The gallery's merged effective run (basedOn chain + docDefaults) is
      // also the preview's formatting — same source, same CSS.
      previewCss: styleGalleryItems(styles).find((item) => item.value === id)?.preview?.css,
    };
    state.description = this.styleDescription(state, byId);
    dialog.show(state);
  }

  /** Prefill and open the New Style dialog. */
  openNewStyle(): void {
    const dialog = this.host.root()?.querySelector("docen-new-style-dialog") as
      | (HTMLElement & { show(state: NewStyleState): void })
      | null;
    const editor = this.host.editor();
    if (!dialog || !editor) return;
    const styles = this.docStyles(editor);
    const byId = styles ? indexParagraphStyles(styles) : new Map();
    const choices: StyleChoice[] = [...byId.entries()]
      .map(([cid, cs]) => ({ id: cid, name: this.styleDisplayName(cid, cs.name) }))
      .sort((a, b) => a.name.localeCompare(b.name));

    let num = 1;
    while ([...byId.values()].some((s) => s.name?.toLowerCase() === `style ${num}`.toLowerCase())) {
      num++;
    }

    dialog.show({
      choices,
      defaultName: `Style ${num}`,
      type: "paragraph",
      basedOn: "Normal",
      next: "Normal",
    });
  }

  /** The description under the preview — the style's own definition read out
   *  the way Word's description box does: the basedOn/next pointers, then the
   *  comma list of the formatting this dialog edits. */
  styleDescription(state: ModifyStyleState, byId: Map<string, StyleEntry>): string {
    const styleName = (sid: string | null): string | null => {
      if (!sid) return null;
      const entry = byId.get(sid);
      return entry ? this.styleDisplayName(sid, entry.name) : sid;
    };
    const lines: string[] = [];
    const basedOn = styleName(state.basedOn);
    const next = styleName(state.next);
    if (basedOn) lines.push(`${t("modifyStyleDialog.basedOn", this.host.element())} ${basedOn}`);
    if (next) lines.push(`${t("modifyStyleDialog.next", this.host.element())} ${next}`);
    const parts: string[] = [];
    if (state.font) parts.push(state.font);
    if (state.size != null) parts.push(`${state.size} ${t("unit.pt", this.host.element())}`);
    if (state.bold) parts.push(t("ribbon.cmd.bold", this.host.element()));
    if (state.italic) parts.push(t("ribbon.cmd.italic", this.host.element()));
    if (state.underline) parts.push(t("ribbon.cmd.underline", this.host.element()));
    if (state.color) parts.push(`#${state.color}`);
    const alignKeys: Record<string, string> = {
      center: "ribbon.cmd.align-center",
      right: "ribbon.cmd.align-right",
      both: "ribbon.cmd.justify",
      left: "ribbon.cmd.align-left",
    };
    if (state.alignment)
      parts.push(t(alignKeys[state.alignment] ?? "ribbon.cmd.align-left", this.host.element()));
    const lineKeys: Record<number, string> = {
      240: "modifyStyleDialog.lineSingle",
      360: "modifyStyleDialog.line15",
      480: "modifyStyleDialog.lineDouble",
    };
    if (state.lineSpacing != null) {
      const key = lineKeys[state.lineSpacing];
      parts.push(
        key
          ? t(key, this.host.element())
          : `${state.lineSpacing / 240} ${t("modifyStyleDialog.lineTimes", this.host.element())}`,
      );
    }
    if (state.spacingBefore != null)
      parts.push(
        `${t("modifyStyleDialog.before", this.host.element())} ${state.spacingBefore / 20} ${t("unit.pt", this.host.element())}`,
      );
    if (state.spacingAfter != null)
      parts.push(
        `${t("modifyStyleDialog.after", this.host.element())} ${state.spacingAfter / 20} ${t("unit.pt", this.host.element())}`,
      );
    if (parts.length) lines.push(parts.join(", "));
    return lines.join("\n");
  }

  /** The Modify Style dialog's Format > Font/Paragraph — open that dialog in
   *  style mode: the fields prefill from the style's own definition and OK
   *  retargets the style (`data-for-style` marker read by the OK routers).
   *  The Modify Style dialog stays open underneath (both are native modal
   *  dialogs; closing the child restores it, mid-edit fields intact). */
  openStyleFormat(id: string, target: "font" | "paragraph"): void {
    const editor = this.host.editor();
    if (!editor || !id) return;
    const styles = this.docStyles(editor);
    const style = styles ? indexParagraphStyles(styles).get(id) : undefined;
    if (target === "paragraph") {
      const dialog = this.host.root()?.querySelector("docen-paragraph-dialog") as
        | (HTMLElement & {
            show(attrs?: Record<string, unknown>, opts?: { styleId?: string }): void;
          })
        | null;
      // StyleEntry unions the paragraph/character shapes — the w:pPr block
      // only exists on the paragraph side.
      const paragraph = (style as { paragraph?: Record<string, unknown> } | undefined)?.paragraph;
      dialog?.show(paragraph ?? {}, { styleId: id });
    } else {
      const dialog = this.host.root()?.querySelector("docen-font-dialog") as
        | (HTMLElement & { show(state: FontDialogPatch, opts?: { styleId?: string }): void })
        | null;
      dialog?.show(this.host.fontPatchOfRun((style?.run ?? {}) as Record<string, unknown>), {
        styleId: id,
      });
    }
  }

  /** Capture the opened styles model (called from #renderDoc). */
  snapshotStyles(): void {
    const editor = this.host.editor();
    const styles = editor ? this.docStyles(editor) : null;
    this.#stylesSnapshot = styles ? JSON.stringify(styles) : null;
  }

  /** Restore the styles model captured at open (the style-set gallery's
   *  "document default" entry) — the same doc-attrs write path as
   *  #applyStylesAttr, so the restore rides undo and re-renders chrome. */
  restoreStylesSnapshot(): void {
    const editor = this.host.editor();
    if (!editor || this.#stylesSnapshot === null) return;
    editor.view.dispatch(
      editor.state.tr.setDocAttribute("styles", JSON.parse(this.#stylesSnapshot)),
    );
    this.host.renderChrome();
  }
}
