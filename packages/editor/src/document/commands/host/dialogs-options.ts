import {
  defaultParagraphStyleId,
  indexParagraphStyles,
  mergeStyleChain,
  type StylesOptions,
} from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import type { EditorState } from "@tiptap/pm/state";

import type { FontDialogPatch } from "../../../ui/components/workspace/font-dialog";
import type { HostCommandDomain } from "./registry";

/** Scalar paragraph properties the Paragraph dialog consumes — the slots the
 *  prefill cascade picks up from the style chain and docDefaults when the
 *  paragraph's own attrs leave them unset. Mirrors the dialog's patch keys. */
const PARAGRAPH_FLAG_KEYS = [
  "alignment",
  "outlineLevel",
  "mirrorIndents",
  "adjustRightInd",
  "snapToGrid",
  "contextualSpacing",
  "widowControl",
  "keepNext",
  "keepLines",
  "pageBreakBefore",
  "suppressLineNumbers",
  "suppressAutoHyphens",
  "kinsoku",
  "wordWrap",
  "overflowPunct",
  "autoSpaceDE",
  "autoSpaceDN",
  "textAlignment",
] as const;

/** The dialog domain's view of the host — only what its command bodies touch. */
export interface DialogsOptionsHostView {
  /** The host element — the shadow-DOM root for every dialog lookup. */
  element(): HTMLElement;
  /** The editor input commands must target (the open story, else the main
   *  editor). */
  activeEditor(): Editor | null | undefined;
  /** The document's styles model (Paragraph dialog prefill cascade). */
  docStyles(editor: Editor): StylesOptions | null;
  /** The selection's effective run state (Font dialog prefill). */
  runState(state: EditorState): FontDialogPatch;
  /** Open the Chart Data grid for the selected chart. */
  chartEditAtSelection(): void;
  /** Open the phonetic-guide (拼音指南) dialog for the selection. */
  phoneticOpen(): void;
  /** Open the two-lines-in-one (中文版式) dialog for the selection. */
  twoInOneOpen(): void;
  /** Open the Define New Multilevel List dialog. */
  defineListOpen(): void;
}

/**
 * Dialog-opening commands split out of the host element: Symbol, the
 * Paragraph/Font dialog prefills, Date & Time, Chart Edit Data, and the
 * Chinese-layout + multilevel-list dialogs.
 */
export class DialogsOptionsHostCommands implements HostCommandDomain {
  constructor(private readonly host: DialogsOptionsHostView) {}

  readonly chrome: readonly string[] = [];

  readonly editor: readonly string[] = [
    "symbol",
    "paragraph-dialog",
    "font-dialog",
    "date-time",
    "chart-edit-data",
    "phonetic-guide",
    "two-lines-in-one",
    "define-new-list",
  ];

  run(event: string, _value?: string): boolean {
    // Symbol — open the character grid dialog; the insertion arrives via the
    // dialog's symbol:insert event (it stays open for several inserts).
    if (event === "symbol") {
      (
        this.host.element().shadowRoot?.querySelector("docen-symbol-dialog") as {
          show(): void;
        } | null
      )?.show();
      return true;
    }
    // Paragraph — open the dialog prefilled from the caret paragraph's attrs;
    // the commit arrives via paragraph:ok (stamped by paragraph-dialog-apply).
    if (event === "paragraph-dialog") {
      const target = this.host.activeEditor();
      const node = target?.state.selection.$from.parent;
      if (target && node?.type.name === "paragraph") {
        // Effective values (direct attrs over the style chain over docDefaults
        // — Word's cascade) drive the prefill: the unit boxes are never empty,
        // and a blind commit must not overwrite an inherited indent or spacing
        // (docDefaults' 8pt after) with an explicit 0.
        const styles = this.host.docStyles(target);
        const byId = styles ? indexParagraphStyles(styles) : new Map();
        const attrs = node.attrs as Record<string, unknown>;
        const chain = mergeStyleChain(
          byId,
          (typeof attrs.style === "string" && attrs.style) || defaultParagraphStyleId(styles),
        ).paragraph;
        const ddParagraph = ((
          styles?.default as { document?: { paragraph?: Record<string, unknown> } } | undefined
        )?.document?.paragraph ?? {}) as Record<string, unknown>;
        const effective = { ...attrs };
        for (const key of ["indent", "spacing"] as const) {
          const inherited = {
            ...(typeof ddParagraph[key] === "object" && ddParagraph[key] ? ddParagraph[key] : {}),
            ...(typeof chain[key] === "object" && chain[key] ? (chain[key] as object) : {}),
          };
          const direct = attrs[key];
          effective[key] = Object.keys(inherited).length
            ? {
                ...inherited,
                ...(typeof direct === "object" && direct ? direct : {}),
              }
            : direct;
        }
        // Scalar paragraph flags cascade the same way: a null attrs slot is the
        // schema's "unset", so an explicit value on the style chain or in
        // docDefaults must win over the dialog's spec-default fallback.
        for (const key of PARAGRAPH_FLAG_KEYS) {
          const direct = attrs[key];
          effective[key] = direct ?? chain?.[key] ?? ddParagraph[key];
        }
        (
          this.host.element().shadowRoot?.querySelector("docen-paragraph-dialog") as {
            show(attrs?: Record<string, unknown>): void;
          } | null
        )?.show(effective);
      }
      return true;
    }
    // Font — open the dialog prefilled from the selection's run marks; the
    // commit arrives via font:ok (#onFontDialogOk).
    if (event === "font-dialog") {
      const target = this.host.activeEditor();
      const dialog = this.host.element().shadowRoot?.querySelector("docen-font-dialog") as {
        show(state: FontDialogPatch): void;
      } | null;
      if (target && dialog) dialog.show(this.host.runState(target.state));
      return true;
    }
    // Date & Time — open the dialog; the commit arrives via date-time:insert
    // (static text, or a DATE field when "update automatically" is checked).
    if (event === "date-time") {
      (
        this.host.element().shadowRoot?.querySelector("docen-date-time-dialog") as {
          show(): void;
        } | null
      )?.show();
      return true;
    }
    // Chart — open the Edit Data grid (Chart Design tab → Data group); the
    // commit arrives via chart:ok → the chart-data-apply command.
    if (event === "chart-edit-data") {
      this.host.chartEditAtSelection();
      return true;
    }
    // Phonetic guide (拼音指南, Home → Font): the per-character reading
    // dialog over the selection.
    if (event === "phonetic-guide") {
      this.host.phoneticOpen();
      return true;
    }
    // Chinese Layout (中文版式, Home → Paragraph): the two-lines-in-one
    // dialog over the selection (合并字符 rides the same dialog).
    if (event === "two-lines-in-one") {
      this.host.twoInOneOpen();
      return true;
    }
    // Multilevel List gallery (Home → Paragraph): the Define New Multilevel
    // List dialog — its last entry.
    if (event === "define-new-list") {
      this.host.defineListOpen();
      return true;
    }
    return false;
  }
}
