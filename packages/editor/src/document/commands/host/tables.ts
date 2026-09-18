import type { Editor } from "@docen/docx/core";

import { tableAncestry } from "../../extensions/commands";
import type { HostCommandDomain } from "./registry";

/** The table domain's view of the host — only what its command bodies touch. */
export interface TablesHostView {
  /** The host element — the shadow-DOM root for the table dialogs. */
  element(): HTMLElement;
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The editor input commands must target (the open story, else the main
   *  editor). */
  activeEditor(): Editor | null | undefined;
  /** The current section's text width (px) — the AutoFit Window payload. */
  contentWidthPx(): number | undefined;
  /** The Draw Border pen state (Table Design → Draw Border). */
  setPenStyle(style: string): void;
  setPenSize(eighths: number): void;
  setPenColor(color: string): void;
  borderPainting(): boolean;
  borderErase(): boolean;
  stopBorderPainting(): void;
  armBorderPainter(erase: boolean): void;
  tableDrawing(): boolean;
  tableEraser(): boolean;
  toggleDrawTable(): void;
  toggleTableEraser(): void;
}

/**
 * Table commands split out of the host element: the Draw Border pen + painter,
 * the insert-table dialog entries, AutoFit Window's layout-derived width, and
 * the Table Properties dialog prefill.
 */
export class TablesHostCommands implements HostCommandDomain {
  constructor(private readonly host: TablesHostView) {}

  readonly chrome: readonly string[] = [
    "pen-style",
    "pen-size",
    "pen-color",
    "border-painter",
    "insert-table",
    "table-dialog",
    "autofit-window",
    "draw-table",
    "table-eraser",
  ];

  readonly editor: readonly string[] = ["table-properties", "insert-quick-table", "insert-excel"];

  run(event: string, value?: string): boolean {
    if (event === "draw-table") {
      this.host.toggleDrawTable();
      return true;
    }
    if (event === "table-eraser") {
      this.host.toggleTableEraser();
      return true;
    }
    // Table Design → Draw Border: the pen pickers stamp the host pen state;
    // the painter split arms the sweep — the face toggles the pen, the
    // drop-down's eraser toggles the erase half (one painter at a time).
    if (event === "pen-style" && typeof value === "string") {
      this.host.setPenStyle(value);
      return true;
    }
    if (event === "pen-size") {
      const size = Number(value);
      if (Number.isFinite(size) && size > 0) this.host.setPenSize(size);
      return true;
    }
    if (event === "pen-color" && typeof value === "string") {
      this.host.setPenColor(value);
      return true;
    }
    if (event === "border-painter") {
      const erase = value === "eraser";
      if (this.host.borderPainting() && this.host.borderErase() === erase) {
        this.host.stopBorderPainting();
      } else {
        this.host.armBorderPainter(erase);
      }
      return true;
    }
    // The Table button's face opens the hover grid; its dropdown's Insert
    // Table opens the classic dialog shape (both insert via table-grid:insert).
    if (event === "insert-table") {
      (
        this.host.element().shadowRoot?.querySelector("docen-table-dialog") as {
          show(m?: string): void;
        } | null
      )?.show("grid");
      return true;
    }
    if (event === "table-dialog") {
      (
        this.host.element().shadowRoot?.querySelector("docen-table-dialog") as {
          show(m?: string): void;
        } | null
      )?.show("form");
      return true;
    }
    // AutoFit Window needs the page's text width — a layout value the command
    // layer can't see, so the host injects it as the twip value (px × 15 at
    // the layout's 96 dpi).
    if (event === "autofit-window") {
      const flow = this.host.contentWidthPx();
      const editor = this.host.editor();
      if (editor && flow != null && flow > 0) {
        (editor.commands as unknown as Record<string, (v?: string) => unknown>)["autofit-window"](
          String(Math.round(flow * 15)),
        );
      }
      return true;
    }
    // Table Properties — open the dialog prefilled from the caret table's
    // attrs; the commit arrives via table-properties:ok
    // (table-properties-apply). No caret table → nothing to show.
    if (event === "table-properties") {
      const target = this.host.activeEditor() ?? this.host.editor();
      const anchor = target ? tableAncestry(target.state) : null;
      const dialog = this.host
        .element()
        .shadowRoot?.querySelector("docen-table-properties-dialog") as {
        show(attrs?: Record<string, unknown>): void;
      } | null;
      if (target && anchor && dialog) {
        dialog.show(
          target.state.selection.$from.node(anchor.tableAt).attrs as Record<string, unknown>,
        );
      }
      return true;
    }
    return false;
  }
}
