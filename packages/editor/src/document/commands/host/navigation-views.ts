import type { Editor } from "@docen/docx/core";

import type { HostCommandDomain } from "./registry";

/** The navigation/views domain's view of the host — only what its command
 *  bodies touch. */
export interface NavigationViewsHostView {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** Toggle a task pane (the navigation/outline pane). */
  togglePane(id: "navigation"): void;
  /** Editing → Find drop-down → Go To (prompts for a page number). */
  goToPage(): void;
  /** Open the nav pane's search box. */
  openSearch(): void;
  /** Open the Find & Replace dialog. */
  openFindReplace(): void;
  /** Current zoom level (percent). */
  zoom(): number;
  setZoom(pct: number): void;
  showZoomDialog(): void;
  zoomPreset(preset: string): void;
  /** Whether the document opened protected (Viewing mode stays protected). */
  docProtected(): boolean;
  /** Re-stamp the tab-row Editing menu to the live editable state. */
  syncEditModeMenu(): void;
  setShowMarks(on: boolean): void;
  getShowMarks(): boolean;
  /** Ruler/gridlines are paint-time view state on the canvas stage. */
  showRuler(): boolean;
  setShowRuler(on: boolean): void;
  showGridlines(): boolean;
  setShowGridlines(on: boolean): void;
  /** Select a document view (Word's View tab buttons). */
  setView(view: string): void;
}

/**
 * Navigation (find/outline/panes), zoom, view-mode, and view-toggle commands
 * split out of the host element.
 */
export class NavigationViewsHostCommands implements HostCommandDomain {
  constructor(private readonly host: NavigationViewsHostView) {}

  readonly chrome: readonly string[] = [
    "toggle-navigation",
    "outline",
    "search",
    "replace",
    "find-dialog",
    "zoom",
    "zoom-100",
  ];

  readonly editor: readonly string[] = [
    "edit-mode",
    "show-marks",
    "toggle-ruler",
    "toggle-gridlines",
    "print-layout",
    "web-layout",
    "read-mode",
    "draft",
  ];

  run(event: string, value?: string): boolean {
    // UI chrome actions are handled locally and need no Tiptap editor.
    if (event === "toggle-navigation") {
      this.host.togglePane("navigation");
      return true;
    }
    // View → Outline: Word's outline view maps to the document-structure
    // pane here (the same tree the navigation pane shows).
    if (event === "outline") {
      this.host.togglePane("navigation");
      return true;
    }
    // Find (ribbon Home → Editing → Find, or Ctrl+F) → open the nav-pane search.
    if (event === "search") {
      // Find drop-down → Go To jumps to a page; the main button and Find
      // open the nav-pane search box.
      if (value === "go-to") this.host.goToPage();
      else this.host.openSearch();
      return true;
    }
    // Replace (ribbon Home → Editing → Replace, or Ctrl+H) → Find & Replace dialog.
    // The Editing group's dialog-box launcher opens the same dialog (Word).
    if (event === "replace" || event === "find-dialog") {
      this.host.openFindReplace();
      return true;
    }
    // Zoom is a canvas action (not a Tiptap command): step in, or apply a
    // preset from the split menu (200/100/75/50/page-width); the split's
    // main button sets 100%.
    if (event === "zoom") {
      this.host.setZoom(this.host.zoom() + 10);
      return true;
    }
    if (event === "zoom-100") {
      if (value === "zoom-dialog") this.host.showZoomDialog();
      else if (value) this.host.zoomPreset(value);
      else this.host.setZoom(100);
      return true;
    }
    const editor = this.host.editor();
    if (!editor) return false;
    // Edit / View mode — toggle the editor's editable state (tab-row "Editing"
    // menu); then re-stamp the menu so its label + checked item follow. A
    // read-only protected document stays protected in Editing mode.
    if (event === "edit-mode") {
      editor.setEditable(value !== "view" && !this.host.docProtected());
      this.host.syncEditModeMenu();
      return true;
    }
    // Formatting marks toggle — canvas-side marks are a later milestone; the
    // host [show-marks] attribute stays the source of truth.
    if (event === "show-marks") {
      this.host.setShowMarks(!this.host.getShowMarks());
      return true;
    }
    // View toggles — ruler and gridlines are paint-time view state (never in
    // the document), so the stage flips the flag and repaints.
    if (event === "toggle-ruler") {
      this.host.setShowRuler(!this.host.showRuler());
      return true;
    }
    if (event === "toggle-gridlines") {
      this.host.setShowGridlines(!this.host.showGridlines());
      return true;
    }
    // View → the four view buttons (Word's View tab): each selects a document
    // view through the `view` attribute — #applyView restages the render.
    const viewOf: Record<string, string> = {
      "print-layout": "print",
      "web-layout": "web",
      "read-mode": "read",
      draft: "draft",
    };
    if (viewOf[event]) {
      this.host.setView(viewOf[event]);
      return true;
    }
    return false;
  }
}
