import { pinImage } from "@docen/core";
import type { Editor } from "@docen/docx/core";
import { NodeSelection } from "@tiptap/pm/state";

import type { DrawingPropertiesState } from "../../../ui/components/workspace/drawing-properties-dialog";
import type { HostCommandDomain } from "./registry";

/** The drawing/pictures domain's view of the host — only what its command
 *  bodies touch. */
export interface DrawingPicturesHostView {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The editor input commands must target (the open story, else the main
   *  editor). */
  activeEditor(): Editor | null | undefined;
  /** The host element — the shadow-DOM root for the drawing dialog. */
  element(): HTMLElement;
  /** Open the Compress Pictures dialog. */
  showCompressPictures(): void;
  /** Arm the Set Transparent Color eyedropper. */
  armTransparentPick(): void;
  /** The multi-selection members (primary + Shift+Click set). */
  drawingMulti():
    | { pos: number; box: { x: number; y: number; width: number; height: number } }[]
    | null
    | undefined;
  /** Open the picture file picker (Insert → Picture). */
  pickImage(): void;
  /** Open the Online Pictures address dialog (Insert → Pictures). */
  openOnlinePictures(): void;
  /** Open the change-picture file picker (Picture Format → Adjust). */
  pickPicture(): void;
  /** Hand the keyboard back to the editing surface. */
  focusBridge(): void;
  /** The selected floating drawing's dialog state (null when not one). */
  drawingState(): DrawingPropertiesState | null;
  /** Enter the bridge's crop mode on the selected image. */
  enterCropMode(): void;
  /** Enter edit points mode on the selected shape. */
  enterEditPointsMode?(): void;
  /** Insert a centered floating text box (no preset) or a shape preset. */
  insertShapeAt(preset?: string): void;
  /** Arm the drag-to-draw shape drawer with a preset. */
  armShapeDrawer(preset: string): void;
  /** Insert the preset-styled WordArt text box. */
  insertWordArt(): void;
}

/**
 * Drawing and picture commands split out of the host element: Compress
 * Pictures, Set Transparent Color, Group/Distribute, the picture file
 * pickers, Reset Picture and Size, Size & Position, Crop, Text Box, Shapes,
 * and WordArt.
 */
export class DrawingPicturesHostCommands implements HostCommandDomain {
  constructor(private readonly host: DrawingPicturesHostView) {}

  readonly chrome: readonly string[] = [
    "compress-pictures",
    "picture-transparent-pick",
    "drawing-group",
    "drawing-distribute",
    "align-objects",
  ];

  readonly editor: readonly string[] = [
    "insert-picture",
    "online-picture",
    "change-picture",
    "reset-picture-size",
    "drawing-properties",
    "drawing-crop",
    "edit-shape",
    "shape-edit-points",
    "text-box",
    "shapes",
    "wordart",
  ];

  run(event: string, value?: string): boolean {
    // Compress Pictures (Picture Format → Adjust) → the compression dialog.
    if (event === "compress-pictures") {
      this.host.showCompressPictures();
      return true;
    }
    // Set Transparent Color (Picture Format → Color menu): arm the canvas
    // eyedropper — the next press on a picture samples its pixel.
    if (event === "picture-transparent-pick") {
      this.host.armTransparentPick();
      return true;
    }
    // Word's Group / Distribute / Align act on the drawing multi-selection — the
    // ribbon event carries no members, so the bridge's Shift+Click set (the
    // primary plus the toggled members) assembles the payload here. Ungroup
    // needs no payload and rides the wired command directly.
    if (event === "drawing-group" || event === "drawing-distribute" || event === "align-objects") {
      const editor = this.host.activeEditor();
      const members = this.host.drawingMulti();
      if (editor && members && members.length >= 2) {
        const payload = JSON.stringify({ members });
        if (event === "drawing-group") editor.commands["drawing-group"](payload);
        else if (event === "drawing-distribute")
          editor.commands["drawing-distribute"](value, payload);
        else
          (editor.commands["align-objects"] as (v?: string, p?: string) => boolean)(value, payload);
        return true;
      }
      if (event === "align-objects" && editor) {
        editor.commands["align-objects"](value);
        return true;
      }
      return true;
    }
    const editor = this.host.editor();
    if (!editor) return false;
    // Picture needs a file picker — open it, then insert the chosen image.
    if (event === "insert-picture") {
      this.host.pickImage();
      return true;
    }
    // Online Pictures — the address dialog (Word's Insert Pictures window;
    // no picture service here, so the dialog takes an image URL). The dialog
    // commit arrives as `online-picture:ok` on the element.
    if (event === "online-picture") {
      this.host.openOnlinePictures();
      return true;
    }
    // Change Picture — a picker over the selected image; the picked source
    // replaces it at the same frame size (Picture Format > Adjust).
    if (event === "change-picture") {
      this.host.pickPicture();
      return true;
    }
    // Reset Picture and Size — the natural size is a decode only the
    // browser-side painter can read (the paint's pin table), so resolve the
    // selected picture's decoded dimensions here and pass them in. An
    // unpinned or undecoded source degrades to the plain Reset Picture.
    if (event === "reset-picture-size") {
      const target = this.host.activeEditor() ?? editor;
      const sel = target.state.selection;
      const src =
        sel instanceof NodeSelection && sel.node.type.name === "image"
          ? (sel.node.attrs as { src?: unknown }).src
          : undefined;
      const image = typeof src === "string" && src ? pinImage(src) : undefined;
      const natural =
        image?.ready && image.width > 0 && image.height > 0
          ? { width: image.width, height: image.height }
          : undefined;
      this.host.focusBridge();
      target.commands["reset-picture-size"](natural);
      return true;
    }
    // Size and Position — open the drawing dialog prefilled from the selected
    // floating drawing; the commit arrives via drawing-properties:ok
    // (drawing-properties-apply).
    if (event === "drawing-properties") {
      const dialog = this.host
        .element()
        .shadowRoot?.querySelector("docen-drawing-properties-dialog") as {
        show(state: DrawingPropertiesState): void;
      } | null;
      const state = this.host.drawingState();
      if (dialog && state) dialog.show(state);
      return true;
    }
    // Crop — the bridge enters crop mode on the selected image (the overlay
    // previews the full source; Enter / a press outside commits, Esc cancels).
    if (event === "drawing-crop") {
      this.host.enterCropMode();
      return true;
    }
    // Edit Shape / Edit Points — enter edit points mode on the selected shape
    if (event === "edit-shape" || event === "shape-edit-points") {
      this.host.enterEditPointsMode?.();
      return true;
    }
    // Text Box — insert a centered floating wps text box. Shapes — arm the
    // drawer with the picked preset (Word's drag-to-draw: the canvas commits
    // the draw through applyShapeDraw, which disarms — one pick, one shape).
    if (event === "text-box") {
      this.host.insertShapeAt(undefined);
      return true;
    }
    if (event === "shapes") {
      this.host.armShapeDrawer(value ?? "rect");
      return true;
    }
    // Insert → Pages menu: WordArt inserts a preset-styled text box.
    if (event === "wordart") {
      this.host.insertWordArt();
      return true;
    }
    return false;
  }
}
