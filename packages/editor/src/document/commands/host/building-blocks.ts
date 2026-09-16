import type { BuildingBlocksSeed } from "../../../ui/components/workspace/building-blocks-dialog";
import type {
  QuickPartSeed,
  QuickPartValues,
} from "../../../ui/components/workspace/quick-part-dialog";
import type { BuildingBlock, BuildingBlockSlice } from "../../building-blocks";
import { createBuildingBlock, isDuplicateBlockName } from "../../building-blocks";
import type { HostCommandDomain } from "./registry";

/** The captured selection for "Save Selection to Quick Part Gallery…". */
export interface QuickPartCapture {
  slice: BuildingBlockSlice;
  preview: string;
  suggestedName: string;
}

/** The building-blocks domain's view of the host — only what its bodies call. */
export interface BuildingBlocksHostView {
  /** The host element — the shadow-DOM root for every dialog lookup. */
  element(): HTMLElement;
  /** True while the active editor accepts edits (Viewing mode = false). */
  editable(): boolean;
  /** The document's building blocks (docen list, else the Word glossary). */
  blocks(): BuildingBlock[];
  /** Replace the document's block list in ONE transaction (one undo step). */
  setBlocks(blocks: readonly BuildingBlock[]): void;
  /** Insert one block at the caret via the editor command. */
  insertBlock(id: string): void;
  /** The selection as a slice + preview text, or null when collapsed. */
  selectionSlice(): QuickPartCapture | null;
  /** Hand the keyboard back to the document after a dialog action. */
  focusBridge(): void;
}

/**
 * Quick Parts / AutoText host commands: the ribbon's Save Selection dialog and
 * Building Blocks Organizer openers, and the gallery's block insertion. The
 * dialog commits (save/rename/delete) come back through {@link commitSave},
 * {@link renameBlock} and {@link deleteBlock}, called by the host's listeners.
 */
export class BuildingBlocksHostCommands implements HostCommandDomain {
  /** The selection captured when the save dialog opened — the dialog returns
   *  only the metadata; content rides this snapshot. */
  #pending?: BuildingBlockSlice;

  constructor(private readonly host: BuildingBlocksHostView) {}

  // The organizer is a chrome (view+manage) surface — it opens without an
  // editor; save/insert mutate the document and need one.
  readonly chrome: readonly string[] = ["building-blocks-organizer"];

  readonly editor: readonly string[] = ["save-quick-part", "quick-parts"];

  run(event: string, value?: string): boolean {
    if (event === "save-quick-part") {
      this.#openSaveDialog();
      return true;
    }
    if (event === "building-blocks-organizer") {
      const dialog = this.#dialog("docen-building-blocks-dialog") as {
        show(seed: BuildingBlocksSeed): void;
      } | null;
      dialog?.show({ blocks: this.host.blocks(), editable: this.host.editable() });
      return true;
    }
    if (event === "quick-parts") {
      if (typeof value === "string" && value) this.insert(value);
      return true;
    }
    return false;
  }

  /** Insert one block at the caret (the gallery item and the organizer's
   *  Insert button both land here). */
  insert(id: string): void {
    this.host.insertBlock(id);
  }

  /** "Create New Building Block" 确定 — append the captured selection. */
  commitSave(values: QuickPartValues): void {
    const slice = this.#pending;
    this.#pending = undefined;
    if (!slice) return;
    const current = this.host.blocks();
    if (isDuplicateBlockName(current, values.name)) return;
    const block = createBuildingBlock({ ...values, content: slice });
    this.host.setBlocks([...current, block]);
    this.host.focusBridge();
  }

  /** Organizer rename — a duplicate name is rejected (the dialog already
   *  reports it; this is the defensive re-check). */
  renameBlock(id: string, name: string): void {
    const current = this.host.blocks();
    const trimmed = name.trim();
    if (!trimmed || isDuplicateBlockName(current, trimmed, id)) return;
    this.host.setBlocks(
      current.map((block) => (block.id === id ? { ...block, name: trimmed } : block)),
    );
    this.host.focusBridge();
  }

  /** Organizer delete — one transaction; undo restores the block. */
  deleteBlock(id: string): void {
    const current = this.host.blocks();
    if (!current.some((block) => block.id === id)) return;
    this.host.setBlocks(current.filter((block) => block.id !== id));
    this.host.focusBridge();
  }

  #openSaveDialog(): void {
    if (!this.host.editable()) return;
    const capture = this.host.selectionSlice();
    // The ribbon item is disabled without a selection — defensive only.
    if (!capture) return;
    this.#pending = capture.slice;
    const dialog = this.#dialog("docen-quick-part-dialog") as {
      show(seed: QuickPartSeed): void;
    } | null;
    dialog?.show({
      preview: capture.preview,
      suggestedName: capture.suggestedName,
      existingNames: this.host.blocks().map((block) => block.name),
    });
  }

  #dialog(tag: string): unknown {
    return this.host.element().shadowRoot?.querySelector(tag) ?? null;
  }
}
