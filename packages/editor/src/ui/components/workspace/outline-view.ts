import type { StylesOptions } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { FASTElement, css, customElement, html, observable } from "@microsoft/fast-element";

import {
  demoteHeading,
  detectNodeHeadingLevel,
  findHeadingSectionRange,
  moveHeadingSection,
  promoteHeading,
  setHeadingLevel,
} from "../../../document/commands/outline";

export interface OutlineViewBlock {
  id: string;
  index: number;
  pos: number;
  level: number | null; // 1-9 for headings, null for body
  text: string;
  hasChildren?: boolean;
}

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    width: 100%;
    height: 100%;
    min-height: 0;
    box-sizing: border-box;
    font-family: var(
      --docen-font-family,
      "Segoe UI",
      -apple-system,
      BlinkMacSystemFont,
      Roboto,
      sans-serif
    );
    background: #ffffff;
    color: #242424;
    user-select: none;
  }

  .outline-toolbar {
    display: flex;
    flex-direction: row;
    align-items: center;
    gap: 6px;
    padding: 6px 10px;
    background: #f8f9fa;
    border-bottom: 1px solid var(--docen-color-divider, #d1d1d1);
    flex-wrap: wrap;
    flex: 0 0 auto;
    font-size: 12px;
  }

  .toolbar-group {
    display: flex;
    align-items: center;
    gap: 2px;
    padding-right: 6px;
    border-right: 1px solid var(--docen-color-divider, #e0e0e0);
  }

  .toolbar-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    height: 24px;
    min-width: 24px;
    padding: 0 6px;
    background: transparent;
    border: 1px solid transparent;
    border-radius: 3px;
    cursor: pointer;
    font-size: 13px;
    color: #333333;
  }

  .toolbar-btn:hover {
    background: rgba(0, 0, 0, 0.06);
    border-color: rgba(0, 0, 0, 0.12);
  }

  .toolbar-btn:active {
    background: rgba(0, 0, 0, 0.12);
  }

  .toolbar-btn.close-btn {
    color: #d13438;
    font-weight: 600;
  }

  .toolbar-select,
  .toolbar-toggle {
    height: 24px;
    font-size: 12px;
    padding: 2px 6px;
    border: 1px solid var(--docen-color-divider, #d1d1d1);
    border-radius: 3px;
    background: #ffffff;
    color: #333333;
  }

  .toggle-label {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    cursor: pointer;
    font-size: 12px;
  }

  .outline-content {
    flex: 1;
    min-height: 0;
    overflow: auto;
    padding: 16px 24px;
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  .outline-item-wrapper {
    position: relative;
    display: flex;
    flex-direction: column;
  }

  .outline-item {
    display: flex;
    flex-direction: row;
    align-items: center;
    padding: 3px 8px;
    border-radius: 3px;
    cursor: default;
    border: 1px solid transparent;
    transition: background 0.1s ease;
  }

  .outline-item:hover {
    background: rgba(0, 0, 0, 0.04);
  }

  .outline-item.selected {
    background: #cce8ff;
    border-color: #99d1ff;
  }

  .outline-item.dragging {
    opacity: 0.5;
  }

  .toggle-fold-btn {
    width: 18px;
    height: 18px;
    flex: 0 0 18px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    margin-right: 6px;
    border-radius: 50%;
    border: 1px solid #a0a0a0;
    background: #ffffff;
    font-size: 12px;
    font-weight: bold;
    cursor: pointer;
    user-select: none;
    line-height: 1;
  }

  .toggle-fold-btn.collapsed {
    background: #0f6cbd;
    border-color: #0f6cbd;
    color: #ffffff;
  }

  .toggle-fold-btn.leaf {
    border: none;
    background: transparent;
    color: #707070;
  }

  .item-text {
    flex: 1;
    min-width: 0;
    white-space: pre-wrap;
    word-break: break-word;
    user-select: text;
  }

  :host([show-first-line-only]) .item-text,
  .first-line-only .item-text {
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
    max-width: 100%;
  }

  .level-1 {
    font-weight: 700;
    font-size: 16px;
    margin-left: 0;
  }
  .level-2 {
    font-weight: 600;
    font-size: 15px;
    margin-left: 18px;
  }
  .level-3 {
    font-weight: 600;
    font-size: 14px;
    margin-left: 36px;
  }
  .level-4 {
    font-weight: 600;
    font-size: 13px;
    margin-left: 54px;
  }
  .level-5 {
    font-weight: 600;
    font-size: 13px;
    margin-left: 72px;
  }
  .level-6 {
    font-weight: 600;
    font-size: 12px;
    margin-left: 90px;
  }
  .level-7 {
    font-weight: 600;
    font-size: 12px;
    margin-left: 108px;
  }
  .level-8 {
    font-weight: 600;
    font-size: 12px;
    margin-left: 126px;
  }
  .level-9 {
    font-weight: 600;
    font-size: 12px;
    margin-left: 144px;
  }
  .level-body {
    font-weight: 400;
    font-size: 12px;
    color: #404040;
    margin-left: 18px;
  }

  .drop-indicator-before {
    position: absolute;
    top: -2px;
    left: 0;
    right: 0;
    height: 3px;
    background: #0f6cbd;
    border-radius: 1.5px;
    z-index: 10;
    pointer-events: none;
  }

  .drop-indicator-after {
    position: absolute;
    bottom: -2px;
    left: 0;
    right: 0;
    height: 3px;
    background: #0f6cbd;
    border-radius: 1.5px;
    z-index: 10;
    pointer-events: none;
  }

  .empty-state {
    padding: 32px;
    text-align: center;
    color: #707070;
    font-size: 13px;
  }
`;

const template = html<DocenOutlineView>`
  <div class="outline-toolbar" part="toolbar">
    <div class="toolbar-group">
      <select
        class="toolbar-select"
        part="level-select"
        title="Outline Level"
        @change="${(x, c) => x.onLevelSelectChange(c.event)}"
      >
        <option value="1">Level 1</option>
        <option value="2">Level 2</option>
        <option value="3">Level 3</option>
        <option value="4">Level 4</option>
        <option value="5">Level 5</option>
        <option value="6">Level 6</option>
        <option value="7">Level 7</option>
        <option value="8">Level 8</option>
        <option value="9">Level 9</option>
        <option value="0">Body Text</option>
      </select>
      <button
        class="toolbar-btn"
        part="btn-promote"
        title="Promote (Alt+Shift+Left)"
        @click="${(x) => x.promoteSelected()}"
      >
        ←
      </button>
      <button
        class="toolbar-btn"
        part="btn-demote"
        title="Demote (Alt+Shift+Right)"
        @click="${(x) => x.demoteSelected()}"
      >
        →
      </button>
      <button
        class="toolbar-btn"
        part="btn-promote-h1"
        title="Promote to Heading 1"
        @click="${(x) => x.promoteToH1()}"
      >
        ⇤
      </button>
      <button
        class="toolbar-btn"
        part="btn-demote-body"
        title="Demote to Body Text"
        @click="${(x) => x.demoteToBody()}"
      >
        ⇥
      </button>
    </div>

    <div class="toolbar-group">
      <button
        class="toolbar-btn"
        part="btn-move-up"
        title="Move Up (Alt+Shift+Up)"
        @click="${(x) => x.moveSelectedUp()}"
      >
        ↑
      </button>
      <button
        class="toolbar-btn"
        part="btn-move-down"
        title="Move Down (Alt+Shift+Down)"
        @click="${(x) => x.moveSelectedDown()}"
      >
        ↓
      </button>
      <button
        class="toolbar-btn"
        part="btn-expand"
        title="Expand"
        @click="${(x) => x.expandSelected()}"
      >
        +
      </button>
      <button
        class="toolbar-btn"
        part="btn-collapse"
        title="Collapse"
        @click="${(x) => x.collapseSelected()}"
      >
        −
      </button>
    </div>

    <div class="toolbar-group">
      <label class="toggle-label" title="Show Level Filter">
        Show Level:
        <select
          class="toolbar-select"
          part="show-level-select"
          @change="${(x, c) => x.onShowLevelChange(c.event)}"
        >
          <option value="0">All Levels</option>
          <option value="1">Level 1</option>
          <option value="2">Level 2</option>
          <option value="3">Level 3</option>
          <option value="4">Level 4</option>
          <option value="5">Level 5</option>
          <option value="6">Level 6</option>
          <option value="7">Level 7</option>
          <option value="8">Level 8</option>
          <option value="9">Level 9</option>
        </select>
      </label>
    </div>

    <div class="toolbar-group">
      <label class="toggle-label" title="Show Body Text">
        <input
          type="checkbox"
          part="toggle-body-text"
          ?checked="${(x) => x.showBodyText}"
          @change="${(x, c) => x.onToggleBodyText(c.event)}"
        />
        Show Body Text
      </label>
      <label class="toggle-label" title="Show First Line Only">
        <input
          type="checkbox"
          part="toggle-first-line"
          ?checked="${(x) => x.showFirstLineOnly}"
          @change="${(x, c) => x.onToggleFirstLineOnly(c.event)}"
        />
        Show First Line Only
      </label>
    </div>

    <div class="toolbar-group" style="border-right: none; margin-left: auto;">
      <button
        class="toolbar-btn close-btn"
        part="btn-close"
        title="Close Outline View"
        @click="${(x) => x.closeOutlineView()}"
      >
        ✕ Close Outline View
      </button>
    </div>
  </div>

  <div
    class="outline-content ${(x) => (x.showFirstLineOnly ? "first-line-only" : "")}"
    part="content"
  >
    ${(x) => x.renderOutlineItems()}
  </div>
`;

/**
 * `<docen-outline-view>` — Full MS Word-style interactive outline view mode component.
 */
@customElement({ name: "docen-outline-view", template, styles })
export class DocenOutlineView extends FASTElement {
  @observable editor: Editor | null = null;
  @observable blocks: OutlineViewBlock[] = [];
  @observable collapsedIds = new Set<string>();
  @observable selectedIds = new Set<string>();
  @observable showLevel = 0; // 0 = all
  @observable showBodyText = true;
  @observable showFirstLineOnly = false;
  @observable dropTarget: { targetId: string; position: "before" | "after" } | null = null;
  @observable draggedId: string | null = null;

  #onTransaction = (): void => {
    this.syncFromEditor();
  };

  #lastAnchorId: string | null = null;

  override connectedCallback(): void {
    super.connectedCallback();
    this.addEventListener("keydown", this.#onKeyDown);
    if (this.editor) this.syncFromEditor();
  }

  override disconnectedCallback(): void {
    this.removeEventListener("keydown", this.#onKeyDown);
    if (this.editor) {
      this.editor.off("transaction", this.#onTransaction);
    }
    super.disconnectedCallback();
  }

  setEditor(editor: Editor): void {
    if (this.editor) {
      this.editor.off("transaction", this.#onTransaction);
    }
    this.editor = editor;
    editor.on("transaction", this.#onTransaction);
    this.syncFromEditor();
  }

  syncFromEditor(): void {
    if (!this.editor) return;
    const doc = this.editor.state.doc;
    const styles = (doc.attrs as { styles?: StylesOptions }).styles;
    const list: OutlineViewBlock[] = [];

    let currentPos = 0;
    for (let i = 0; i < doc.childCount; i++) {
      const node = doc.child(i);
      const level = detectNodeHeadingLevel(node, styles);
      list.push({
        id: `block-${i}`,
        index: i,
        pos: currentPos,
        level,
        text: node.textContent || "(Empty)",
        hasChildren: false,
      });
      currentPos += node.nodeSize;
    }

    // Determine hasChildren for headings
    for (let i = 0; i < list.length; i++) {
      const item = list[i]!;
      if (item.level != null) {
        if (i + 1 < list.length) {
          const next = list[i + 1]!;
          if (next.level == null || next.level > item.level) {
            item.hasChildren = true;
          }
        }
      }
    }

    this.blocks = list;
  }

  /**
   * Check if a heading is collapsed directly or via an ancestor heading.
   */
  isHeadingCollapsed(heading: OutlineViewBlock): boolean {
    if (heading.level == null) return false;
    if (this.collapsedIds.has(heading.id)) return true;
    for (let i = heading.index - 1; i >= 0; i--) {
      const b = this.blocks[i];
      if (b && b.level != null && b.level < heading.level) {
        if (this.collapsedIds.has(b.id)) return true;
        if (this.isHeadingCollapsed(b)) return true;
      }
    }
    return false;
  }

  /**
   * Determine whether an item should be visible given folding and filters.
   */
  isItemVisible(item: OutlineViewBlock): boolean {
    // 1. Body text filter
    if (item.level == null && !this.showBodyText) {
      return false;
    }

    // 2. Heading level filter
    if (this.showLevel > 0) {
      if (item.level != null && item.level > this.showLevel) {
        return false;
      }
      if (item.level == null) {
        let ancestorLvl: number | null = null;
        for (let i = item.index - 1; i >= 0; i--) {
          const b = this.blocks[i];
          if (b && b.level != null) {
            ancestorLvl = b.level;
            break;
          }
        }
        if (ancestorLvl != null && ancestorLvl > this.showLevel) {
          return false;
        }
      }
    }

    // 3. Folding check: is this item inside any collapsed heading?
    if (item.level != null) {
      // Check if any ancestor heading is collapsed
      for (let i = item.index - 1; i >= 0; i--) {
        const b = this.blocks[i];
        if (b && b.level != null) {
          if (b.level < item.level && this.collapsedIds.has(b.id)) {
            return false;
          }
          if (b.level < item.level && this.isHeadingCollapsed(b)) {
            return false;
          }
        }
      }
    } else {
      // Body paragraph: find direct owning heading
      for (let i = item.index - 1; i >= 0; i--) {
        const b = this.blocks[i];
        if (b && b.level != null) {
          if (this.isHeadingCollapsed(b)) return false;
          break;
        }
      }
    }

    return true;
  }

  renderOutlineItems(): ReturnType<typeof html> {
    const visible = this.blocks.filter((b) => this.isItemVisible(b));
    if (visible.length === 0) {
      return html`<div class="empty-state">No outline items to display.</div>`;
    }

    return html`
      ${visible.map(
        (item) => html`
          <div class="outline-item-wrapper" data-id="${item.id}">
            ${
              this.dropTarget?.targetId === item.id && this.dropTarget.position === "before"
                ? html`<div class="drop-indicator-before"></div>`
                : ""
            }
            <div
              class="outline-item ${this.selectedIds.has(item.id) ? "selected" : ""} ${this.draggedId === item.id ? "dragging" : ""}"
              draggable="true"
              @click="${(x: DocenOutlineView, c) => x.onItemClick(item.id, c.event as MouseEvent)}"
              @dragstart="${(x: DocenOutlineView, c) => x.onDragStart(c.event as DragEvent, item.id)}"
              @dragover="${(x: DocenOutlineView, c) => x.onDragOver(c.event as DragEvent, item.id)}"
              @dragleave="${(x: DocenOutlineView, c) => x.onDragLeave(c.event as DragEvent, item.id)}"
              @drop="${(x: DocenOutlineView, c) => x.onDrop(c.event as DragEvent, item.id)}"
            >
              <button
                class="toggle-fold-btn ${item.level != null ? (item.hasChildren ? (this.collapsedIds.has(item.id) ? "collapsed" : "expanded") : "leaf") : "leaf"}"
                @click="${(x: DocenOutlineView, c) => {
                  c.event.stopPropagation();
                  x.toggleCollapse(item.id);
                }}"
              >
                ${
                  item.level != null
                    ? item.hasChildren
                      ? this.collapsedIds.has(item.id)
                        ? "+"
                        : "−"
                      : "○"
                    : "•"
                }
              </button>
              <div class="item-text ${item.level != null ? `level-${item.level}` : "level-body"}">
                ${item.text}
              </div>
            </div>
            ${
              this.dropTarget?.targetId === item.id && this.dropTarget.position === "after"
                ? html`<div class="drop-indicator-after"></div>`
                : ""
            }
          </div>
        `,
      )}
    `;
  }

  toggleCollapse(id: string): void {
    const next = new Set(this.collapsedIds);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    this.collapsedIds = next;
  }

  collapseAll(): void {
    const next = new Set<string>();
    for (const b of this.blocks) {
      if (b.level != null && b.hasChildren) next.add(b.id);
    }
    this.collapsedIds = next;
  }

  expandAll(): void {
    this.collapsedIds = new Set();
  }

  onItemClick(id: string, event: MouseEvent): void {
    const nextSelected = new Set(this.selectedIds);
    if (event.ctrlKey || event.metaKey) {
      if (nextSelected.has(id)) nextSelected.delete(id);
      else nextSelected.add(id);
      this.#lastAnchorId = id;
    } else if (event.shiftKey && this.#lastAnchorId) {
      const idxA = this.blocks.findIndex((b) => b.id === this.#lastAnchorId);
      const idxB = this.blocks.findIndex((b) => b.id === id);
      if (idxA !== -1 && idxB !== -1) {
        nextSelected.clear();
        const start = Math.min(idxA, idxB);
        const end = Math.max(idxA, idxB);
        for (let i = start; i <= end; i++) {
          nextSelected.add(this.blocks[i]!.id);
        }
      }
    } else {
      nextSelected.clear();
      nextSelected.add(id);
      this.#lastAnchorId = id;
    }
    this.selectedIds = nextSelected;

    // Place caret at that item in editor
    const item = this.blocks.find((b) => b.id === id);
    if (item && this.editor) {
      this.editor.commands.setTextSelection(item.pos + 1);
    }
  }

  onDragStart(event: DragEvent, id: string): void {
    this.draggedId = id;
    event.dataTransfer?.setData("text/plain", id);
    if (!this.selectedIds.has(id)) {
      this.selectedIds = new Set([id]);
      this.#lastAnchorId = id;
    }
  }

  onDragOver(event: DragEvent, id: string): void {
    event.preventDefault();
    const currentTarget = event.currentTarget as HTMLElement | null;
    if (!currentTarget) return;
    const rect = currentTarget.getBoundingClientRect();
    const relY = event.clientY - rect.top;
    const position: "before" | "after" = relY < rect.height / 2 ? "before" : "after";
    this.dropTarget = { targetId: id, position };
  }

  onDragLeave(_event: DragEvent, id: string): void {
    if (this.dropTarget?.targetId === id) {
      this.dropTarget = null;
    }
  }

  onDrop(event: DragEvent, targetId: string): void {
    event.preventDefault();
    const draggedId = this.draggedId;
    const dropTarget = this.dropTarget;
    this.draggedId = null;
    this.dropTarget = null;
    if (!draggedId || !dropTarget || !this.editor) return;

    const srcBlock = this.blocks.find((b) => b.id === draggedId);
    const tgtBlock = this.blocks.find((b) => b.id === targetId);
    if (!srcBlock || !tgtBlock || srcBlock.id === tgtBlock.id) return;

    if (srcBlock.level != null) {
      moveHeadingSection(this.editor, srcBlock.index, tgtBlock.index, dropTarget.position);
    } else {
      // Reorder single block
      const doc = this.editor.state.doc;
      const tr = this.editor.state.tr;
      const blocks: any[] = [];
      for (let i = 0; i < doc.childCount; i++) blocks.push(doc.child(i));
      const [moved] = blocks.splice(srcBlock.index, 1);
      let insertIdx = dropTarget.position === "before" ? tgtBlock.index : tgtBlock.index + 1;
      if (srcBlock.index < insertIdx) insertIdx--;
      blocks.splice(insertIdx, 0, moved);
      const newDoc = tr.doc.type.create(tr.doc.attrs, blocks);
      tr.replaceWith(0, tr.doc.content.size, newDoc.content);
      this.editor.view.dispatch(tr);
    }
    this.syncFromEditor();
  }

  promoteSelected(): boolean {
    if (!this.editor) return false;
    const selectedIndices = Array.from(this.selectedIds)
      .map((id) => this.blocks.find((b) => b.id === id)?.index)
      .filter((idx): idx is number => idx != null)
      .sort((a, b) => a - b);

    if (selectedIndices.length === 0) {
      return promoteHeading(this.editor);
    }

    let changed = false;
    for (const idx of selectedIndices) {
      if (promoteHeading(this.editor, idx)) changed = true;
    }
    this.syncFromEditor();
    return changed;
  }

  demoteSelected(): boolean {
    if (!this.editor) return false;
    const selectedIndices = Array.from(this.selectedIds)
      .map((id) => this.blocks.find((b) => b.id === id)?.index)
      .filter((idx): idx is number => idx != null)
      .sort((a, b) => b - a);

    if (selectedIndices.length === 0) {
      return demoteHeading(this.editor);
    }

    let changed = false;
    for (const idx of selectedIndices) {
      if (demoteHeading(this.editor, idx)) changed = true;
    }
    this.syncFromEditor();
    return changed;
  }

  promoteToH1(): boolean {
    if (!this.editor) return false;
    const selectedIndices = Array.from(this.selectedIds)
      .map((id) => this.blocks.find((b) => b.id === id)?.index)
      .filter((idx): idx is number => idx != null);

    if (selectedIndices.length === 0) {
      return setHeadingLevel(this.editor, 1);
    }

    for (const idx of selectedIndices) {
      setHeadingLevel(this.editor, 1, idx);
    }
    this.syncFromEditor();
    return true;
  }

  demoteToBody(): boolean {
    if (!this.editor) return false;
    const selectedIndices = Array.from(this.selectedIds)
      .map((id) => this.blocks.find((b) => b.id === id)?.index)
      .filter((idx): idx is number => idx != null);

    if (selectedIndices.length === 0) {
      return setHeadingLevel(this.editor, 0);
    }

    for (const idx of selectedIndices) {
      setHeadingLevel(this.editor, 0, idx);
    }
    this.syncFromEditor();
    return true;
  }

  moveSelectedUp(): boolean {
    if (!this.editor) return false;
    const selected = Array.from(this.selectedIds)
      .map((id) => this.blocks.find((b) => b.id === id))
      .filter((b): b is OutlineViewBlock => b != null)
      .sort((a, b) => a.index - b.index);

    if (selected.length === 0) return false;
    const first = selected[0]!;
    if (first.index <= 0) return false;

    if (first.level != null) {
      const doc = this.editor.state.doc;
      const styles = (doc.attrs as { styles?: StylesOptions }).styles;
      const srcRange = findHeadingSectionRange(doc, first.index, styles);
      if (srcRange.startIndex <= 0) return false;
      moveHeadingSection(this.editor, first.index, srcRange.startIndex - 1, "before");
    } else {
      const doc = this.editor.state.doc;
      const tr = this.editor.state.tr;
      const blocks: any[] = [];
      for (let i = 0; i < doc.childCount; i++) blocks.push(doc.child(i));
      const [moved] = blocks.splice(first.index, 1);
      blocks.splice(first.index - 1, 0, moved);
      const newDoc = tr.doc.type.create(tr.doc.attrs, blocks);
      tr.replaceWith(0, tr.doc.content.size, newDoc.content);
      this.editor.view.dispatch(tr);
    }
    this.syncFromEditor();
    return true;
  }

  moveSelectedDown(): boolean {
    if (!this.editor) return false;
    const selected = Array.from(this.selectedIds)
      .map((id) => this.blocks.find((b) => b.id === id))
      .filter((b): b is OutlineViewBlock => b != null)
      .sort((a, b) => b.index - a.index);

    if (selected.length === 0) return false;
    const target = selected[0]!;
    if (target.index >= this.blocks.length - 1) return false;

    if (target.level != null) {
      const doc = this.editor.state.doc;
      const styles = (doc.attrs as { styles?: StylesOptions }).styles;
      const srcRange = findHeadingSectionRange(doc, target.index, styles);
      if (srcRange.endIndex >= doc.childCount) return false;
      moveHeadingSection(this.editor, target.index, srcRange.endIndex, "after");
    } else {
      const doc = this.editor.state.doc;
      const tr = this.editor.state.tr;
      const blocks: any[] = [];
      for (let i = 0; i < doc.childCount; i++) blocks.push(doc.child(i));
      const [moved] = blocks.splice(target.index, 1);
      blocks.splice(target.index + 1, 0, moved);
      const newDoc = tr.doc.type.create(tr.doc.attrs, blocks);
      tr.replaceWith(0, tr.doc.content.size, newDoc.content);
      this.editor.view.dispatch(tr);
    }
    this.syncFromEditor();
    return true;
  }

  expandSelected(): void {
    const next = new Set(this.collapsedIds);
    for (const id of this.selectedIds) {
      next.delete(id);
    }
    this.collapsedIds = next;
  }

  collapseSelected(): void {
    const next = new Set(this.collapsedIds);
    for (const id of this.selectedIds) {
      const b = this.blocks.find((x) => x.id === id);
      if (b && b.level != null && b.hasChildren) {
        next.add(id);
      }
    }
    this.collapsedIds = next;
  }

  onLevelSelectChange(event: Event): void {
    const select = event.target as HTMLSelectElement;
    const lvl = parseInt(select.value, 10);
    if (!this.editor) return;
    const selectedIndices = Array.from(this.selectedIds)
      .map((id) => this.blocks.find((b) => b.id === id)?.index)
      .filter((idx): idx is number => idx != null);

    if (selectedIndices.length === 0) {
      setHeadingLevel(this.editor, lvl);
    } else {
      for (const idx of selectedIndices) {
        setHeadingLevel(this.editor, lvl, idx);
      }
    }
    this.syncFromEditor();
  }

  onShowLevelChange(event: Event): void {
    const select = event.target as HTMLSelectElement;
    this.showLevel = parseInt(select.value, 10);
  }

  onToggleBodyText(event: Event): void {
    this.showBodyText = (event.target as HTMLInputElement).checked;
  }

  onToggleFirstLineOnly(event: Event): void {
    this.showFirstLineOnly = (event.target as HTMLInputElement).checked;
    this.toggleAttribute("show-first-line-only", this.showFirstLineOnly);
  }

  closeOutlineView(): void {
    this.dispatchEvent(
      new CustomEvent("outline-view:close", {
        bubbles: true,
        composed: true,
      }),
    );
    const hostEl =
      (this.closest("docen-document") as HTMLElement & {
        setAttribute: (k: string, v: string) => void;
      }) ?? document.querySelector("docen-document");
    if (hostEl) {
      hostEl.setAttribute("view", "print");
    }
  }

  #onKeyDown = (event: KeyboardEvent): void => {
    if (event.altKey && event.shiftKey) {
      if (event.key === "ArrowLeft" || event.key === "Left") {
        event.preventDefault();
        this.promoteSelected();
      } else if (event.key === "ArrowRight" || event.key === "Right") {
        event.preventDefault();
        this.demoteSelected();
      } else if (event.key === "ArrowUp" || event.key === "Up") {
        event.preventDefault();
        this.moveSelectedUp();
      } else if (event.key === "ArrowDown" || event.key === "Down") {
        event.preventDefault();
        this.moveSelectedDown();
      }
    }
  };
}

export default DocenOutlineView;
