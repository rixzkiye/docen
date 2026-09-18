import type { Editor } from "@docen/docx/core";

export interface SelectionRange {
  from: number;
  to: number;
}

/**
 * Manages non-contiguous (multi-range) text selections.
 * Allows multiple disjoint text selections across the document,
 * supporting unified formatting, deletion, and copying.
 */
export class MultiSelectionManager {
  private _ranges: SelectionRange[] = [];

  get ranges(): readonly SelectionRange[] {
    return this._ranges;
  }

  hasRanges(): boolean {
    return this._ranges.length > 0;
  }

  clear(): void {
    this._ranges = [];
  }

  setRanges(ranges: SelectionRange[]): void {
    this._ranges = ranges
      .filter((r) => r.from !== r.to)
      .map((r) => ({
        from: Math.min(r.from, r.to),
        to: Math.max(r.from, r.to),
      }))
      .sort((a, b) => a.from - b.from);
  }

  addRange(range: SelectionRange): void {
    if (range.from === range.to) return;
    const norm = {
      from: Math.min(range.from, range.to),
      to: Math.max(range.from, range.to),
    };
    // Merge overlapping or adjacent ranges
    const merged: SelectionRange[] = [];
    let placed = false;
    for (const r of this._ranges) {
      if (placed) {
        merged.push(r);
      } else if (norm.to < r.from) {
        merged.push(norm);
        merged.push(r);
        placed = true;
      } else if (norm.from > r.to) {
        merged.push(r);
      } else {
        norm.from = Math.min(norm.from, r.from);
        norm.to = Math.max(norm.to, r.to);
      }
    }
    if (!placed) merged.push(norm);
    this._ranges = merged;
  }

  allRanges(primary?: SelectionRange): SelectionRange[] {
    if (!primary || primary.from === primary.to) return [...this._ranges];
    const norm = {
      from: Math.min(primary.from, primary.to),
      to: Math.max(primary.from, primary.to),
    };
    const list = [...this._ranges];
    if (!list.some((r) => r.from === norm.from && r.to === norm.to)) {
      list.push(norm);
    }
    return list.sort((a, b) => a.from - b.from);
  }

  /**
   * Applies an inline formatting mark to all selected ranges.
   */
  applyMark(editor: Editor, markType: string, attrs?: Record<string, unknown>): boolean {
    const primary = {
      from: editor.state.selection.from,
      to: editor.state.selection.to,
    };
    const ranges = this.allRanges(primary);
    if (!ranges.length) return false;

    const mark = editor.schema.marks[markType];
    if (!mark) return false;

    const tr = editor.state.tr;
    for (const r of ranges) {
      tr.addMark(r.from, r.to, mark.create(attrs));
    }
    editor.view.dispatch(tr);
    return true;
  }

  /**
   * Deletes all selected ranges in reverse document order.
   */
  deleteContents(editor: Editor): boolean {
    const primary = {
      from: editor.state.selection.from,
      to: editor.state.selection.to,
    };
    const ranges = this.allRanges(primary);
    if (!ranges.length) return false;

    const tr = editor.state.tr;
    for (let i = ranges.length - 1; i >= 0; i--) {
      const r = ranges[i]!;
      tr.delete(r.from, r.to);
    }
    this.clear();
    editor.view.dispatch(tr);
    return true;
  }

  /**
   * Copies text of all selected ranges joined by newlines.
   */
  copyText(editor: Editor): string {
    const primary = {
      from: editor.state.selection.from,
      to: editor.state.selection.to,
    };
    const ranges = this.allRanges(primary);
    const parts: string[] = [];
    for (const r of ranges) {
      const text = editor.state.doc.textBetween(r.from, r.to, "\n");
      if (text) parts.push(text);
    }
    return parts.join("\n");
  }
}
