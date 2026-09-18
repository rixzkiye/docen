import type { Editor } from "@docen/docx/core";

import type { SelectionRange } from "./multi-range";

export interface BlockBounds {
  page: number;
  minX: number;
  maxX: number;
  minY: number;
  maxY: number;
}

/**
 * Computes rectangular text ranges for Alt+drag block selection.
 */
export function computeBlockRanges(
  caretMap: any,
  page: number,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
): SelectionRange[] {
  if (!caretMap || !caretMap.lines) return [];
  const minX = Math.min(x1, x2);
  const maxX = Math.max(x1, x2);
  const minY = Math.min(y1, y2);
  const maxY = Math.max(y1, y2);

  const ranges: SelectionRange[] = [];

  for (const entry of caretMap.lines) {
    if (entry.page !== page) continue;
    const lineTop = entry.yPx;
    const lineBottom = entry.yPx + entry.line.heightPx;
    if (lineBottom < minY || lineTop > maxY) continue;

    const lineMidY = lineTop + entry.line.heightPx / 2;

    if (typeof caretMap.posAtPoint === "function") {
      const p1 = caretMap.posAtPoint(page, minX, lineMidY);
      const p2 = caretMap.posAtPoint(page, maxX, lineMidY);
      if (p1 != null && p2 != null && p1 !== p2) {
        ranges.push({
          from: Math.min(p1, p2),
          to: Math.max(p1, p2),
        });
      }
      continue;
    }

    const items = entry.line.items;
    if (!items || items.length === 0) continue;

    let charOffset = 0;
    let rangeStart = -1;
    let rangeEnd = -1;

    for (const item of items) {
      if (item.kind === "text") {
        const itemX = entry.xPx + item.xPx;
        const charWidth = item.widthPx / Math.max(1, item.text.length);
        for (let i = 0; i < item.text.length; i++) {
          const cx = itemX + i * charWidth;
          if (cx + charWidth >= minX && cx <= maxX) {
            if (rangeStart === -1) rangeStart = charOffset;
            rangeEnd = charOffset + 1;
          }
          charOffset++;
        }
      } else {
        const itemX = entry.xPx + item.xPx;
        if (itemX + item.widthPx >= minX && itemX <= maxX) {
          if (rangeStart === -1) rangeStart = charOffset;
          rangeEnd = charOffset + 1;
        }
        charOffset++;
      }
    }

    if (rangeStart !== -1 && rangeEnd > rangeStart) {
      const lineBase = entry.owner?.innerPos ?? entry.startChar;
      ranges.push({
        from: lineBase + rangeStart,
        to: lineBase + rangeEnd,
      });
    }
  }

  return ranges;
}

/**
 * Deletes all ranges in block selection from end to start.
 */
export function deleteBlockRanges(editor: Editor, ranges: SelectionRange[]): boolean {
  if (!ranges.length) return false;
  const sorted = [...ranges].sort((a, b) => b.from - a.from);
  const tr = editor.state.tr;
  for (const r of sorted) {
    tr.delete(r.from, r.to);
  }
  editor.view.dispatch(tr);
  return true;
}

/**
 * Copies rectangular block text lines.
 */
export function copyBlockRanges(editor: Editor, ranges: SelectionRange[]): string {
  if (!ranges.length) return "";
  const sorted = [...ranges].sort((a, b) => a.from - b.from);
  const lines: string[] = [];
  for (const r of sorted) {
    lines.push(editor.state.doc.textBetween(r.from, r.to));
  }
  return lines.join("\n");
}
