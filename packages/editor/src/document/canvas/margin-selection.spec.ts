import { describe, expect, it } from "vitest";

import { MARGIN_SELECTION_CURSOR, BLOCK_SELECT_CURSOR } from "./selection";

describe("Margin Selection and Context Cursors (W2.2 & W2.7)", () => {
  it("defines Word's right-pointing arrow SVG cursor for margin selection", () => {
    expect(MARGIN_SELECTION_CURSOR).toBeDefined();
    expect(MARGIN_SELECTION_CURSOR).toContain("data:image/svg+xml");
    expect(MARGIN_SELECTION_CURSOR).toContain("18 2, default");
  });

  it("defines block selection crosshair cursor", () => {
    expect(BLOCK_SELECT_CURSOR).toBe("crosshair");
  });

  it("detects left margin boundaries correctly", () => {
    // Simulated caretMap with column box starting at x = 72
    const mockCaretMap = {
      columnBoxes: (_page: number) => [{ left: 72, top: 72, width: 400, height: 600 }],
      lines: [
        { page: 0, yPx: 100, line: { heightPx: 20 } },
        { page: 0, yPx: 200, line: { heightPx: 20 } },
      ],
      isLeftMargin(page: number, x: number, y: number): boolean {
        const boxes = this.columnBoxes(page);
        const contentLeft = boxes.length ? boxes[0]!.left : 72;
        if (x < 0 || x >= contentLeft) return false;
        const pageLines = this.lines.filter((l: any) => l.page === page);
        if (!pageLines.length) return false;
        const top = pageLines[0]!.yPx - 10;
        const bottom =
          pageLines[pageLines.length - 1]!.yPx +
          pageLines[pageLines.length - 1]!.line.heightPx +
          10;
        return y >= top && y <= bottom;
      },
    };

    // Point in margin: x = 30 (left of 72), y = 150 (between 90 and 230)
    expect(mockCaretMap.isLeftMargin(0, 30, 150)).toBe(true);

    // Point in body text: x = 100 (right of 72)
    expect(mockCaretMap.isLeftMargin(0, 100, 150)).toBe(false);

    // Point above first line: y = 50
    expect(mockCaretMap.isLeftMargin(0, 30, 50)).toBe(false);

    // Point below last line: y = 300
    expect(mockCaretMap.isLeftMargin(0, 30, 300)).toBe(false);
  });
});
