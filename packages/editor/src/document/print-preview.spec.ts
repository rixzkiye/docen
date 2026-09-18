// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";

import {
  calculateBookletImposition,
  DocenPrintPreview,
  expandCopies,
  parsePageRange,
  resolvePaperDimensions,
} from "../ui/components/workspace/print-preview";

describe("Print Preview: parsePageRange", () => {
  it("returns all pages for 'all'", () => {
    expect(parsePageRange({ type: "all", totalPages: 5 })).toEqual([1, 2, 3, 4, 5]);
  });

  it("returns current page clamped within range", () => {
    expect(parsePageRange({ type: "current", currentPage: 3, totalPages: 5 })).toEqual([3]);
    expect(parsePageRange({ type: "current", currentPage: 10, totalPages: 5 })).toEqual([5]);
    expect(parsePageRange({ type: "current", currentPage: -2, totalPages: 5 })).toEqual([1]);
  });

  it("returns selection pages if valid, else falls back to current page", () => {
    expect(
      parsePageRange({
        type: "selection",
        selectionPages: [2, 4],
        currentPage: 1,
        totalPages: 5,
      }),
    ).toEqual([2, 4]);

    expect(
      parsePageRange({
        type: "selection",
        selectionPages: [],
        currentPage: 2,
        totalPages: 5,
      }),
    ).toEqual([2]);
  });

  it("parses custom ranges with hyphenated spans and comma separated numbers", () => {
    expect(
      parsePageRange({
        type: "custom",
        custom: "1-3, 5, 7-10",
        totalPages: 10,
      }),
    ).toEqual([1, 2, 3, 5, 7, 8, 9, 10]);

    // Handles whitespace and reversed ranges
    expect(
      parsePageRange({
        type: "custom",
        custom: " 4 - 2 , 6 ",
        totalPages: 6,
      }),
    ).toEqual([2, 3, 4, 6]);

    // Out of bounds pages clamped / ignored
    expect(
      parsePageRange({
        type: "custom",
        custom: "1, 20, 3-100",
        totalPages: 4,
      }),
    ).toEqual([1, 3, 4]);
  });

  it("falls back to all pages on empty or invalid custom range", () => {
    expect(parsePageRange({ type: "custom", custom: "", totalPages: 3 })).toEqual([1, 2, 3]);
    expect(parsePageRange({ type: "custom", custom: "xyz", totalPages: 3 })).toEqual([1, 2, 3]);
  });

  it("returns empty array for zero or negative totalPages", () => {
    expect(parsePageRange({ type: "all", totalPages: 0 })).toEqual([]);
    expect(parsePageRange({ type: "all", totalPages: -1 })).toEqual([]);
  });
});

describe("Print Preview: expandCopies", () => {
  it("leaves single copy untouched", () => {
    expect(expandCopies([1, 2, 3], 1, true)).toEqual([1, 2, 3]);
    expect(expandCopies([1, 2, 3], 1, false)).toEqual([1, 2, 3]);
  });

  it("expands collated copies (1,2,3, 1,2,3...)", () => {
    expect(expandCopies([1, 2, 3], 2, true)).toEqual([1, 2, 3, 1, 2, 3]);
    expect(expandCopies([1, 2], 3, true)).toEqual([1, 2, 1, 2, 1, 2]);
  });

  it("expands uncollated copies (1,1, 2,2, 3,3...)", () => {
    expect(expandCopies([1, 2, 3], 2, false)).toEqual([1, 1, 2, 2, 3, 3]);
    expect(expandCopies([1, 2], 3, false)).toEqual([1, 1, 1, 2, 2, 2]);
  });

  it("clamps copies to at least 1", () => {
    expect(expandCopies([1, 2], 0, true)).toEqual([1, 2]);
    expect(expandCopies([1, 2], -5, false)).toEqual([1, 2]);
  });
});

describe("Print Preview: calculateBookletImposition (2-up)", () => {
  it("imposes a 4-page document into 1 physical sheet (2 sides)", () => {
    const sheets = calculateBookletImposition(4);
    expect(sheets).toHaveLength(2);

    // Sheet 0 Front: Left = 4, Right = 1
    expect(sheets[0]).toEqual({
      sheetIndex: 0,
      side: "front",
      leftPage: 4,
      rightPage: 1,
    });

    // Sheet 0 Back: Left = 2, Right = 3
    expect(sheets[1]).toEqual({
      sheetIndex: 0,
      side: "back",
      leftPage: 2,
      rightPage: 3,
    });
  });

  it("imposes an 8-page document into 2 physical sheets (4 sides)", () => {
    const sheets = calculateBookletImposition(8);
    expect(sheets).toHaveLength(4);

    // Sheet 0 Front: Left = 8, Right = 1
    expect(sheets[0]).toEqual({
      sheetIndex: 0,
      side: "front",
      leftPage: 8,
      rightPage: 1,
    });

    // Sheet 0 Back: Left = 2, Right = 7
    expect(sheets[1]).toEqual({
      sheetIndex: 0,
      side: "back",
      leftPage: 2,
      rightPage: 7,
    });

    // Sheet 1 Front: Left = 6, Right = 3
    expect(sheets[2]).toEqual({
      sheetIndex: 1,
      side: "front",
      leftPage: 6,
      rightPage: 3,
    });

    // Sheet 1 Back: Left = 4, Right = 5
    expect(sheets[3]).toEqual({
      sheetIndex: 1,
      side: "back",
      leftPage: 4,
      rightPage: 5,
    });
  });

  it("pads odd or non-multiple of 4 pages with null (blank slots)", () => {
    const sheets = calculateBookletImposition(5);
    // 5 pages pads to 8 pages -> 2 sheets
    expect(sheets).toHaveLength(4);

    // Pages > 5 are null
    expect(sheets[0]).toEqual({
      sheetIndex: 0,
      side: "front",
      leftPage: null, // 8 > 5
      rightPage: 1,
    });
    expect(sheets[1]).toEqual({
      sheetIndex: 0,
      side: "back",
      leftPage: 2,
      rightPage: null, // 7 > 5
    });
    expect(sheets[2]).toEqual({
      sheetIndex: 1,
      side: "front",
      leftPage: null, // 6 > 5
      rightPage: 3,
    });
    expect(sheets[3]).toEqual({
      sheetIndex: 1,
      side: "back",
      leftPage: 4,
      rightPage: 5,
    });
  });

  it("returns empty array for 0 pages", () => {
    expect(calculateBookletImposition(0)).toEqual([]);
  });
});

describe("Print Preview: resolvePaperDimensions", () => {
  it("preserves base dimensions with auto preset and auto orientation", () => {
    const dims = resolvePaperDimensions("auto", "auto", 800, 1000);
    expect(dims).toEqual({ width: 800, height: 1000 });
  });

  it("applies standard preset sizes", () => {
    const a4 = resolvePaperDimensions("a4", "auto", 500, 500);
    expect(a4.width).toBeCloseTo(793.7, 1);
    expect(a4.height).toBeCloseTo(1122.5, 1);

    const letter = resolvePaperDimensions("letter", "auto", 500, 500);
    expect(letter.width).toBe(816);
    expect(letter.height).toBe(1056);
  });

  it("overrides orientation to portrait or landscape", () => {
    // Landscape base forced to portrait -> width < height
    const portrait = resolvePaperDimensions("auto", "portrait", 1000, 600);
    expect(portrait.width).toBe(600);
    expect(portrait.height).toBe(1000);

    // Portrait base forced to landscape -> width > height
    const landscape = resolvePaperDimensions("auto", "landscape", 600, 1000);
    expect(landscape.width).toBe(1000);
    expect(landscape.height).toBe(600);
  });
});

describe("Print Preview: DocenPrintPreview component", () => {
  it("instantiates and seeds snapshots properly", () => {
    const el = new DocenPrintPreview();
    el.seed({
      snapshots: [
        { width: 800, height: 1000, url: "data:image/png;base64,1" },
        { width: 800, height: 1000, url: "data:image/png;base64,2" },
        { width: 800, height: 1000, url: "data:image/png;base64,3" },
      ],
      filename: "Report.docx",
      currentPage: 2,
    });

    expect(el.getSnapshots()).toHaveLength(3);
    expect(el.getEffectivePageNumbers()).toEqual([1, 2, 3]);
  });
});
