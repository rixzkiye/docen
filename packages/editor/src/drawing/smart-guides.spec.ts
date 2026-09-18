// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";

import type { Box } from "./geometry";
import { computeSmartGuides, SmartGuidesOverlay, type PageMargins } from "./smart-guides";

const margins: PageMargins = {
  pageWidth: 800,
  pageHeight: 1000,
  contentLeft: 100,
  contentTop: 100,
  contentWidth: 600,
  contentHeight: 800,
};

describe("computeSmartGuides", () => {
  it("snaps to left margin when within 5px threshold", () => {
    // Left margin is at x = 100. Candidate box at x = 103 (diff = 3px)
    const candidate: Box = { x: 103, y: 300, width: 150, height: 100 };
    const result = computeSmartGuides(candidate, margins, []);

    expect(result.snappedX).toBe(true);
    expect(result.snappedBox.x).toBe(100);
    expect(result.guides.some((g) => g.orientation === "vertical" && g.position === 100)).toBe(
      true,
    );
  });

  it("does not snap when distance exceeds threshold", () => {
    // Candidate box at x = 110 (diff = 10px > 5px)
    const candidate: Box = { x: 110, y: 300, width: 150, height: 100 };
    const result = computeSmartGuides(candidate, margins, [], { threshold: 5 });

    expect(result.snappedX).toBe(false);
    expect(result.snappedBox.x).toBe(110);
  });

  it("snaps to center of margins and page center", () => {
    // Margin center = 100 + 300 = 400.
    // Candidate with width 100 has center = x + 50. If x = 348, center = 398 (diff = 2px).
    const candidate: Box = { x: 348, y: 300, width: 100, height: 100 };
    const result = computeSmartGuides(candidate, margins, []);

    expect(result.snappedX).toBe(true);
    expect(result.snappedBox.x).toBe(350); // center at 400
    expect(result.guides.some((g) => g.orientation === "vertical" && g.position === 400)).toBe(
      true,
    );
  });

  it("snaps to right margin", () => {
    // Right margin = 700.
    // Candidate with width 100 has right = x + 100. If x = 598, right = 698 (diff = 2px).
    const candidate: Box = { x: 598, y: 300, width: 100, height: 100 };
    const result = computeSmartGuides(candidate, margins, []);

    expect(result.snappedX).toBe(true);
    expect(result.snappedBox.x).toBe(600); // right edge at 700
  });

  it("snaps to sibling object edges and centers", () => {
    const sibling: Box = { x: 250, y: 200, width: 100, height: 100 };
    // Test left-to-left alignment: candidate at x = 252 (diff = 2px from 250)
    const candidate: Box = { x: 252, y: 500, width: 80, height: 80 };
    const result = computeSmartGuides(candidate, margins, [sibling]);

    expect(result.snappedX).toBe(true);
    expect(result.snappedBox.x).toBe(250);
    expect(result.guides.some((g) => g.orientation === "vertical" && g.position === 250)).toBe(
      true,
    );
  });

  it("snaps to equal horizontal spacing between siblings", () => {
    const sib1: Box = { x: 100, y: 200, width: 100, height: 100 };
    const sib2: Box = { x: 250, y: 200, width: 100, height: 100 };
    // Gap between sib1 and sib2 = 250 - (100 + 100) = 50px.
    // Candidate placed after sib2 should snap to x = 250 + 100 + 50 = 400.
    const candidate: Box = { x: 403, y: 200, width: 100, height: 100 }; // diff = 3px
    const result = computeSmartGuides(candidate, margins, [sib1, sib2]);

    expect(result.snappedX).toBe(true);
    expect(result.snappedBox.x).toBe(400);
  });

  it("snaps to equal vertical spacing between siblings", () => {
    const sib1: Box = { x: 200, y: 100, width: 100, height: 80 };
    const sib2: Box = { x: 200, y: 220, width: 100, height: 80 };
    // Gap between sib1 and sib2 = 220 - (100 + 80) = 40px.
    // Candidate placed after sib2 should snap to y = 220 + 80 + 40 = 340.
    const candidate: Box = { x: 200, y: 342, width: 100, height: 80 }; // diff = 2px
    const result = computeSmartGuides(candidate, margins, [sib1, sib2]);

    expect(result.snappedY).toBe(true);
    expect(result.snappedBox.y).toBe(340);
  });

  it("constrains drag to strictly horizontal or vertical when Shift is held", () => {
    const origin: Box = { x: 200, y: 200, width: 100, height: 100 };
    // Dominant horizontal movement (dx = 60, dy = 15)
    const candH: Box = { x: 260, y: 215, width: 100, height: 100 };
    const resH = computeSmartGuides(candH, margins, [], { shiftKey: true, origin });
    expect(resH.snappedBox.y).toBe(200); // dy locked to 0

    // Dominant vertical movement (dx = 10, dy = 70)
    const candV: Box = { x: 210, y: 270, width: 100, height: 100 };
    const resV = computeSmartGuides(candV, margins, [], { shiftKey: true, origin });
    expect(resV.snappedBox.x).toBe(200); // dx locked to 0
  });
});

describe("SmartGuidesOverlay", () => {
  it("renders SVG guide lines and clears them", () => {
    const overlay = new SmartGuidesOverlay();
    expect(overlay.el.style.display).toBe("none");

    overlay.render([
      {
        orientation: "vertical",
        position: 100,
        start: 50,
        end: 300,
        targetType: "margin",
      },
      {
        orientation: "horizontal",
        position: 200,
        start: 100,
        end: 400,
        targetType: "object",
      },
    ]);

    expect(overlay.el.style.display).toBe("block");
    expect(overlay.el.children.length).toBe(2);

    const l1 = overlay.el.children[0] as SVGLineElement;
    expect(l1.getAttribute("x1")).toBe("100");
    expect(l1.getAttribute("y1")).toBe("50");
    expect(l1.getAttribute("stroke-dasharray")).toBe("4,3");

    overlay.clear();
    expect(overlay.el.style.display).toBe("none");
    expect(overlay.el.children.length).toBe(0);
  });
});
