import type { Box } from "./geometry";

export interface PageMargins {
  pageWidth: number;
  pageHeight: number;
  contentLeft: number;
  contentTop: number;
  contentWidth: number;
  contentHeight: number;
}

export interface GuideLine {
  orientation: "vertical" | "horizontal";
  position: number;
  start: number;
  end: number;
  targetType: "margin" | "object" | "spacing";
  label?: string;
}

export interface SnapOptions {
  threshold?: number;
  shiftKey?: boolean;
  origin?: Box;
}

export interface SnapResult {
  snappedBox: Box;
  guides: GuideLine[];
  snappedX: boolean;
  snappedY: boolean;
}

/**
 * Calculates alignment snaps against page margins, sibling objects, and equal spacing.
 * Magnetic snap within 5px threshold (default).
 * Constrains axis to strictly horizontal or vertical if Shift key is held.
 */
export function computeSmartGuides(
  candidateBox: Box,
  margins: PageMargins,
  siblings: readonly Box[],
  options: SnapOptions = {},
): SnapResult {
  const threshold = options.threshold ?? 5;
  let { x, y, width, height } = candidateBox;

  // Shift key constraint: lock to strictly horizontal or vertical drag
  if (options.shiftKey && options.origin) {
    const dx = x - options.origin.x;
    const dy = y - options.origin.y;
    if (Math.abs(dx) >= Math.abs(dy)) {
      y = options.origin.y;
    } else {
      x = options.origin.x;
    }
  }

  const guides: GuideLine[] = [];
  let snappedX = false;
  let snappedY = false;

  // ── X-Axis Snapping ──
  // Dragged points
  const left = x;
  const centerX = x + width / 2;
  const right = x + width;

  interface XTarget {
    pos: number;
    targetType: "margin" | "object";
    spanY: [number, number];
  }

  const xTargets: XTarget[] = [
    // Page margins
    {
      pos: margins.contentLeft,
      targetType: "margin",
      spanY: [margins.contentTop, margins.contentTop + margins.contentHeight],
    },
    {
      pos: margins.contentLeft + margins.contentWidth / 2,
      targetType: "margin",
      spanY: [margins.contentTop, margins.contentTop + margins.contentHeight],
    },
    {
      pos: margins.contentLeft + margins.contentWidth,
      targetType: "margin",
      spanY: [margins.contentTop, margins.contentTop + margins.contentHeight],
    },
    // Page bounds & center
    { pos: 0, targetType: "margin", spanY: [0, margins.pageHeight] },
    { pos: margins.pageWidth / 2, targetType: "margin", spanY: [0, margins.pageHeight] },
    { pos: margins.pageWidth, targetType: "margin", spanY: [0, margins.pageHeight] },
  ];

  // Sibling bounds
  for (const s of siblings) {
    const sSpanY: [number, number] = [s.y, s.y + s.height];
    xTargets.push(
      { pos: s.x, targetType: "object", spanY: sSpanY },
      { pos: s.x + s.width / 2, targetType: "object", spanY: sSpanY },
      { pos: s.x + s.width, targetType: "object", spanY: sSpanY },
    );
  }

  let bestDeltaX = Infinity;
  let bestTargetX: XTarget | null = null;
  let bestSnappedX = x;

  for (const t of xTargets) {
    // 1. Align left to target
    const dLeft = t.pos - left;
    if (Math.abs(dLeft) <= threshold && Math.abs(dLeft) < Math.abs(bestDeltaX)) {
      bestDeltaX = dLeft;
      bestTargetX = t;
      bestSnappedX = t.pos;
    }
    // 2. Align center to target
    const dCenter = t.pos - centerX;
    if (Math.abs(dCenter) <= threshold && Math.abs(dCenter) < Math.abs(bestDeltaX)) {
      bestDeltaX = dCenter;
      bestTargetX = t;
      bestSnappedX = t.pos - width / 2;
    }
    // 3. Align right to target
    const dRight = t.pos - right;
    if (Math.abs(dRight) <= threshold && Math.abs(dRight) < Math.abs(bestDeltaX)) {
      bestDeltaX = dRight;
      bestTargetX = t;
      bestSnappedX = t.pos - width;
    }
  }

  // Equal spacing horizontally among siblings
  if (siblings.length >= 2) {
    const sortedH = [...siblings].sort((a, b) => a.x - b.x);
    for (let i = 0; i < sortedH.length - 1; i++) {
      const prev = sortedH[i]!;
      const next = sortedH[i + 1]!;
      const gap = next.x - (prev.x + prev.width);
      if (gap > 0) {
        // Test placing candidate after next with same gap: candidate.x = next.x + next.width + gap
        const targetPosAfter = next.x + next.width + gap;
        const dAfter = targetPosAfter - left;
        if (Math.abs(dAfter) <= threshold && Math.abs(dAfter) < Math.abs(bestDeltaX)) {
          bestDeltaX = dAfter;
          bestSnappedX = targetPosAfter;
          bestTargetX = {
            pos: targetPosAfter,
            targetType: "margin",
            spanY: [Math.min(next.y, y), Math.max(next.y + next.height, y + height)],
          };
        }
        // Test placing candidate before prev with same gap: candidate.x + candidate.width = prev.x - gap
        const targetPosBefore = prev.x - gap - width;
        const dBefore = targetPosBefore - left;
        if (Math.abs(dBefore) <= threshold && Math.abs(dBefore) < Math.abs(bestDeltaX)) {
          bestDeltaX = dBefore;
          bestSnappedX = targetPosBefore;
          bestTargetX = {
            pos: targetPosBefore + width,
            targetType: "margin",
            spanY: [Math.min(prev.y, y), Math.max(prev.y + prev.height, y + height)],
          };
        }
      }
    }
  }

  if (bestTargetX && Math.abs(bestDeltaX) <= threshold) {
    x = Math.round(bestSnappedX);
    snappedX = true;
    const startY = Math.min(y, bestTargetX.spanY[0]);
    const endY = Math.max(y + height, bestTargetX.spanY[1]);
    guides.push({
      orientation: "vertical",
      position: bestTargetX.pos,
      start: startY,
      end: endY,
      targetType: bestTargetX.targetType,
    });
  }

  // ── Y-Axis Snapping ──
  const top = y;
  const centerY = y + height / 2;
  const bottom = y + height;

  interface YTarget {
    pos: number;
    targetType: "margin" | "object";
    spanX: [number, number];
  }

  const yTargets: YTarget[] = [
    // Page margins
    {
      pos: margins.contentTop,
      targetType: "margin",
      spanX: [margins.contentLeft, margins.contentLeft + margins.contentWidth],
    },
    {
      pos: margins.contentTop + margins.contentHeight / 2,
      targetType: "margin",
      spanX: [margins.contentLeft, margins.contentLeft + margins.contentWidth],
    },
    {
      pos: margins.contentTop + margins.contentHeight,
      targetType: "margin",
      spanX: [margins.contentLeft, margins.contentLeft + margins.contentWidth],
    },
    // Page bounds & middle
    { pos: 0, targetType: "margin", spanX: [0, margins.pageWidth] },
    { pos: margins.pageHeight / 2, targetType: "margin", spanX: [0, margins.pageWidth] },
    { pos: margins.pageHeight, targetType: "margin", spanX: [0, margins.pageWidth] },
  ];

  for (const s of siblings) {
    const sSpanX: [number, number] = [s.x, s.x + s.width];
    yTargets.push(
      { pos: s.y, targetType: "object", spanX: sSpanX },
      { pos: s.y + s.height / 2, targetType: "object", spanX: sSpanX },
      { pos: s.y + s.height, targetType: "object", spanX: sSpanX },
    );
  }

  let bestDeltaY = Infinity;
  let bestTargetY: YTarget | null = null;
  let bestSnappedY = y;

  for (const t of yTargets) {
    // 1. Align top to target
    const dTop = t.pos - top;
    if (Math.abs(dTop) <= threshold && Math.abs(dTop) < Math.abs(bestDeltaY)) {
      bestDeltaY = dTop;
      bestTargetY = t;
      bestSnappedY = t.pos;
    }
    // 2. Align middle to target
    const dCenter = t.pos - centerY;
    if (Math.abs(dCenter) <= threshold && Math.abs(dCenter) < Math.abs(bestDeltaY)) {
      bestDeltaY = dCenter;
      bestTargetY = t;
      bestSnappedY = t.pos - height / 2;
    }
    // 3. Align bottom to target
    const dBottom = t.pos - bottom;
    if (Math.abs(dBottom) <= threshold && Math.abs(dBottom) < Math.abs(bestDeltaY)) {
      bestDeltaY = dBottom;
      bestTargetY = t;
      bestSnappedY = t.pos - height;
    }
  }

  // Equal spacing vertically among siblings
  if (siblings.length >= 2) {
    const sortedV = [...siblings].sort((a, b) => a.y - b.y);
    for (let i = 0; i < sortedV.length - 1; i++) {
      const prev = sortedV[i]!;
      const next = sortedV[i + 1]!;
      const gap = next.y - (prev.y + prev.height);
      if (gap > 0) {
        // After next
        const targetPosAfter = next.y + next.height + gap;
        const dAfter = targetPosAfter - top;
        if (Math.abs(dAfter) <= threshold && Math.abs(dAfter) < Math.abs(bestDeltaY)) {
          bestDeltaY = dAfter;
          bestSnappedY = targetPosAfter;
          bestTargetY = {
            pos: targetPosAfter,
            targetType: "margin",
            spanX: [Math.min(next.x, x), Math.max(next.x + next.width, x + width)],
          };
        }
        // Before prev
        const targetPosBefore = prev.y - gap - height;
        const dBefore = targetPosBefore - top;
        if (Math.abs(dBefore) <= threshold && Math.abs(dBefore) < Math.abs(bestDeltaY)) {
          bestDeltaY = dBefore;
          bestSnappedY = targetPosBefore;
          bestTargetY = {
            pos: targetPosBefore + height,
            targetType: "margin",
            spanX: [Math.min(prev.x, x), Math.max(prev.x + prev.width, x + width)],
          };
        }
      }
    }
  }

  if (bestTargetY && Math.abs(bestDeltaY) <= threshold) {
    y = Math.round(bestSnappedY);
    snappedY = true;
    const startX = Math.min(x, bestTargetY.spanX[0]);
    const endX = Math.max(x + width, bestTargetY.spanX[1]);
    guides.push({
      orientation: "horizontal",
      position: bestTargetY.pos,
      start: startX,
      end: endX,
      targetType: bestTargetY.targetType,
    });
  }

  return {
    snappedBox: { x, y, width, height },
    guides,
    snappedX,
    snappedY,
  };
}

/**
 * Overlay renderer for smart guide lines (1px dashed accent color lines).
 */
export class SmartGuidesOverlay {
  readonly el: SVGSVGElement;

  constructor() {
    this.el = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    this.el.setAttribute("data-docen-smart-guides", "");
    Object.assign(this.el.style, {
      position: "absolute",
      inset: "0",
      width: "100%",
      height: "100%",
      pointerEvents: "none",
      zIndex: "8",
      display: "none",
    } satisfies Partial<CSSStyleDeclaration>);
  }

  render(guides: readonly GuideLine[], scale = 1): void {
    while (this.el.firstChild) {
      this.el.removeChild(this.el.firstChild);
    }
    if (guides.length === 0) {
      this.el.style.display = "none";
      return;
    }
    this.el.style.display = "block";

    for (const g of guides) {
      const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
      const accent = "#2b7cd3";

      if (g.orientation === "vertical") {
        const x = g.position * scale;
        const y1 = g.start * scale;
        const y2 = g.end * scale;
        line.setAttribute("x1", String(x));
        line.setAttribute("y1", String(y1));
        line.setAttribute("x2", String(x));
        line.setAttribute("y2", String(y2));
      } else {
        const y = g.position * scale;
        const x1 = g.start * scale;
        const x2 = g.end * scale;
        line.setAttribute("x1", String(x1));
        line.setAttribute("y1", String(y));
        line.setAttribute("x2", String(x2));
        line.setAttribute("y2", String(y));
      }

      line.setAttribute("stroke", accent);
      line.setAttribute("stroke-width", "1");
      line.setAttribute("stroke-dasharray", "4,3");
      this.el.append(line);
    }
  }

  clear(): void {
    while (this.el.firstChild) {
      this.el.removeChild(this.el.firstChild);
    }
    this.el.style.display = "none";
  }
}
