import type { LaidOutMathItem, LayoutMathElement, LayoutTextStyle } from "@docen/layout";

import type { IGroup } from "./kit";
import { Group, Line, Path, Rect, Text } from "./kit";

/**
 * Paint a mathematical formula onto the Leafer canvas scene graph.
 * Renders fractions, radicals, operators, scripts, delimiters, and placeholder slots
 * with high visual fidelity without performing any DOM measurement.
 */
export function paintMath(
  tree: IGroup,
  item: LaidOutMathItem,
  originX: number,
  originY: number,
  style?: LayoutTextStyle,
): void {
  const data = item.data;
  const ink = style?.color ?? "#000000";
  const fontFamily =
    typeof style?.family === "string" ? style.family : "Cambria Math, Segoe UI Symbol, serif";
  const fontSize = style?.sizePx ?? 14;

  // Fallback: render single dashed slot with structural label if no layout data
  if (!data || !data.elements || data.elements.length === 0) {
    tree.add(
      new Rect({
        x: originX,
        y: originY,
        width: item.widthPx,
        height: item.heightPx,
        fill: "rgba(149, 166, 190, 0.12)",
        stroke: "#9aa6be",
        strokeWidth: 1,
        dashPattern: [3, 2],
        hittable: false,
      }),
    );
    tree.add(
      new Text({
        x: originX,
        y: originY,
        width: item.widthPx,
        height: item.heightPx,
        text: item.label,
        fontFamily,
        fontSize: Math.round(fontSize * 0.85),
        italic: true,
        fill: "#707d95",
        textAlign: "center",
        verticalAlign: "middle",
        hittable: false,
      }),
    );
    return;
  }

  const container = new Group({
    x: originX,
    y: originY,
    hittable: false,
  });

  for (const el of data.elements) {
    paintMathElement(container, el, ink, fontFamily, fontSize);
  }

  tree.add(container);
}

function paintMathElement(
  group: IGroup,
  el: LayoutMathElement,
  ink: string,
  fontFamily: string,
  baseFontSize: number,
): void {
  const elGroup = new Group({
    x: el.xPx,
    y: el.yPx,
    hittable: false,
  });

  // 1. Fraction
  if (el.kind === "fraction" && el.fractionLine) {
    const { x1, y1, x2, y2 } = el.fractionLine;
    elGroup.add(
      new Line({
        points: [x1, y1, x2, y2],
        stroke: ink,
        strokeWidth: Math.max(1, Math.round(baseFontSize * 0.08)),
        hittable: false,
      }),
    );
  }

  // 2. Radical
  if (el.kind === "radical") {
    const surdW = el.radicalBar?.surdWidth ?? baseFontSize * 0.7;
    // Draw radical surd shape: tick up-down-up
    const surdPath = `M 0 ${el.heightPx * 0.6} L ${surdW * 0.3} ${el.heightPx * 0.55} L ${surdW * 0.6} ${el.heightPx - 1} L ${surdW} 1 L ${el.widthPx} 1`;
    elGroup.add(
      new Path({
        path: surdPath,
        stroke: ink,
        strokeWidth: Math.max(1, Math.round(baseFontSize * 0.08)),
        hittable: false,
      }),
    );
    if (el.symbol) {
      // Degree of root (e.g. 3 for cube root)
      elGroup.add(
        new Text({
          x: 0,
          y: 0,
          text: el.symbol,
          fontFamily,
          fontSize: Math.round(baseFontSize * 0.6),
          fill: ink,
          hittable: false,
        }),
      );
    }
  }

  // 3. Large operator (Sum / Integral)
  if (el.kind === "nary" && el.symbol) {
    elGroup.add(
      new Text({
        x: 0,
        y: (el.heightPx - baseFontSize * 1.3) / 2,
        text: el.symbol,
        fontFamily,
        fontSize: Math.round(baseFontSize * 1.3),
        fill: ink,
        hittable: false,
      }),
    );
  }

  // 4. Delimiters
  if (el.kind === "delimiter" && el.symbol) {
    const open = el.symbol[0] ?? "(";
    const close = el.symbol[1] ?? ")";
    elGroup.add(
      new Text({
        x: 0,
        y: 0,
        text: open,
        fontFamily,
        fontSize: Math.round(baseFontSize * 1.2),
        fill: ink,
        hittable: false,
      }),
    );
    elGroup.add(
      new Text({
        x: el.widthPx - baseFontSize * 0.4,
        y: 0,
        text: close,
        fontFamily,
        fontSize: Math.round(baseFontSize * 1.2),
        fill: ink,
        hittable: false,
      }),
    );
  }

  // 5. Plain text
  if (el.kind === "text" && el.text) {
    elGroup.add(
      new Text({
        x: 0,
        y: 0,
        text: el.text,
        fontFamily,
        fontSize: baseFontSize,
        fontStyle: /[a-zA-Z]/.test(el.text) ? "italic" : "normal",
        fill: ink,
        hittable: false,
      }),
    );
  }

  // Render slots (empty argument box □ or slot text)
  if (el.slots) {
    for (const slot of el.slots) {
      if (slot.empty || !slot.text) {
        // Word's □ empty argument slot
        elGroup.add(
          new Rect({
            x: slot.xPx,
            y: slot.yPx,
            width: Math.max(10, slot.widthPx),
            height: Math.max(12, slot.heightPx),
            fill: "rgba(149, 166, 190, 0.08)",
            stroke: "#9aa6be",
            strokeWidth: 1,
            dashPattern: [2, 2],
            hittable: false,
          }),
        );
      } else {
        elGroup.add(
          new Text({
            x: slot.xPx,
            y: slot.yPx,
            width: slot.widthPx,
            height: slot.heightPx,
            text: slot.text,
            fontFamily,
            fontSize: Math.round(baseFontSize * (slot.heightPx < baseFontSize ? 0.8 : 1)),
            fontStyle: /[a-zA-Z]/.test(slot.text) ? "italic" : "normal",
            fill: ink,
            hittable: false,
          }),
        );
      }
    }
  }

  group.add(elGroup);
}
