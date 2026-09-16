/**
 * Mathematical Layout Engine (OMML Math Layout)
 *
 * Implements pure-geometry layout calculations for mathematical formulas
 * without DOM measurement: fractions (stacked numerator/denominator), radicals
 * (surd + overline), scripts (superscript/subscript/subsup), n-ary operators
 * (sum/integral with limits), delimiters (stretchy brackets), and matrices.
 */

export interface LayoutMathSlot {
  id?: string;
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  text?: string;
  empty?: boolean;
}

export interface LayoutMathElement {
  kind: "fraction" | "radical" | "script" | "nary" | "delimiter" | "matrix" | "text" | "slot";
  xPx: number;
  yPx: number;
  widthPx: number;
  heightPx: number;
  text?: string;
  symbol?: string;
  fractionLine?: { x1: number; y1: number; x2: number; y2: number };
  radicalBar?: { x1: number; y1: number; x2: number; y2: number; surdWidth: number };
  slots?: LayoutMathSlot[];
  children?: LayoutMathElement[];
}

export interface LayoutMathData {
  elements: LayoutMathElement[];
  widthPx: number;
  heightPx: number;
  baselinePx: number;
}

/** Standard estimate of glyph advance width based on font size. */
export function estimateCharWidth(fontSizePx: number): number {
  return fontSizePx * 0.55;
}

/** Measure a flat text string without DOM. */
export function measureMathText(
  text: string,
  fontSizePx: number,
): { widthPx: number; heightPx: number } {
  const widthPx = Math.max(8, text.length * estimateCharWidth(fontSizePx));
  const heightPx = fontSizePx * 1.2;
  return { widthPx, heightPx };
}

/** Extract text representation from a math node argument/slot. */
function extractSlotText(slot: unknown): string {
  if (typeof slot === "string") return slot;
  if (!slot || typeof slot !== "object") return "";
  const rec = slot as Record<string, unknown>;
  if (typeof rec.text === "string") return rec.text;
  if (Array.isArray(rec.children)) {
    return rec.children.map(extractSlotText).join("");
  }
  return "";
}

/**
 * Lay out a math tree from OMML-style JSON representation into coordinate-addressed
 * LayoutMathElement tree with exact dimensions and relative offsets.
 */
export function layoutMathTree(math: Record<string, unknown>, baseFontSizePx = 14): LayoutMathData {
  const elements: LayoutMathElement[] = [];
  let curX = 2;
  const children = Array.isArray(math.children) ? math.children : [math];

  let maxHeight = baseFontSizePx * 1.4;
  const baselinePx = baseFontSizePx * 0.9;

  for (const child of children) {
    if (!child || typeof child !== "object") continue;
    const rec = child as Record<string, unknown>;

    // 1. Fraction (m:f): numerator stacked over denominator
    if ("fraction" in rec && rec.fraction && typeof rec.fraction === "object") {
      const f = rec.fraction as Record<string, unknown>;
      const numText = extractSlotText(f.numerator) || "□";
      const denText = extractSlotText(f.denominator) || "□";

      const subFontSize = baseFontSizePx * 0.85;
      const numW = measureMathText(numText, subFontSize).widthPx;
      const denW = measureMathText(denText, subFontSize).widthPx;
      const fracW = Math.max(numW, denW) + 8;
      const lineH = Math.max(1, Math.round(baseFontSizePx * 0.08));
      const partH = subFontSize * 1.1;
      const totalH = partH * 2 + lineH + 4;

      const numX = (fracW - numW) / 2;
      const denX = (fracW - denW) / 2;
      const lineY = partH + 2;

      const el: LayoutMathElement = {
        kind: "fraction",
        xPx: curX,
        yPx: 0,
        widthPx: fracW,
        heightPx: totalH,
        fractionLine: { x1: 0, y1: lineY, x2: fracW, y2: lineY },
        slots: [
          {
            xPx: numX,
            yPx: 0,
            widthPx: numW,
            heightPx: partH,
            text: numText === "□" ? "" : numText,
            empty: numText === "□",
          },
          {
            xPx: denX,
            yPx: lineY + lineH + 2,
            widthPx: denW,
            heightPx: partH,
            text: denText === "□" ? "" : denText,
            empty: denText === "□",
          },
        ],
      };

      elements.push(el);
      curX += fracW + 4;
      if (totalH > maxHeight) maxHeight = totalH;
      continue;
    }

    // 2. Radical (m:rad): square root or n-th root
    if ("radical" in rec && rec.radical && typeof rec.radical === "object") {
      const rad = rec.radical as Record<string, unknown>;
      const bodyText = extractSlotText(rad.children) || "□";
      const degText = extractSlotText(rad.degree);

      const bodyW = measureMathText(bodyText, baseFontSizePx).widthPx;
      const surdWidth = baseFontSizePx * 0.7;
      const totalW = surdWidth + bodyW + 4;
      const totalH = baseFontSizePx * 1.3;

      const el: LayoutMathElement = {
        kind: "radical",
        xPx: curX,
        yPx: 0,
        widthPx: totalW,
        heightPx: totalH,
        symbol: degText || undefined,
        radicalBar: { x1: surdWidth, y1: 1, x2: totalW, y2: 1, surdWidth },
        slots: [
          {
            xPx: surdWidth + 2,
            yPx: 3,
            widthPx: bodyW,
            heightPx: totalH - 3,
            text: bodyText === "□" ? "" : bodyText,
            empty: bodyText === "□",
          },
        ],
      };

      elements.push(el);
      curX += totalW + 4;
      if (totalH > maxHeight) maxHeight = totalH;
      continue;
    }

    // 3. SuperScript / SubScript / SubSup
    if (
      ("superScript" in rec || "subScript" in rec || "subSup" in rec) &&
      typeof rec === "object"
    ) {
      const baseText =
        extractSlotText(
          rec.children ||
            (rec.superScript as Record<string, unknown>)?.children ||
            (rec.subScript as Record<string, unknown>)?.children,
        ) || "□";
      const supText = extractSlotText(
        rec.superScript && typeof rec.superScript === "object"
          ? (rec.superScript as Record<string, unknown>).superScript
          : rec.superScript,
      );
      const subText = extractSlotText(
        rec.subScript && typeof rec.subScript === "object"
          ? (rec.subScript as Record<string, unknown>).subScript
          : rec.subScript,
      );

      const baseW = measureMathText(baseText, baseFontSizePx).widthPx;
      const scriptFontSize = baseFontSizePx * 0.75;
      const supW = supText ? measureMathText(supText, scriptFontSize).widthPx : 0;
      const subW = subText ? measureMathText(subText, scriptFontSize).widthPx : 0;
      const scriptW = Math.max(supW, subW);
      const totalW = baseW + scriptW + 2;
      const totalH = baseFontSizePx * 1.4;

      const slots: LayoutMathSlot[] = [
        {
          xPx: 0,
          yPx: baseFontSizePx * 0.2,
          widthPx: baseW,
          heightPx: baseFontSizePx,
          text: baseText === "□" ? "" : baseText,
          empty: baseText === "□",
        },
      ];

      if (supText) {
        slots.push({
          xPx: baseW + 1,
          yPx: 0,
          widthPx: supW,
          heightPx: scriptFontSize,
          text: supText,
        });
      }
      if (subText) {
        slots.push({
          xPx: baseW + 1,
          yPx: baseFontSizePx * 0.6,
          widthPx: subW,
          heightPx: scriptFontSize,
          text: subText,
        });
      }

      elements.push({
        kind: "script",
        xPx: curX,
        yPx: 0,
        widthPx: totalW,
        heightPx: totalH,
        slots,
      });

      curX += totalW + 4;
      if (totalH > maxHeight) maxHeight = totalH;
      continue;
    }

    // 4. Large operator (Sum / Integral)
    if (("sum" in rec || "integral" in rec) && typeof rec === "object") {
      const isSum = "sum" in rec;
      const op = ((isSum ? rec.sum : rec.integral) as Record<string, unknown>) || {};
      const opSymbol = isSum ? "∑" : "∫";
      const bodyText = extractSlotText(op.children) || "□";
      const subText = extractSlotText(op.subScript);
      const supText = extractSlotText(op.superScript);

      const opW = baseFontSizePx * 0.9;
      const scriptFontSize = baseFontSizePx * 0.7;
      const subW = subText ? measureMathText(subText, scriptFontSize).widthPx : 0;
      const supW = supText ? measureMathText(supText, scriptFontSize).widthPx : 0;
      const bodyW = measureMathText(bodyText, baseFontSizePx).widthPx;

      const totalW = opW + Math.max(subW, supW) + bodyW + 6;
      const totalH = baseFontSizePx * 1.6;

      const slots: LayoutMathSlot[] = [];
      if (supText) {
        slots.push({ xPx: opW, yPx: 0, widthPx: supW, heightPx: scriptFontSize, text: supText });
      }
      if (subText) {
        slots.push({
          xPx: opW,
          yPx: totalH - scriptFontSize,
          widthPx: subW,
          heightPx: scriptFontSize,
          text: subText,
        });
      }
      slots.push({
        xPx: opW + Math.max(subW, supW) + 2,
        yPx: totalH * 0.25,
        widthPx: bodyW,
        heightPx: baseFontSizePx,
        text: bodyText === "□" ? "" : bodyText,
        empty: bodyText === "□",
      });

      elements.push({
        kind: "nary",
        xPx: curX,
        yPx: 0,
        widthPx: totalW,
        heightPx: totalH,
        symbol: opSymbol,
        slots,
      });

      curX += totalW + 4;
      if (totalH > maxHeight) maxHeight = totalH;
      continue;
    }

    // 5. Delimiter / Brackets
    if ("delimiter" in rec && rec.delimiter && typeof rec.delimiter === "object") {
      const d = rec.delimiter as Record<string, unknown>;
      const open = typeof d.open === "string" ? d.open : "(";
      const close = typeof d.close === "string" ? d.close : ")";
      const innerText = extractSlotText(d.children) || "□";

      const innerW = measureMathText(innerText, baseFontSizePx).widthPx;
      const bracketW = baseFontSizePx * 0.4;
      const totalW = bracketW * 2 + innerW + 4;
      const totalH = baseFontSizePx * 1.3;

      elements.push({
        kind: "delimiter",
        xPx: curX,
        yPx: 0,
        widthPx: totalW,
        heightPx: totalH,
        symbol: `${open}${close}`,
        slots: [
          {
            xPx: bracketW + 2,
            yPx: 1,
            widthPx: innerW,
            heightPx: totalH,
            text: innerText === "□" ? "" : innerText,
            empty: innerText === "□",
          },
        ],
      });

      curX += totalW + 4;
      if (totalH > maxHeight) maxHeight = totalH;
      continue;
    }

    // 6. Plain text or single variable
    const t = extractSlotText(rec);
    if (t) {
      const { widthPx, heightPx } = measureMathText(t, baseFontSizePx);
      elements.push({
        kind: "text",
        xPx: curX,
        yPx: (maxHeight - heightPx) / 2,
        widthPx,
        heightPx,
        text: t,
      });
      curX += widthPx + 2;
    }
  }

  const finalW = Math.max(16, curX + 2);
  return {
    elements,
    widthPx: finalW,
    heightPx: maxHeight,
    baselinePx,
  };
}
