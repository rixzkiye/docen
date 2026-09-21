/**
 * Scene painter — walks a laid-out page (the @docen/layout result) and builds
 * the LeaferJS tree. The layout engine owns ALL geometry; painting positions
 * what it is given and only measures painted "chrome" text (chart/marks
 * labels) with the deterministic measureChromeText helper, so the browser
 * canvas and the server PDF place it identically. Text elements carry
 * explicit width AND height — Leafer never paints an element whose height is
 * still 0.
 *
 * @module
 */
import type {
  FlowItem,
  LaidOutBalloon,
  LaidOutBlock,
  LaidOutEndnoteArea,
  LaidOutFootnoteArea,
  LaidOutStackItem,
} from "@docen/layout";
import { columnBoxesOf } from "@docen/layout";

import type { PaintColumn, PaintContext } from "./paint/context";
import { paintBreakRow, paintParagraph } from "./paint/paragraph";
import { paintTable } from "./paint/table";
import { measureChromeText } from "./paint/text-measure";

export * from "./paint/context";
export * from "./paint/kit";
export * from "./paint/image";
export * from "./paint/drawing";
export * from "./paint/table";
export * from "./paint/paragraph";
export * from "./paint/glyph-painter";

import { leaferKit, withKit } from "./paint/kit";

export function paintScene(tree: any, items: readonly FlowItem[], ctx: PaintContext): void {
  // One content-positioned group holds the item leaves (origin = the content
  // box origin) — per-item repaint hangs the same walk on per-item groups at
  // their flow positions instead (see paintItem). Deferred floats still land
  // page-local in the caller's tree, so the target rides the context.
  const kit = ctx.kit ?? leaferKit;
  withKit(kit, () => {
    const content = kit.createGroup({ x: ctx.flow.contentLeftPx, y: ctx.flow.contentTopPx });
    tree.add(content);
    const ictx: PaintContext = {
      ...ctx,
      origin: { x: ctx.flow.contentLeftPx, y: ctx.flow.contentTopPx },
      floatsTarget: { behind: tree, body: tree },
    };
    for (const item of items) paintItem(content, item, ictx);
  });
}

/** Paint one flow item relative to the group content hangs from (ctx.origin —
 *  paintScene pins it at the content box, per-item repaint at the item's flow
 *  position): leaves land group-local, so an unchanged item repaints by a
 *  group translation alone. */
export function paintItem(tree: any, item: FlowItem, ctx: PaintContext): void {
  const kit = ctx.kit ?? leaferKit;
  withKit(kit, () => {
    const x = ctx.flow.contentLeftPx + (item.xPx ?? 0) - (ctx.origin?.x ?? 0);
    const y = ctx.flow.contentTopPx + item.yPx - (ctx.origin?.y ?? 0);
    const cols = columnBoxesOf(ctx.flow.contentWidthPx, ctx.columns);
    paintBlock(
      tree,
      item.block,
      x,
      y,
      ctx,
      item.xPx != null
        ? {
            width:
              cols.find((c) => item.xPx! >= c.xPx - 0.01 && item.xPx! < c.xPx + c.widthPx)
                ?.widthPx ?? cols[0]!.widthPx,
            inCell: false,
          }
        : undefined,
    );
  });
}

/** Paint the document grid (Word's View → Gridlines): one horizontal rule
 *  every `linePitchPx` across the content box — the pitch the docGrid snaps
 *  body lines to, so the overlay shows why lines sit where they sit. */
export function paintGridlines(tree: any, ctx: PaintContext): void {
  const pitch = ctx.flow.linePitchPx;
  if (!ctx.showGridlines || !pitch || pitch <= 0) return;
  const kit = ctx.kit ?? leaferKit;
  const left = ctx.flow.contentLeftPx;
  const right = left + ctx.flow.contentWidthPx;
  for (
    let y = ctx.flow.contentTopPx;
    y <= ctx.flow.contentTopPx + ctx.flow.contentHeightPx;
    y += pitch
  ) {
    tree.add(
      kit.createLine({
        points: [left, y, right, y],
        stroke: "#c5d3ee",
        strokeWidth: 1,
        hittable: false,
      }),
    );
  }
}

/** Paint this page's footnotes at the bottom of the content box:
 *  the separator line (Word default: 2 inches = 192 px, 1px stroke) followed
 *  by each laid footnote note. */
export function paintFootnotes(
  tree: any,
  footnotes: LaidOutFootnoteArea | undefined,
  ctx: PaintContext,
): void {
  if (!footnotes || footnotes.items.length === 0) return;
  const kit = ctx.kit ?? leaferKit;
  // Footnote separator line: 10px below the top of the footnote area
  const sepY = ctx.flow.contentTopPx + footnotes.yPx + 10;
  const sepX = ctx.flow.contentLeftPx;
  tree.add(
    kit.createLine({
      points: [sepX, sepY, sepX + footnotes.separatorWidthPx, sepY],
      stroke: "#000000",
      strokeWidth: 1,
      hittable: false,
    }),
  );
  for (const item of footnotes.items) {
    paintBlock(
      tree,
      item.block,
      ctx.flow.contentLeftPx,
      ctx.flow.contentTopPx + footnotes.yPx + item.yPx,
      ctx,
    );
  }
}

/** Paint this page's endnotes (at section end or document end):
 *  the separator line (Word default: 2 inches = 192 px, 1px stroke) followed
 *  by each laid endnote note. */
export function paintEndnotes(
  tree: any,
  endnotes: LaidOutEndnoteArea | undefined,
  ctx: PaintContext,
): void {
  if (!endnotes || endnotes.items.length === 0) return;
  const kit = ctx.kit ?? leaferKit;
  // Endnote separator line: 10px below the top of the endnote area
  const sepY = ctx.flow.contentTopPx + endnotes.yPx + 10;
  const sepX = ctx.flow.contentLeftPx;
  tree.add(
    kit.createLine({
      points: [sepX, sepY, sepX + endnotes.separatorWidthPx, sepY],
      stroke: "#000000",
      strokeWidth: 1,
      hittable: false,
    }),
  );
  for (const item of endnotes.items) {
    paintBlock(
      tree,
      item.block,
      ctx.flow.contentLeftPx,
      ctx.flow.contentTopPx + endnotes.yPx + item.yPx,
      ctx,
    );
  }
}

/** Paint the section's column separator lines (w:cols/@w:sep) — one vertical
 *  line centered in each gap between neighboring columns, spanning the
 *  content box. */
export function paintColumnSeparators(tree: any, ctx: PaintContext): void {
  const cols = ctx.columns;
  if (!cols?.separate || cols.count < 2) return;
  const kit = ctx.kit ?? leaferKit;
  const boxes = columnBoxesOf(ctx.flow.contentWidthPx, cols);
  for (let i = 0; i < boxes.length - 1; i++) {
    const x = ctx.flow.contentLeftPx + boxes[i]!.xPx + boxes[i]!.widthPx + cols.spacePx / 2;
    tree.add(
      kit.createLine({
        points: [x, ctx.flow.contentTopPx, x, ctx.flow.contentTopPx + ctx.flow.contentHeightPx],
        stroke: "#000000",
        strokeWidth: 1,
        hittable: false,
      }),
    );
  }
}

/** Paint this page's line numbers (w:lnNumType) in the left margin — each
 *  number's right edge sits `distancePx` left of the text margin and its box
 *  top aligns with the counted line's text box (same strut size, so the
 *  baselines agree). Auto placement (w:distance omitted, `distancePx: null` —
 *  the OOXML default) keeps a small fixed gap from the text margin — about
 *  the midpoint of the stage's crop-mark leg, where the number reads as
 *  beside the text, not stranded mid-margin.
 *  The marks arrive pre-counted from the stage; each number is right-aligned
 *  with the deterministic chrome measurement. */
export function paintLineNumbers(tree: any, ctx: PaintContext): void {
  const ln = ctx.lineNumbers;
  if (!ln || ln.marks.length === 0) return;
  const kit = ctx.kit ?? leaferKit;
  const widest = Math.max(...ln.marks.map((m) => m.sizePx));
  const boxWidth = String(ln.marks[ln.marks.length - 1]!.num).length * widest * 0.62;
  const { distancePx } = ln.config;
  const labelX =
    distancePx != null
      ? ctx.flow.contentLeftPx - distancePx - boxWidth
      : Math.max(0, ctx.flow.contentLeftPx - 12 - boxWidth);
  for (const mark of ln.marks) {
    const text = String(mark.num);
    // Right-align explicitly (deterministic chrome width): Leafer would
    // otherwise shift by the platform font it substitutes for the canvas.
    const textWidth = measureChromeText(text, mark.sizePx);
    tree.add(
      kit.createText({
        x: labelX + boxWidth - textWidth,
        y: ctx.flow.contentTopPx + mark.yPx,
        width: boxWidth,
        height: mark.sizePx * 1.4,
        text,
        fill: "#000000",
        fontSize: mark.sizePx,
        textAlign: "left",
      }),
    );
  }
}

/** Paint a pre-laid header/footer stack at its page position. */
export function paintFurnitureStack(
  tree: any,
  stack: readonly LaidOutStackItem[],
  x: number,
  y: number,
  ctx: PaintContext,
): void {
  for (const item of stack) {
    paintBlock(tree, item.block, x, y + item.yPx, ctx);
  }
}

export function paintBlock(
  tree: any,
  block: LaidOutBlock,
  x: number,
  y: number,
  ctx: PaintContext,
  col?: PaintColumn,
): void {
  const kit = ctx.kit ?? leaferKit;
  withKit(kit, () => {
    switch (block.kind) {
      case "paragraph":
        paintParagraph(tree, block, x, y, ctx, col);
        return;
      // Only paragraphs can carry drawings; the behind pass therefore skips
      // every other block so nothing paints twice.
      case "table":
      case "placeholder":
      case "pageBreak":
        if (ctx.layer === "behind") return;
        if (block.kind === "table") paintTable(tree, block, x, y, ctx);
        else if (block.kind === "placeholder") paintPlaceholder(tree, block, x, y, ctx);
        else if (ctx.showMarks)
          paintBreakRow(tree, block, x, y, col?.width ?? ctx.flow.contentWidthPx, ctx);
        return;
      case "group":
        for (const child of block.children) {
          paintBlock(tree, child.block, x, y + child.yPx, ctx, col);
        }
        return;
    }
  });
}

function paintPlaceholder(
  tree: any,
  block: { heightPx: number; label?: string },
  x: number,
  y: number,
  ctx?: PaintContext,
): void {
  const kit = ctx?.kit ?? leaferKit;
  const width = 240;
  const height = Math.max(20, block.heightPx);
  tree.add(
    kit.createRect({
      x,
      y,
      width,
      height,
      fill: "#fafafa",
      stroke: "#d0d0d0",
      strokeWidth: 1,
      dashPattern: [4, 3],
    }),
  );
  if (block.label) {
    tree.add(
      kit.createText({
        x: x + 8,
        y: y + 4,
        text: `${block.label} (not rendered yet)`,
        fill: "#9a9a9a",
        fontFamily: "Inter, sans-serif",
        fontSize: 11,
      }),
    );
  }
}

// ── Margin balloons ──────────────────────────────────────────────────────────

/** One balloon card's hit rectangle, page-local semantic px — the stage's
 *  click scan and hover routing. */
export interface BalloonHitBox {
  id: number;
  kind: "comment" | "revision";
  x: number;
  y: number;
  width: number;
  height: number;
}

/** What {@link paintBalloons} produced: the click boxes and each card's paint
 *  group (keyed `kind:id` — the stage toggles the group's opacity on hover). */
export interface BalloonPaint {
  boxes: BalloonHitBox[];
  groups: Map<string, any>;
}

/** Balloon card metrics — must match the flow's packing constants (the flow
 *  measured the card; the painter draws inside that box and never measures). */
const BALLOON_LINE_H = 14;
const BALLOON_PAD = 8;
const BALLOON_FONT = "Inter, sans-serif";

/** The balloon a click's page-local point lands on: the topmost painted card
 *  wins (the painter walks in order, so the last box drawn is checked first). */
export function balloonAt(
  boxes: readonly BalloonHitBox[],
  lx: number,
  ly: number,
): BalloonHitBox | null {
  for (let i = boxes.length - 1; i >= 0; i--) {
    const box = boxes[i]!;
    if (lx >= box.x && lx <= box.x + box.width && ly >= box.y && ly <= box.y + box.height) {
      return box;
    }
  }
  return null;
}

/** Paint one page's right-margin balloon stack: rounded cards in the page
 *  margin (the connector meets the content edge at the anchor line), author
 *  header in the revision color plus the wrapped body lines the flow already
 *  measured. Cards paint into per-card groups so a hover can tone one without
 *  a repaint. */
export function paintBalloons(
  tree: any,
  balloons: readonly LaidOutBalloon[] | undefined,
  ctx: PaintContext,
): BalloonPaint {
  const kit = ctx.kit ?? leaferKit;
  const boxes: BalloonHitBox[] = [];
  const groups = new Map<string, any>();
  if (!balloons || balloons.length === 0) return { boxes, groups };
  for (const balloon of balloons) {
    const x = ctx.flow.contentLeftPx + balloon.xPx;
    const y = ctx.flow.contentTopPx + balloon.yPx;
    const edge = `#${balloon.color}`;
    // The connector: card's left edge (at its vertical center) back to the
    // anchored line on the content box's right edge.
    tree.add(
      kit.createLine({
        points: [
          ctx.flow.contentLeftPx + balloon.anchorXPx,
          ctx.flow.contentTopPx + balloon.anchorYPx,
          x,
          y + balloon.heightPx / 2,
        ],
        stroke: edge,
        strokeWidth: 1,
        hittable: false,
      }),
    );
    const group = kit.createGroup({ x, y });
    tree.add(group);
    group.add(
      kit.createBox({
        width: balloon.widthPx,
        height: balloon.heightPx,
        fill: "#ffffff",
        stroke: edge,
        strokeWidth: 1,
        cornerRadius: 3,
        hittable: false,
      }),
    );
    // Accent rail down the card's left edge (Word's colored markup bar).
    group.add(
      kit.createBox({
        width: 3,
        height: Math.max(1, balloon.heightPx - 2),
        y: 1,
        fill: edge,
        cornerRadius: [2, 0, 0, 2],
        hittable: false,
      }),
    );
    let cursor = BALLOON_PAD;
    if (balloon.label) {
      group.add(
        kit.createText({
          x: BALLOON_PAD,
          y: cursor - 2,
          width: Math.max(1, balloon.widthPx - BALLOON_PAD * 2),
          height: BALLOON_LINE_H,
          text: balloon.label,
          fill: edge,
          fontFamily: BALLOON_FONT,
          fontSize: 11,
          fontWeight: 700,
          textWrap: "none",
          hittable: false,
        }),
      );
      cursor += BALLOON_LINE_H;
    }
    for (const line of balloon.lines) {
      group.add(
        kit.createText({
          x: BALLOON_PAD,
          y: cursor - 2,
          width: Math.max(1, balloon.widthPx - BALLOON_PAD * 2),
          height: BALLOON_LINE_H,
          text: line,
          fill: "#3b3b3b",
          fontFamily: BALLOON_FONT,
          fontSize: 11,
          textWrap: "none",
          hittable: false,
        }),
      );
      cursor += BALLOON_LINE_H;
    }
    boxes.push({
      id: balloon.id,
      kind: balloon.kind,
      x,
      y,
      width: balloon.widthPx,
      height: balloon.heightPx,
    });
    groups.set(`${balloon.kind}:${balloon.id}`, group);
  }
  return { boxes, groups };
}
