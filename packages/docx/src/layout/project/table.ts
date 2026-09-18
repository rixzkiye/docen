// Table projection: direct tblPr over the table style — widths, insets,
// borders, per-row height/header/cantSplit, and vertical merges expanded
// into the engine's rowspan shape.

import {
  twipToPx,
  type LayoutBlock,
  type LayoutBorderEdge,
  type LayoutCellInsets,
  type LayoutTable,
  type LayoutTableWidth,
} from "@docen/layout";
import type { TableCellOptions, TableOptions } from "@office-open/docx";

import {
  indexTableStyles,
  resolveTableCellStyle,
  resolveTableLook,
  resolveTableStyle,
  type EffectiveTableCellStyle,
  type TableCellPosition,
} from "../../style-cascade";
import type { ProjectContext } from "./context";
import { eighthPtToPx, isRecord, measureTwip, num, str, type LayoutCell, type Rec } from "./guards";
import { projectChild } from "./page";
import { formatIndicatorOf, balloonKinds } from "./runs";

// ── table projection ──

function toTableWidth(w: unknown): LayoutTableWidth | undefined {
  if (!isRecord(w)) return undefined;
  if (w.type === "auto" || w.type === "nil") return undefined;
  if (w.type === "percent") {
    const size = w.size;
    const pct =
      typeof size === "string" && size.endsWith("%") ? Number(size.slice(0, -1)) : num(size);
    if (pct != null && Number.isFinite(pct)) return { type: "percent", percent: pct };
    return undefined;
  }
  const tw = measureTwip(w.size);
  return tw != null && tw > 0 ? { type: "px", px: twipToPx(tw) } : undefined;
}

function toCellInsets(m: unknown): LayoutCellInsets | undefined {
  if (!isRecord(m)) return undefined;
  const side = (v: unknown): number | undefined => {
    const size = isRecord(v) ? measureTwip(v.size) : undefined;
    return size != null ? twipToPx(size) : undefined;
  };
  const insets = {
    top: side(m.top),
    right: side(m.right),
    bottom: side(m.bottom),
    left: side(m.left),
  };
  return insets.top != null || insets.right != null || insets.bottom != null || insets.left != null
    ? insets
    : undefined;
}

/** Word's application default when neither the table nor its style declares
 *  w:tblCellMar: 108 twips left/right, 0 top/bottom. Without it cells wrap at
 *  the full column width and their text paints over the borders. */
const WORD_DEFAULT_CELL_INSETS: LayoutCellInsets = {
  left: twipToPx(108),
  right: twipToPx(108),
};

type CellBorders = NonNullable<LayoutCell["borders"]>;

/** One w:tcBorders/w:tblBorders edge → px + color (nil/none survive as declared
 *  zero-weight edges; the conflict resolver skips them). */
function toBorderEdge(v: unknown): LayoutBorderEdge | undefined {
  if (!isRecord(v)) return undefined;
  const size = num(v.size);
  const color = typeof v.color === "string" && v.color !== "auto" ? v.color : undefined;
  return {
    style: typeof v.style === "string" ? v.style : undefined,
    px: size != null ? eighthPtToPx(size) : undefined,
    color,
  };
}

function toBorders(b: unknown): CellBorders | undefined {
  if (!isRecord(b)) return undefined;
  const out = {
    top: toBorderEdge(b.top),
    right: toBorderEdge(b.right),
    bottom: toBorderEdge(b.bottom),
    left: toBorderEdge(b.left),
    tl2br: toBorderEdge(b.tl2br ?? b.topLeftToBottomRight),
    tr2bl: toBorderEdge(b.tr2bl ?? b.topRightToBottomLeft),
  };
  return out.top || out.right || out.bottom || out.left || out.tl2br || out.tr2bl ? out : undefined;
}

function toCellBorders(direct: unknown, styleBorders: unknown): CellBorders | undefined {
  const d = isRecord(direct) ? direct : undefined;
  const s = isRecord(styleBorders) ? styleBorders : undefined;
  if (!d && !s) return undefined;
  const edge = (side: string, alt?: string): LayoutBorderEdge | undefined =>
    toBorderEdge(d?.[side]) ??
    (alt ? toBorderEdge(d?.[alt]) : undefined) ??
    toBorderEdge(s?.[side]) ??
    (alt ? toBorderEdge(s?.[alt]) : undefined);
  const out = {
    top: edge("top"),
    right: edge("right"),
    bottom: edge("bottom"),
    left: edge("left"),
    tl2br: edge("tl2br", "topLeftToBottomRight"),
    tr2bl: edge("tr2bl", "topRightToBottomLeft"),
  };
  return out.top || out.right || out.bottom || out.left || out.tl2br || out.tr2bl ? out : undefined;
}

/** w:tblBorders → the engine's table-level defaults, merging the direct
 *  tblPr borders over the table style's per side. */
function toTableBorders(direct: unknown, styleTable: unknown): LayoutTable["borders"] | undefined {
  const d = isRecord(direct) ? direct : undefined;
  const s = isRecord(styleTable)
    ? isRecord(styleTable.borders)
      ? styleTable.borders
      : undefined
    : undefined;
  if (!d && !s) return undefined;
  const edge = (side: string): LayoutBorderEdge | undefined =>
    toBorderEdge(d?.[side]) ?? toBorderEdge(s?.[side]);
  const out = {
    top: edge("top"),
    bottom: edge("bottom"),
    left: edge("left"),
    right: edge("right"),
    insideHorizontal: edge("insideHorizontal"),
    insideVertical: edge("insideVertical"),
  };
  return out.top ||
    out.bottom ||
    out.left ||
    out.right ||
    out.insideHorizontal ||
    out.insideVertical
    ? out
    : undefined;
}

function projectCell(
  c: TableCellOptions,
  ctx: ProjectContext,
  rowspan?: number,
  style?: EffectiveTableCellStyle,
): LayoutCell {
  const shd = isRecord(c.shading) ? c.shading : undefined;
  const styleShd = style?.cell?.shading;
  const directFill =
    shd && typeof shd.fill === "string" && shd.fill !== "auto" && shd.type !== "nil"
      ? shd.fill
      : undefined;
  const styleFill =
    styleShd &&
    typeof styleShd.fill === "string" &&
    styleShd.fill !== "auto" &&
    styleShd.type !== "nil"
      ? styleShd.fill
      : undefined;
  const fill = directFill ?? styleFill;

  const cellCtx: ProjectContext =
    style?.paragraph || style?.run
      ? {
          ...ctx,
          tableCellDefaults: {
            paragraph: style.paragraph as Record<string, unknown> | undefined,
            run: style.run as Record<string, unknown> | undefined,
          },
        }
      : ctx;

  const verticalAlign =
    c.verticalAlign === "center" || c.verticalAlign === "bottom"
      ? c.verticalAlign
      : style?.cell?.verticalAlign === "center" || style?.cell?.verticalAlign === "bottom"
        ? style?.cell?.verticalAlign
        : undefined;

  const blocks = c.children
    .map((child) => projectChild(child, cellCtx))
    .filter((b): b is LayoutBlock => b !== null);

  const tcRev = isRecord((c as any).tcPrChange)
    ? (c as any).tcPrChange
    : isRecord(c.revision)
      ? c.revision
      : undefined;
  if (tcRev && balloonKinds(ctx).revisions) {
    const indicator = formatIndicatorOf(ctx, tcRev);
    const change = (indicator as { formatChange?: { color: string } }).formatChange;
    if (change) {
      const anchor = {
        id: num(tcRev.id) ?? 0,
        kind: "revision" as const,
        color: change.color,
        label: str(tcRev.author) ?? "",
        inlineIndex: -1,
      };
      if (blocks.length > 0 && blocks[0].kind === "paragraph") {
        blocks[0].balloons = [...(blocks[0].balloons ?? []), anchor];
      }
    }
  }

  return {
    colspan: c.columnSpan,
    rowspan: rowspan ?? 1,
    insets: toCellInsets(c.margins) ?? toCellInsets(style?.cell?.margins),
    borders: toCellBorders(c.borders, style?.cell?.borders) ?? toBorders(c.borders),
    fill,
    verticalAlign,
    textDirection:
      c.textDirection === "tbRl" || c.textDirection === "btLr" || c.textDirection === "lrTb"
        ? c.textDirection
        : undefined,
    blocks,
  };
}

/** Expand OOXML vertical merges into the layout's rowspan shape — the single
 *  projection point where the two models meet. A `restart` cell absorbs every
 *  `continue` cell below it in the same grid columns: the continuation rows
 *  drop those cells (OOXML gives them just an empty <w:p>/) and the restart's
 *  rowspan counts them. Returns the merged-cell rowspan per restart cell. */
function collectRowSpans(rows: { cells: unknown[] }[]): Map<TableCellOptions, number> {
  const spans = new Map<TableCellOptions, number>();
  // Grid column → the restart cell currently absorbing continuations below.
  const open = new Map<number, TableCellOptions>();
  for (const row of rows) {
    let col = 0;
    for (const raw of row.cells) {
      if (!isRecord(raw) || !("children" in raw)) continue;
      const cell = raw as unknown as TableCellOptions;
      const span = cell.columnSpan ?? 1;
      if (cell.verticalMerge === "continue") {
        const owner = open.get(col);
        if (owner) spans.set(owner, (spans.get(owner) ?? 1) + 1);
      } else if (cell.verticalMerge === "restart") {
        spans.set(cell, 1);
        for (let c = col; c < col + span; c++) open.set(c, cell);
        col += span;
        continue;
      } else {
        for (let c = col; c < col + span; c++) open.delete(c);
        col += span;
        continue;
      }
      col += span;
    }
  }
  return spans;
}

export function projectTable(t: TableOptions, ctx: ProjectContext): LayoutTable {
  // Only cell rows project; vMerge continuation cells fold into their restart.
  const cellRows = (t.rows ?? []).filter(
    (row): row is Extract<(typeof t.rows)[number], { cells: unknown[] }> =>
      "cells" in row && Array.isArray(row.cells),
  );
  const rowSpans = collectRowSpans(cellRows);
  const resolvedStyle = t.style ? resolveTableStyle(ctx.styles?.tableStyles, t.style) : undefined;
  const look = resolveTableLook(t.tableLook);
  const totalRows = cellRows.length;
  const totalCols =
    t.columnWidths && t.columnWidths.length > 0
      ? t.columnWidths.length
      : Math.max(
          1,
          ...cellRows.map((r) =>
            r.cells.reduce(
              (sum, c) =>
                sum +
                (isRecord(c) && "columnSpan" in c && typeof c.columnSpan === "number"
                  ? c.columnSpan
                  : 1),
              0,
            ),
          ),
        );

  const rows: LayoutTable["rows"] = [];
  for (let r = 0; r < cellRows.length; r++) {
    const row = cellRows[r]!;
    const trHeight: Rec = isRecord(row.height) ? row.height : {};
    const heightValue = measureTwip(trHeight.value);
    const height =
      heightValue != null && heightValue > 0
        ? {
            rule: trHeight.rule === "exact" ? ("exact" as const) : ("atLeast" as const),
            px: twipToPx(heightValue),
          }
        : undefined;

    const projectedCells: LayoutCell[] = [];
    let col = 0;
    for (const raw of row.cells) {
      if (!isRecord(raw) || !("children" in raw)) continue;
      const cell = raw as unknown as TableCellOptions;
      const span = cell.columnSpan ?? 1;
      if (cell.verticalMerge === "continue") {
        col += span;
        continue;
      }
      const rowSpan = rowSpans.get(cell) ?? 1;
      const cellPos: TableCellPosition = {
        rowIndex: r,
        colIndex: col,
        rowSpan,
        colSpan: span,
        totalRows,
        totalCols,
      };
      const cellStyle = resolvedStyle
        ? resolveTableCellStyle(
            resolvedStyle,
            cellPos,
            look,
            t.styleRowBandSize,
            t.styleColBandSize,
          )
        : undefined;
      projectedCells.push(projectCell(cell, ctx, rowSpan, cellStyle));
      col += span;
    }

    const trRev = isRecord((row as any).trPrChange)
      ? (row as any).trPrChange
      : isRecord((row as any).revision)
        ? (row as any).revision
        : undefined;
    if (trRev && balloonKinds(ctx).revisions) {
      const indicator = formatIndicatorOf(ctx, trRev);
      const change = (indicator as { formatChange?: { color: string } }).formatChange;
      if (change && projectedCells.length > 0) {
        const anchor = {
          id: num(trRev.id) ?? 0,
          kind: "revision" as const,
          color: change.color,
          label: str(trRev.author) ?? "",
          inlineIndex: -1,
        };
        const firstBlock = projectedCells[0]?.blocks[0];
        if (firstBlock && firstBlock.kind === "paragraph") {
          firstBlock.balloons = [...(firstBlock.balloons ?? []), anchor];
        }
      }
    }

    rows.push({
      cells: projectedCells,
      height,
      tableHeader: row.tableHeader || undefined,
      cantSplit: row.cantSplit || undefined,
    });
  }

  const columnWidthsPx = t.columnWidths?.map((w) => twipToPx(measureTwip(w) ?? 0));
  const styleTable =
    resolvedStyle?.table ??
    (t.style ? indexTableStyles(ctx.styles).get(t.style)?.table : undefined);
  const alignment = t.alignment ?? styleTable?.alignment;

  const tblRev = isRecord((t as any).tblPrChange)
    ? (t as any).tblPrChange
    : isRecord((t as any).revision)
      ? (t as any).revision
      : undefined;
  if (tblRev && balloonKinds(ctx).revisions) {
    const indicator = formatIndicatorOf(ctx, tblRev);
    const change = (indicator as { formatChange?: { color: string } }).formatChange;
    if (change && rows.length > 0) {
      const anchor = {
        id: num(tblRev.id) ?? 0,
        kind: "revision" as const,
        color: change.color,
        label: str(tblRev.author) ?? "",
        inlineIndex: -1,
      };
      const firstBlock = rows[0]?.cells[0]?.blocks[0];
      if (firstBlock && firstBlock.kind === "paragraph") {
        firstBlock.balloons = [...(firstBlock.balloons ?? []), anchor];
      }
    }
  }

  return {
    kind: "table",
    width: toTableWidth(t.width) ?? toTableWidth(styleTable?.width),
    layout: t.layout,
    align:
      alignment === "center"
        ? "center"
        : alignment === "right" || alignment === "end"
          ? "right"
          : undefined,
    columnWidthsPx: columnWidthsPx && columnWidthsPx.length > 0 ? columnWidthsPx : undefined,
    cellInsets:
      toCellInsets(t.margins) ?? toCellInsets(styleTable?.margins) ?? WORD_DEFAULT_CELL_INSETS,
    borders: toTableBorders(t.borders, styleTable),
    rows,
  };
}
