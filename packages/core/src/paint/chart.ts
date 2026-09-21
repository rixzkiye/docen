import type { LayoutDrawingMember } from "@docen/layout";

import type { ChartHitContext, ChartPartHit, ChartPartShape } from "./context";
import type { IGroup } from "./kit";
import { Ellipse, Group, Path as LeaferPath, Rect, Text } from "./kit";

// ── chart member painter ──
//
// Draws a `kind: "chart"` member from its verbatim ChartSpaceOptions payload
// (the office-open core chart model, consumed structurally — this package
// holds no office-open types). Word's default look is the target: an Office
// accent palette for series, gray axis text, hairline gridlines, a centered
// title band and a legend strip outside the plot.

type Rec = Record<string, unknown>;

const isRecord = (v: unknown): v is Rec => typeof v === "object" && v !== null && !Array.isArray(v);
const num = (v: unknown): number | undefined => (typeof v === "number" ? v : undefined);
const str = (v: unknown): string | undefined => (typeof v === "string" ? v : undefined);

/** Office theme accent palette (Word 2013+ default chart colors), order the
 *  series cycle through. */
const ACCENTS = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"];

const AXIS_TEXT = "#595959";
const GRID_LINE = "#D9D9D9";
const AXIS_LINE = "#BFBFBF";
const LABEL_PX = 12;
const TITLE_PX = 15;

/** Series fill: the explicit solid fill wins; otherwise the accent cycle. */
function seriesFillOf(series: Rec, index: number): string {
  const sp = isRecord(series.shapeProperties) ? series.shapeProperties : undefined;
  const fill = sp?.fill;
  if (isRecord(fill) && fill.type === "solid") {
    const color = isRecord(fill.color) ? (str(fill.color.hex) ?? fill.color) : str(fill.color);
    if (typeof color === "string") return color.replace("#", "").toUpperCase();
  }
  return ACCENTS[index % ACCENTS.length]!;
}

/** The chart model the painter works from, flattened out of the verbatim
 *  ChartSpaceOptions record. */
interface ChartModel {
  type: string;
  categories: string[];
  series: Rec[];
  /** "clustered" | "stacked" | "percentStacked" (bar/line/area groups). */
  grouping: string;
  markers: boolean;
  title: string | undefined;
  legend: boolean;
  legendPosition: string;
  holeSize: number;
  /** "standard" | "marker" | "filled" (radar polygons). */
  radarStyle: string;
  /** Bubble size scale, percent (100 = default max radius). */
  bubbleScale: number;
  /** What bubbleSize maps to: "area" (sqrt) or "width" (linear). */
  sizeRepresents: string;
  /** Of-Pie secondary type: "pie" | "bar". */
  ofPieType: string;
  /** Split type for ofPie: "position" | "value" | "percent" | "custom". */
  splitType: string;
  /** Split position (number of slices moved to secondary chart). */
  splitPosition: number;
  /** Stock style: "candlestick" | "standard". */
  stockStyle: string;
  dataLabels?: Rec;
  valueAxis?: Rec;
  categoryAxis?: Rec;
  secondaryValueAxis?: Rec;
  style?: number;
  wireframe?: boolean;
}

function readModel(chart: Rec): ChartModel | undefined {
  const type = str(chart.type);
  if (!type) return undefined;
  const series: Rec[] = Array.isArray(chart.series) ? chart.series.filter(isRecord) : [];
  const categories = Array.isArray(chart.categories)
    ? chart.categories.filter((c): c is string => typeof c === "string")
    : [];
  // A legend shows by default once the chart has more than one series to
  // distinguish (Word's c:autoTitleDeleted-style default; c:legend's absence
  // means the app default, not "off").
  const legend = chart.showLegend === true || (chart.showLegend === undefined && series.length > 1);

  const axes = Array.isArray(chart.axes) ? chart.axes.filter(isRecord) : [];
  const valAxes = axes.filter((a) => a.kind === "value" || (!a.kind && str(a.type) === "val"));
  const catAxes = axes.filter((a) => a.kind === "category" || (!a.kind && str(a.type) === "cat"));
  const valueAxis = isRecord(chart.valueAxis)
    ? chart.valueAxis
    : (valAxes[0] ?? (isRecord(chart.valAx) ? chart.valAx : undefined));
  const secondaryValueAxis = isRecord(chart.secondaryValueAxis)
    ? chart.secondaryValueAxis
    : (valAxes[1] ?? undefined);
  const categoryAxis = isRecord(chart.categoryAxis)
    ? chart.categoryAxis
    : (catAxes[0] ?? (isRecord(chart.catAx) ? chart.catAx : undefined));

  return {
    type,
    categories,
    series,
    grouping: str(chart.grouping) ?? "clustered",
    markers: chart.markers === true,
    title:
      typeof chart.title === "string" && chart.title
        ? chart.title
        : isRecord(chart.title) && typeof chart.title.text === "string"
          ? chart.title.text
          : undefined,
    legend,
    legendPosition: str(chart.legendPosition) ?? "bottom",
    holeSize: num(chart.holeSize) ?? 50,
    radarStyle: str(chart.radarStyle) ?? "standard",
    bubbleScale: num(chart.bubbleScale) ?? 100,
    sizeRepresents: str(chart.sizeRepresents) ?? "area",
    ofPieType: str(chart.ofPieType) ?? "pie",
    splitType: str(chart.splitType) ?? "position",
    splitPosition: num(chart.splitPosition) ?? num(chart.splitPos) ?? 2,
    stockStyle: str(chart.stockStyle) ?? "candlestick",
    dataLabels: isRecord(chart.dataLabels) ? chart.dataLabels : undefined,
    valueAxis,
    categoryAxis,
    secondaryValueAxis,
    style: num(chart.style) ?? num(chart.chartStyle),
    wireframe: chart.wireframe === true,
  };
}

/** Series values as numbers (c:val / c:yVal). Scatter/bubble series read the
 *  same slot through `values` here — the projection hands yValues over. */
function valuesOf(series: Rec): number[] {
  const raw = Array.isArray(series.values)
    ? series.values
    : Array.isArray(series.yValues)
      ? series.yValues
      : [];
  return raw.filter((v): v is number => typeof v === "number");
}

interface AxisConfig {
  min?: number;
  max?: number;
  majorUnit?: number;
  logBase?: number;
  numberFormat?: string;
}

function resolveAxisConfig(axis: Rec | undefined): AxisConfig | undefined {
  if (!axis) return undefined;
  const scaling = isRecord(axis.scaling) ? axis.scaling : undefined;
  const min = num(scaling?.min) ?? num(axis.min);
  const max = num(scaling?.max) ?? num(axis.max);
  const logBase = num(scaling?.logBase);
  const majorUnit = num(axis.majorUnit);
  const numFmt =
    str(axis.numberFormat) ??
    (isRecord(axis.numberFormat) ? str(axis.numberFormat.formatCode) : undefined);
  return { min, max, majorUnit, logBase, numberFormat: numFmt };
}

/** Nice axis bounds: a 1/2/5×10^k step covering [min, max], zero-based unless
 *  the data is entirely off zero (Word keeps a zero baseline when it can). */
function niceBounds(
  min: number,
  max: number,
  axisCfg?: AxisConfig,
): { min: number; max: number; step: number; logBase?: number } {
  if (axisCfg?.logBase) {
    const lo =
      axisCfg.min != null && axisCfg.min > 0 ? axisCfg.min : Math.max(1, min > 0 ? min : 1);
    const hi = axisCfg.max != null && axisCfg.max > lo ? axisCfg.max : Math.max(10, max);
    return { min: lo, max: hi, step: axisCfg.logBase, logBase: axisCfg.logBase };
  }
  let lo = axisCfg?.min != null ? axisCfg.min : min;
  let hi = axisCfg?.max != null ? axisCfg.max : max;
  let step = axisCfg?.majorUnit != null ? axisCfg.majorUnit : undefined;

  if (!Number.isFinite(lo) || !Number.isFinite(hi) || lo === hi) {
    return { min: lo === hi ? lo - 1 : lo, max: hi === lo ? hi + 1 : hi, step: step ?? 1 };
  }
  if (step == null) {
    const raw = (hi - lo) / 5;
    const mag = Math.pow(10, Math.floor(Math.log10(raw || 1)));
    step = ([1, 2, 5, 10] as const).map((m) => m * mag).find((s) => s >= raw) ?? 10 * mag;
    if (axisCfg?.min == null) {
      lo = Math.floor(lo / step) * step;
      if (min >= 0 && lo > 0) lo = 0;
    }
    if (axisCfg?.max == null) {
      hi = Math.ceil(hi / step) * step;
      if (max <= 0 && hi < 0) hi = 0;
    }
  }
  return { min: lo, max: hi, step: Math.max(step, 0.0001) };
}

const fmtTick = (v: number): string =>
  Math.abs(v) >= 1000 ? `${Math.round(v / 100) / 10}k` : `${Math.round(v * 100) / 100}`;

function formatValue(v: number, numberFormat?: string): string {
  if (!numberFormat) return fmtTick(v);
  if (numberFormat.includes("%")) {
    const decimals = (numberFormat.split(".")[1] || "").replace(/[^0#]/g, "").length;
    return `${(v * 100).toFixed(decimals)}%`;
  }
  if (numberFormat.includes("$") || numberFormat.includes("¥") || numberFormat.includes("€")) {
    const symbol = numberFormat.includes("$") ? "$" : numberFormat.includes("¥") ? "¥" : "€";
    return `${symbol}${Math.round(v * 100) / 100}`;
  }
  if (numberFormat.includes("0.0")) {
    const decimals = (numberFormat.split(".")[1] || "").replace(/[^0#]/g, "").length;
    return v.toFixed(decimals);
  }
  return fmtTick(v);
}

function dataLabelTextOf(
  series: Rec,
  val: number,
  catIndex: number,
  categories: string[],
  totalVal?: number,
  customDl?: Rec,
): string | undefined {
  const dl = customDl ?? (isRecord(series.dataLabels) ? series.dataLabels : undefined);
  if (!dl || dl.delete === true) return undefined;
  const numFmt = str(dl.numberFormat);
  const parts: string[] = [];
  if (dl.showSerName && typeof series.name === "string") parts.push(series.name);
  if (dl.showCatName && categories[catIndex]) parts.push(categories[catIndex]!);
  if (dl.showVal !== false && !dl.showPercent) parts.push(formatValue(val, numFmt));
  else if (dl.showVal === true) parts.push(formatValue(val, numFmt));
  if (dl.showPercent && totalVal && totalVal > 0)
    parts.push(`${Math.round((val / totalVal) * 100)}%`);
  if (parts.length === 0) parts.push(formatValue(val, numFmt));
  const separator = str(dl.separator) ?? " ";
  return parts.join(separator);
}

function dataLabelPos(series: Rec, model: ChartModel, fallback: string): string {
  const sDl = isRecord(series.dataLabels) ? series.dataLabels : undefined;
  const mDl = isRecord(model.dataLabels) ? model.dataLabels : undefined;
  return str(sDl?.position) ?? str(mDl?.position) ?? fallback;
}

function paintDataLabel(
  tree: IGroup,
  text: string,
  x: number,
  y: number,
  align: "center" | "left" | "right" = "center",
): void {
  if (!text) return;
  label(tree, text, x, y, LABEL_PX - 2, align);
}

function paintTrendlines(
  tree: IGroup,
  series: Rec,
  pts: { x: number; y: number }[],
  _plot: PlotBox,
): void {
  const trendlines = Array.isArray(series.trendlines) ? series.trendlines.filter(isRecord) : [];
  if (trendlines.length === 0 || pts.length < 2) return;

  for (const tl of trendlines) {
    const type = str(tl.type) ?? "linear";
    const n = pts.length;
    if (type === "linear") {
      let sumX = 0;
      let sumY = 0;
      let sumXY = 0;
      let sumXX = 0;
      for (const p of pts) {
        sumX += p.x;
        sumY += p.y;
        sumXY += p.x * p.y;
        sumXX += p.x * p.x;
      }
      const denom = n * sumXX - sumX * sumX;
      if (Math.abs(denom) < 1e-9) continue;
      const slope = (n * sumXY - sumX * sumY) / denom;
      const intercept = (sumY - slope * sumX) / n;
      const x0 = pts[0]!.x;
      const x1 = pts[pts.length - 1]!.x;
      const y0 = slope * x0 + intercept;
      const y1 = slope * x1 + intercept;
      tree.add(
        new LeaferPath({
          path: `M ${x0} ${y0} L ${x1} ${y1}`,
          stroke: "#595959",
          strokeWidth: 1.5,
          dashPattern: [4, 4],
        }),
      );
      if (tl.dispEq || tl.displayEquation) {
        const sign = intercept >= 0 ? "+" : "-";
        const eq = `y = ${Math.round(slope * 100) / 100}x ${sign} ${Math.round(Math.abs(intercept) * 100) / 100}`;
        label(tree, eq, x1, y1 - 8, LABEL_PX - 2, "right");
      }
    } else if (type === "movingAverage") {
      const period = Math.max(num(tl.period) ?? 2, 2);
      const maPts: { x: number; y: number }[] = [];
      for (let i = period - 1; i < n; i++) {
        let avgY = 0;
        for (let j = 0; j < period; j++) avgY += pts[i - j]!.y;
        avgY /= period;
        maPts.push({ x: pts[i]!.x, y: avgY });
      }
      if (maPts.length >= 2) {
        const path = maPts.map((p, idx) => `${idx === 0 ? "M" : "L"} ${p.x} ${p.y}`).join(" ");
        tree.add(
          new LeaferPath({
            path,
            stroke: "#595959",
            strokeWidth: 1.5,
            dashPattern: [4, 4],
          }),
        );
      }
    } else if (type === "exponential") {
      const x0 = pts[0]!.x;
      const x1 = pts[pts.length - 1]!.x;
      const steps = 10;
      const curvePts: { x: number; y: number }[] = [];
      const y0 = pts[0]!.y;
      const y1 = pts[pts.length - 1]!.y;
      for (let i = 0; i <= steps; i++) {
        const t = i / steps;
        const cx = x0 + t * (x1 - x0);
        const cy = y0 * Math.pow(Math.max(y1 / (y0 || 1), 0.1), t);
        curvePts.push({ x: cx, y: cy });
      }
      const path = curvePts.map((p, idx) => `${idx === 0 ? "M" : "L"} ${p.x} ${p.y}`).join(" ");
      tree.add(
        new LeaferPath({
          path,
          stroke: "#595959",
          strokeWidth: 1.5,
          dashPattern: [4, 4],
        }),
      );
    }
  }
}

function paintErrorBars(
  tree: IGroup,
  series: Rec,
  pts: { x: number; y: number; val: number }[],
  bounds: { min: number; max: number },
  plot: PlotBox,
): void {
  const eb = isRecord(series.errorBars) ? series.errorBars : undefined;
  if (!eb) return;
  const barType = str(eb.barType) ?? "both";
  const valueType = str(eb.valueType) ?? "fixedValue";
  const rawVal = num(eb.value) ?? 5;
  const noEndCap = eb.noEndCap === true;
  const span = bounds.max - bounds.min || 1;

  for (const p of pts) {
    let err = rawVal;
    if (valueType === "percentage") {
      err = Math.abs(p.val) * (rawVal / 100);
    }
    const errPx = (err / span) * plot.height;
    if (errPx <= 0) continue;

    if (barType === "both" || barType === "plus") {
      segment(tree, p.x, p.y, p.x, p.y - errPx, "#595959", 1.2);
      if (!noEndCap) segment(tree, p.x - 3, p.y - errPx, p.x + 3, p.y - errPx, "#595959", 1.2);
    }
    if (barType === "both" || barType === "minus") {
      segment(tree, p.x, p.y, p.x, p.y + errPx, "#595959", 1.2);
      if (!noEndCap) segment(tree, p.x - 3, p.y + errPx, p.x + 3, p.y + errPx, "#595959", 1.2);
    }
  }
}

/** Text label, horizontally centered on x (or left/right-anchored by align).
 *  With a fixed maxWidth box Leafer anchors the Text at its top-left corner,
 *  so center/right anchors shift the box to keep x on the anchor edge. */
function label(
  tree: IGroup,
  text: string,
  x: number,
  y: number,
  size: number,
  align: "center" | "left" | "right",
  maxWidth?: number,
): void {
  if (!text) return;
  const bx =
    maxWidth == null || align === "left" ? x : align === "center" ? x - maxWidth / 2 : x - maxWidth;
  tree.add(
    new Text({
      x: bx,
      y,
      text,
      fontSize: size,
      fill: AXIS_TEXT,
      textAlign: align,
      verticalAlign: "middle",
      ...(maxWidth != null ? { width: maxWidth, overflow: "ellipsis" } : {}),
    }),
  );
}

/** Gap between neighboring legend entries in a horizontal row. */
const LEGEND_GAP = 24;

/** The painted width of a label, for centering a legend row — the hidden
 *  canvas the break-row marks measure with (node-safe: without a document,
 *  a per-character estimate). */
let legendCtx: CanvasRenderingContext2D | null | undefined;
function measureLabelWidth(text: string, px: number): number {
  legendCtx ??=
    typeof document === "undefined" ? null : document.createElement("canvas").getContext("2d");
  if (!legendCtx) return text.length * px;
  legendCtx.font = `${px}px sans-serif`;
  return legendCtx.measureText(text).width;
}

/** One straight hairline. */
function segment(
  tree: IGroup,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  color: string,
  width = 1,
): void {
  tree.add(
    new LeaferPath({
      path: `M ${x0} ${y0} L ${x1} ${y1}`,
      stroke: color,
      strokeWidth: width,
    }),
  );
}

/** Scatter x/y pairs (c:xVal / c:yVal) as equal-length number arrays. */
function scatterPairs(series: Rec): { x: number[]; y: number[] } {
  const x = Array.isArray(series.xValues)
    ? series.xValues.filter((v): v is number => typeof v === "number")
    : [];
  const y = valuesOf(series);
  const n = Math.min(x.length, y.length);
  return { x: x.slice(0, n), y: y.slice(0, n) };
}

/** One plot rectangle the series render inside. */
interface PlotBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

// ── value-axis charts (column / bar / line / area) ──

/** A chart sub-element's hit registration — the box and any shape arrive in
 *  chart-local coordinates, the hit table is page-local (the registrar in
 *  {@link paintChartMember} folds the origin in). */
type ElementReg = (part: ChartPartHit, x: number, y: number, width: number, height: number) => void;

function paintValueChart(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const horizontal = model.type === "bar";
  const all = model.series.map(valuesOf);
  const flat = all.flat();
  if (flat.length === 0) return;
  const stacked = model.grouping === "stacked" || model.grouping === "percentStacked";
  // Per-category stacks resolve against the stack total (percentStacked
  // normalizes it to 100); plain clusters use the raw data range.
  const catCount = Math.max(model.categories.length, ...all.map((v) => v.length), 1);
  const axisCfg = resolveAxisConfig(model.valueAxis);
  let bounds: { min: number; max: number; step: number; logBase?: number };
  if (stacked) {
    const totals: number[] = [];
    for (let c = 0; c < catCount; c++) {
      const sum = all.reduce((acc, vals) => acc + (vals[c] ?? 0), 0);
      totals.push(model.grouping === "percentStacked" ? 100 : sum);
    }
    bounds = niceBounds(Math.min(0, ...totals), Math.max(0, ...totals), axisCfg);
  } else {
    bounds = niceBounds(Math.min(0, ...flat), Math.max(0, ...flat), axisCfg);
  }
  const toPx = (v: number): number =>
    horizontal
      ? plot.x + ((v - bounds.min) / (bounds.max - bounds.min)) * plot.width
      : plot.y + plot.height - ((v - bounds.min) / (bounds.max - bounds.min)) * plot.height;
  // The px → value inverse the editor's value-drag gesture reads (Excel's
  // drag-a-point editing). percentStacked bars paint shares, not the raw
  // values a drag would write, so they stay fixed.
  const span = bounds.max - bounds.min;
  const valueDrag =
    model.grouping === "percentStacked"
      ? undefined
      : horizontal
        ? { a: bounds.min - (plot.x * span) / plot.width, b: span / plot.width, horizontal: true }
        : {
            a: bounds.min + ((plot.y + plot.height) * span) / plot.height,
            b: -span / plot.height,
          };

  // Value axis ticks + gridlines (the category axis shares its baseline).
  const ticks: number[] = [];
  for (let v = bounds.min; v <= bounds.max + bounds.step / 2; v += bounds.step) ticks.push(v);
  for (const t of ticks) {
    const p = toPx(t);
    const tickStr = formatValue(t, axisCfg?.numberFormat);
    if (horizontal) {
      segment(tree, p, plot.y, p, plot.y + plot.height, GRID_LINE);
      label(tree, tickStr, p, plot.y + plot.height + 4, LABEL_PX - 1, "center");
    } else {
      segment(tree, plot.x, p, plot.x + plot.width, p, t === bounds.min ? AXIS_LINE : GRID_LINE);
      label(tree, tickStr, plot.x - 6, p, LABEL_PX - 1, "right");
    }
  }
  // Category axis: one band per category (labels under/left of the baseline).
  const band = (horizontal ? plot.height : plot.width) / catCount;
  const catAxisAt = toPx(Math.max(bounds.min, 0));
  for (let c = 0; c < catCount; c++) {
    const text = model.categories[c] ?? String(c + 1);
    if (horizontal) {
      const y = plot.y + (c + 0.5) * band;
      label(tree, text, plot.x - 6, y, LABEL_PX - 1, "right");
    } else {
      const x = plot.x + (c + 0.5) * band;
      label(tree, text, x, catAxisAt + 4, LABEL_PX - 1, "center", band);
    }
  }

  // Series: bars pack beside each other inside the band, lines and areas run
  // through the band midpoints. Stacking stacks instead.
  const slot = band / (stacked ? 1 : model.series.length);
  model.series.forEach((series, si) => {
    const vals = all[si]!;
    const fill = seriesFillOf(series, si);
    const isLine = model.type === "line";
    const isArea = model.type === "area";
    if (isLine || isArea) {
      const pts = vals
        .map((v, c) => ({ x: plot.x + (c + 0.5) * band, y: toPx(v), c, val: v }))
        .filter((p) => p.x >= plot.x && p.x <= plot.x + plot.width);
      if (pts.length === 0) return;
      const d = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x} ${p.y}`).join(" ");
      if (isArea) {
        const base = toPx(Math.max(bounds.min, 0));
        tree.add(
          new LeaferPath({
            path: `${d} L ${pts[pts.length - 1]!.x} ${base} L ${pts[0]!.x} ${base} Z`,
            fill: `#${fill}`,
            opacity: 0.55,
          }),
        );
      }
      tree.add(
        new LeaferPath({ path: d, stroke: `#${fill}`, strokeWidth: 2, strokeJoin: "round" }),
      );
      if (model.markers) {
        for (const p of pts) {
          tree.add(new Ellipse({ x: p.x - 3, y: p.y - 3, width: 6, height: 6, fill: `#${fill}` }));
        }
      }
      for (const p of pts) {
        const dlText = dataLabelTextOf(
          series,
          p.val,
          p.c,
          model.categories,
          undefined,
          model.dataLabels,
        );
        if (dlText) {
          const pos = dataLabelPos(series, model, "top");
          let lx = p.x;
          let ly = p.y - 8;
          let align: "center" | "left" | "right" = "center";
          if (pos === "bottom") ly = p.y + 8;
          else if (pos === "left") {
            lx = p.x - 8;
            ly = p.y;
            align = "right";
          } else if (pos === "right") {
            lx = p.x + 8;
            ly = p.y;
            align = "left";
          } else if (pos === "center") {
            lx = p.x;
            ly = p.y;
            align = "center";
          }
          paintDataLabel(tree, dlText, lx, ly, align);
        }
      }
      paintTrendlines(tree, series, pts, plot);
      paintErrorBars(tree, series, pts, bounds, plot);

      if (reg) {
        // The series shape first, the data points after — the click's
        // topmost-last scan makes a point win over the line it sits on.
        const xs = pts.map((p) => p.x);
        const ys = pts.map((p) => p.y);
        reg(
          {
            series: si,
            shape: {
              kind: "poly",
              pts: pts.map((p) => [p.x, p.y] as [number, number]),
              ...(isArea ? { closed: true } : { width: 10 }),
            },
          },
          Math.min(...xs),
          Math.min(...ys),
          Math.max(...xs) - Math.min(...xs),
          Math.max(...ys) - Math.min(...ys),
        );
        for (const p of pts) reg({ series: si, point: p.c, valueDrag }, p.x - 4, p.y - 4, 8, 8);
      }
      return;
    }
    // Column / bar rectangles.
    vals.forEach((v, c) => {
      const base = toPx(Math.max(bounds.min, 0));
      const p = toPx(v);
      if (stacked) {
        // Word stacks in series order; the running sum is the bar's far edge.
        const below = all.slice(0, si).reduce((acc, prev) => acc + (prev[c] ?? 0), 0);
        const from = toPx(below);
        const to = toPx(below + v);
        const y0 = Math.min(from, to);
        const h = Math.abs(p - toPx(below));
        const bar = horizontal
          ? { x: y0, y: plot.y + c * band, width: h, height: band }
          : { x: plot.x + c * band, y: y0, width: band, height: h };
        tree.add(new Rect({ ...bar, fill: `#${fill}` }));
        const dlText = dataLabelTextOf(series, v, c, model.categories, undefined, model.dataLabels);
        if (dlText) {
          const pos = dataLabelPos(series, model, "outsideEnd");
          let lx = bar.x + bar.width / 2;
          let ly = bar.y - 8;
          let align: "center" | "left" | "right" = "center";
          if (horizontal) {
            if (pos === "insideEnd") {
              lx = bar.x + bar.width - 4;
              ly = bar.y + bar.height / 2;
              align = "right";
            } else if (pos === "center") {
              lx = bar.x + bar.width / 2;
              ly = bar.y + bar.height / 2;
              align = "center";
            } else if (pos === "insideBase") {
              lx = bar.x + 4;
              ly = bar.y + bar.height / 2;
              align = "left";
            } else {
              lx = bar.x + bar.width + 4;
              ly = bar.y + bar.height / 2;
              align = "left";
            }
          } else {
            if (pos === "insideEnd") {
              lx = bar.x + bar.width / 2;
              ly = bar.y + 8;
            } else if (pos === "center") {
              lx = bar.x + bar.width / 2;
              ly = bar.y + bar.height / 2;
            } else if (pos === "insideBase") {
              lx = bar.x + bar.width / 2;
              ly = base - 8;
            } else {
              lx = bar.x + bar.width / 2;
              ly = bar.y - 8;
            }
          }
          paintDataLabel(tree, dlText, lx, ly, align);
        }
        reg?.({ series: si, point: c, valueDrag }, bar.x, bar.y, bar.width, bar.height);
        return;
      }
      const offset = si * slot;
      const bar = horizontal
        ? {
            x: Math.min(base, p),
            y: plot.y + c * band + offset + slot * 0.1,
            width: Math.abs(p - base),
            height: slot * 0.8,
          }
        : {
            x: plot.x + c * band + offset + slot * 0.1,
            y: Math.min(base, p),
            width: slot * 0.8,
            height: Math.abs(p - base),
          };
      tree.add(new Rect({ ...bar, fill: `#${fill}` }));
      const dlText = dataLabelTextOf(series, v, c, model.categories, undefined, model.dataLabels);
      if (dlText) {
        const pos = dataLabelPos(series, model, "outsideEnd");
        let lx = bar.x + bar.width / 2;
        let ly = bar.y - 8;
        let align: "center" | "left" | "right" = "center";
        if (horizontal) {
          if (pos === "insideEnd") {
            lx = bar.x + bar.width - 4;
            ly = bar.y + bar.height / 2;
            align = "right";
          } else if (pos === "center") {
            lx = bar.x + bar.width / 2;
            ly = bar.y + bar.height / 2;
            align = "center";
          } else if (pos === "insideBase") {
            lx = bar.x + 4;
            ly = bar.y + bar.height / 2;
            align = "left";
          } else {
            lx = bar.x + bar.width + 4;
            ly = bar.y + bar.height / 2;
            align = "left";
          }
        } else {
          if (pos === "insideEnd") {
            lx = bar.x + bar.width / 2;
            ly = bar.y + 8;
          } else if (pos === "center") {
            lx = bar.x + bar.width / 2;
            ly = bar.y + bar.height / 2;
          } else if (pos === "insideBase") {
            lx = bar.x + bar.width / 2;
            ly = base - 8;
          } else {
            lx = bar.x + bar.width / 2;
            ly = bar.y - 8;
          }
        }
        paintDataLabel(tree, dlText, lx, ly, align);
      }
      reg?.({ series: si, point: c, valueDrag }, bar.x, bar.y, bar.width, bar.height);
    });

    const pts = vals.map((v, c) => {
      const offset = si * slot;
      const x = horizontal ? toPx(v) : plot.x + c * band + offset + slot * 0.5;
      const y = horizontal ? plot.y + c * band + offset + slot * 0.5 : toPx(v);
      return { x, y, val: v };
    });
    paintTrendlines(tree, series, pts, plot);
    paintErrorBars(tree, series, pts, bounds, plot);
  });
}

// ── pie / doughnut ──

function paintPie(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const series = model.series[0];
  if (!series) return;
  const vals = valuesOf(series);
  const total = vals.reduce((a, b) => a + Math.max(0, b), 0);
  if (total <= 0) return;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const r = Math.min(plot.width, plot.height) / 2;
  const hole = model.type === "doughnut" ? (r * model.holeSize) / 100 : 0;
  let angle = -90; // Word starts the first slice at 12 o'clock.
  vals.forEach((v, i) => {
    const sweep = (Math.max(0, v) / total) * 360;
    if (sweep <= 0) return;
    const a0 = (angle * Math.PI) / 180;
    const a1 = ((angle + sweep) * Math.PI) / 180;
    angle += sweep;
    if (sweep >= 360) {
      tree.add(
        new Ellipse({
          x: cx - r,
          y: cy - r,
          width: r * 2,
          height: r * 2,
          fill: `#${seriesFillOf(series, i)}`,
        }),
      );
      const dlText = dataLabelTextOf(series, v, i, model.categories, total, model.dataLabels);
      if (dlText) paintDataLabel(tree, dlText, cx, cy);
      reg?.({ series: 0, point: i }, cx - r, cy - r, r * 2, r * 2);
      return;
    }
    const large = sweep > 180 ? 1 : 0;
    const x0 = cx + r * Math.cos(a0);
    const y0 = cy + r * Math.sin(a0);
    const x1 = cx + r * Math.cos(a1);
    const y1 = cy + r * Math.sin(a1);
    const d = hole
      ? (() => {
          const hx0 = cx + hole * Math.cos(a0);
          const hy0 = cy + hole * Math.sin(a0);
          const hx1 = cx + hole * Math.cos(a1);
          const hy1 = cy + hole * Math.sin(a1);
          return (
            `M ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} ` +
            `L ${hx1} ${hy1} A ${hole} ${hole} 0 ${large} 0 ${hx0} ${hy0} Z`
          );
        })()
      : `M ${cx} ${cy} L ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} Z`;
    tree.add(
      new LeaferPath({
        path: d,
        fill: `#${seriesFillOf(series, i)}`,
        stroke: "#FFFFFF",
        strokeWidth: 1,
      }),
    );
    const dlText = dataLabelTextOf(series, v, i, model.categories, total, model.dataLabels);
    if (dlText) {
      const pos = dataLabelPos(series, model, "bestFit");
      const mid = (a0 + a1) / 2;
      const lr =
        pos === "outsideEnd"
          ? r * 1.15
          : pos === "insideEnd"
            ? r * 0.85
            : hole
              ? (r + hole) / 2
              : r * 0.65;
      paintDataLabel(tree, dlText, cx + lr * Math.cos(mid), cy + lr * Math.sin(mid));
    }
    reg?.(
      {
        series: 0,
        point: i,
        shape: { kind: "wedge", cx, cy, r, a0, a1, ...(hole ? { hole } : {}) },
      },
      cx - r,
      cy - r,
      r * 2,
      r * 2,
    );
  });
}

// ── scatter ──

function paintScatter(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const pairs = model.series.map(scatterPairs);
  const flat = pairs.flat();
  if (flat.length === 0) return;
  const xs = flat.flatMap((p) => p.x);
  const ys = flat.flatMap((p) => p.y);
  const bx = niceBounds(Math.min(...xs), Math.max(...xs));
  const by = niceBounds(Math.min(0, ...ys), Math.max(0, ...ys), resolveAxisConfig(model.valueAxis));
  const px = (v: number) => plot.x + ((v - bx.min) / (bx.max - bx.min)) * plot.width;
  const py = (v: number) => plot.y + plot.height - ((v - by.min) / (by.max - by.min)) * plot.height;
  xyGrid(tree, bx, by, plot, px, py);
  model.series.forEach((series, si) => {
    const fill = seriesFillOf(series, si);
    const pair = pairs[si];
    if (!pair) return;
    pair.x.forEach((xv, pi) => {
      tree.add(
        new Ellipse({
          x: px(xv) - 3.5,
          y: py(pair.y[pi]) - 3.5,
          width: 7,
          height: 7,
          fill: `#${fill}`,
        }),
      );
      const dlText = dataLabelTextOf(
        series,
        pair.y[pi]!,
        pi,
        model.categories,
        undefined,
        model.dataLabels,
      );
      if (dlText) {
        paintDataLabel(tree, dlText, px(xv) + 8, py(pair.y[pi]!), "left");
      }
      reg?.({ series: si, point: pi }, px(xv) - 4.5, py(pair.y[pi]) - 4.5, 9, 9);
    });
    const pts = pair.x.map((xv, pi) => ({ x: px(xv), y: py(pair.y[pi]!), val: pair.y[pi]! }));
    paintTrendlines(tree, series, pts, plot);
    paintErrorBars(tree, series, pts, by, plot);
  });
}

const ticksOf = (b: { min: number; max: number; step: number }): number[] => {
  const out: number[] = [];
  const step = Math.max(b.step, 0.0001);
  for (let v = b.min; v <= b.max + step / 2; v += step) out.push(v);
  return out;
};

/** The scatter/bubble value grid: one hairline + tick label per nice step
 *  on both axes. */
function xyGrid(
  tree: IGroup,
  bx: { min: number; max: number; step: number },
  by: { min: number; max: number; step: number },
  plot: PlotBox,
  px: (v: number) => number,
  py: (v: number) => number,
): void {
  for (const t of ticksOf(bx)) {
    segment(tree, px(t), plot.y, px(t), plot.y + plot.height, GRID_LINE);
    label(tree, fmtTick(t), px(t), plot.y + plot.height + 4, LABEL_PX - 1, "center");
  }
  for (const t of ticksOf(by)) {
    segment(tree, plot.x, py(t), plot.x + plot.width, py(t), GRID_LINE);
    label(tree, fmtTick(t), plot.x - 6, py(t), LABEL_PX - 1, "right");
  }
}

// ── radar ──

/** Radar (c:radar): a polygonal web — one vertex per category on the
 *  inscribed circle, series as closed polygons over it (Word's web lines).
 *  Value maps linearly to the radius; a vertex carries a radial value-drag
 *  (Excel drags the vertex along its spoke). */
function paintRadar(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const all = model.series.map(valuesOf);
  const flat = all.flat();
  if (flat.length === 0) return;
  const catCount = Math.max(model.categories.length, ...all.map((v) => v.length), 3);
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const radius = Math.min(plot.width, plot.height) / 2;
  const bounds = niceBounds(
    Math.min(0, ...flat),
    Math.max(0, ...flat),
    resolveAxisConfig(model.valueAxis),
  );
  const span = bounds.max - bounds.min || 1;
  const pointAt = (c: number, v: number): [number, number] => {
    const angle = (Math.PI * 2 * c) / catCount - Math.PI / 2;
    const r = (Math.max(0, v - bounds.min) / span) * radius;
    return [cx + r * Math.cos(angle), cy + r * Math.sin(angle)];
  };
  // Concentric polygon rings per tick; spokes from the center to the outer
  // ring's vertices; category labels just outside those vertices.
  for (const t of ticksOf(bounds)) {
    const ring = Array.from({ length: catCount }, (_, c) => pointAt(c, t));
    tree.add(
      new LeaferPath({
        path: `M ${ring.map(([x, y]) => `${x} ${y}`).join(" L ")} Z`,
        stroke: GRID_LINE,
        strokeWidth: 1,
      }),
    );
  }
  const outer = Array.from({ length: catCount }, (_, c) => {
    const angle = (Math.PI * 2 * c) / catCount - Math.PI / 2;
    const [x, y] = pointAt(c, bounds.max);
    return { x, y, angle };
  });
  for (const v of outer) segment(tree, cx, cy, v.x, v.y, GRID_LINE);
  const drag = { a: bounds.min, b: span / radius, radial: { cx, cy } };
  model.series.forEach((series, si) => {
    const vals = all[si];
    if (!vals) return;
    const pts = Array.from({ length: catCount }, (_, c) => pointAt(c, vals[c] ?? bounds.min));
    const d = `M ${pts.map(([x, y]) => `${x} ${y}`).join(" L ")} Z`;
    const fill = seriesFillOf(series, si);
    if (model.radarStyle === "filled") {
      tree.add(new LeaferPath({ path: d, fill: `#${fill}`, opacity: 0.55 }));
    } else {
      tree.add(
        new LeaferPath({ path: d, stroke: `#${fill}`, strokeWidth: 2, strokeJoin: "round" }),
      );
      if (model.radarStyle === "marker" || model.markers) {
        for (const [x, y] of pts) {
          tree.add(new Ellipse({ x: x - 3, y: y - 3, width: 6, height: 6, fill: `#${fill}` }));
        }
      }
    }
    pts.forEach(([x, y], c) => {
      const dlText = dataLabelTextOf(
        series,
        vals[c] ?? 0,
        c,
        model.categories,
        undefined,
        model.dataLabels,
      );
      if (dlText) paintDataLabel(tree, dlText, x, y - 8, "center");
      reg?.({ series: si, point: c, valueDrag: drag }, x - 4, y - 4, 8, 8);
    });
  });
  outer.forEach((v, c) =>
    label(
      tree,
      model.categories[c] ?? String(c + 1),
      v.x + 12 * Math.cos(v.angle),
      v.y + 12 * Math.sin(v.angle),
      LABEL_PX - 1,
      "center",
    ),
  );
}

// ── bubble ──

/** Bubble (c:bubble): scatter points whose radius encodes c:bubbleSize —
 *  proportionally to the value (c:sizeRepresents "width") or to its square
 *  root (the default "area"), the largest at `bubbleScale`% of a quarter of
 *  the plot's short side. No value drag: the radius is a second value, not
 *  the y the pointer would move. */
function paintBubble(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const series = model.series.map((s, si) => {
    const p = scatterPairs(s);
    // Re-typed from a category chart the series carries values only — index
    // the categories instead of dropping the series.
    if (p.x.length === 0) p.x = p.y.map((_, c) => c);
    const raw = Array.isArray(s.bubbleSize)
      ? s.bubbleSize.filter((v): v is number => typeof v === "number")
      : [];
    const sizes = p.y.map((_, c) => raw[c] ?? 1);
    return { p, sizes, fill: seriesFillOf(s, si) };
  });
  const pts = series.flatMap((s) => s.p.x.map((x, c) => ({ x, y: s.p.y[c], s: s.sizes[c] })));
  if (pts.length === 0) return;
  const xs = pts.map((d) => d.x);
  const ys = pts.map((d) => d.y);
  const bx = niceBounds(Math.min(...xs), Math.max(...xs));
  const by = niceBounds(Math.min(0, ...ys), Math.max(0, ...ys), resolveAxisConfig(model.valueAxis));
  const px = (v: number) => plot.x + ((v - bx.min) / (bx.max - bx.min)) * plot.width;
  const py = (v: number) => plot.y + plot.height - ((v - by.min) / (by.max - by.min)) * plot.height;
  xyGrid(tree, bx, by, plot, px, py);
  const maxSize = Math.max(...pts.map((d) => d.s), 0);
  const maxR = ((Math.min(plot.width, plot.height) / 4) * model.bubbleScale) / 100;
  const radiusOf = (s: number): number =>
    model.sizeRepresents === "width"
      ? (maxR * s) / (maxSize || 1)
      : maxR * Math.sqrt(s / (maxSize || 1));
  series.forEach((s, si) => {
    s.p.x.forEach((xv, c) => {
      const r = radiusOf(s.sizes[c]);
      const x = px(xv);
      const y = py(s.p.y[c]);
      tree.add(
        new Ellipse({
          x: x - r,
          y: y - r,
          width: r * 2,
          height: r * 2,
          fill: `#${s.fill}`,
          opacity: 0.75,
        }),
      );
      const dlText = dataLabelTextOf(
        model.series[si]!,
        s.p.y[c]!,
        c,
        model.categories,
        undefined,
        model.dataLabels,
      );
      if (dlText) paintDataLabel(tree, dlText, x + r + 4, y, "left");
      reg?.({ series: si, point: c }, x - r - 1, y - r - 1, r * 2 + 2, r * 2 + 2);
    });
  });
}

// ── stock ──

/** Stock (c:stock): Japanese candlesticks (open/high/low/close) or high-low-close lines. */
function paintStock(tree: IGroup, model: ChartModel, plot: PlotBox): void {
  const all = model.series.map(valuesOf);
  const flat = all.flat();
  if (flat.length === 0) return;
  const catCount = Math.max(model.categories.length, ...all.map((v) => v.length), 1);
  const bounds = niceBounds(
    Math.min(0, ...flat),
    Math.max(0, ...flat),
    resolveAxisConfig(model.valueAxis),
  );
  const py = (v: number) =>
    plot.y + plot.height - ((v - bounds.min) / (bounds.max - bounds.min || 1)) * plot.height;
  for (const t of ticksOf(bounds)) {
    segment(tree, plot.x, py(t), plot.x + plot.width, py(t), GRID_LINE);
    label(tree, fmtTick(t), plot.x - 6, py(t), LABEL_PX - 1, "right");
  }
  const band = plot.width / catCount;
  const four = all.length >= 4;
  const high = (four ? all[1] : all[0]) ?? [];
  const low = (four ? all[2] : all[1]) ?? [];
  const close = (four ? all[3] : all[2]) ?? [];
  const open = four ? (all[0] ?? []) : undefined;
  for (let c = 0; c < catCount; c++) {
    const cx = plot.x + (c + 0.5) * band;
    const hi = high[c];
    const lo = low[c];
    if (hi == null || lo == null) continue;

    const op = open ? open[c] : undefined;
    const cl = close[c];

    if (op != null && cl != null) {
      // Japanese candlestick: vertical wick through center + box body
      segment(tree, cx, py(hi), cx, py(lo), "#595959", 1.5);
      const topY = Math.min(py(op), py(cl));
      const botY = Math.max(py(op), py(cl));
      const h = Math.max(botY - topY, 2);
      const w = Math.min(band * 0.55, 18);
      const isUp = cl >= op;
      tree.add(
        new Rect({
          x: cx - w / 2,
          y: topY,
          width: w,
          height: h,
          fill: isUp ? "#FFFFFF" : "#ED7D31",
          stroke: isUp ? "#26A69A" : "#595959",
          strokeWidth: 1.5,
        }),
      );
    } else {
      segment(tree, cx, py(hi), cx, py(lo), "#595959", 1.5);
      const tick = band * 0.18;
      if (open && open[c] != null)
        segment(tree, cx - tick, py(open[c]), cx, py(open[c]), "#595959", 1.5);
      if (close[c] != null)
        segment(tree, cx, py(close[c]), cx + tick, py(close[c]), "#595959", 1.5);
    }

    label(
      tree,
      model.categories[c] ?? String(c + 1),
      cx,
      plot.y + plot.height + 4,
      LABEL_PX - 1,
      "center",
    );
  }
}

// ── surface ──

/** Surface (c:surface): 3D topographical contour mesh with elevation bands and wireframe. */
function paintSurface(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const all = model.series.map(valuesOf);
  const flat = all.flat();
  if (flat.length === 0) return;
  const catCount = Math.max(model.categories.length, ...all.map((v) => v.length), 2);
  const seriesCount = Math.max(model.series.length, 2);
  const bounds = niceBounds(
    Math.min(0, ...flat),
    Math.max(0, ...flat),
    resolveAxisConfig(model.valueAxis),
  );

  // 3D Oblique parameters
  const originX = plot.x + 35;
  const originY = plot.y + plot.height - 30;
  const depthDx = (plot.width * 0.3) / Math.max(seriesCount - 1, 1);
  const depthDy = -(plot.height * 0.22) / Math.max(seriesCount - 1, 1);
  const colDx = (plot.width * 0.6) / Math.max(catCount - 1, 1);
  const valHeight = plot.height * 0.55;
  const span = bounds.max - bounds.min || 1;

  const projectPoint = (r: number, c: number, v: number): [number, number] => {
    const x = originX + c * colDx + r * depthDx;
    const normV = (v - bounds.min) / span;
    const y = originY + r * depthDy - normV * valHeight;
    return [x, y];
  };

  // Base elevation axis ticks at front left
  for (const t of ticksOf(bounds)) {
    const norm = (t - bounds.min) / span;
    const ty = originY - norm * valHeight;
    segment(tree, originX - 4, ty, originX, ty, AXIS_LINE, 1);
    label(tree, fmtTick(t), originX - 8, ty, LABEL_PX - 2, "right");
  }
  segment(tree, originX, originY, originX, originY - valHeight, AXIS_LINE, 1.5);

  // Elevation color palette
  const ELEVATION_COLORS = ["#5B9BD5", "#70AD47", "#FFC000", "#ED7D31", "#C00000"];

  // Render quads from back to front: r = seriesCount - 2 down to 0, c = 0 to catCount - 2
  for (let r = seriesCount - 2; r >= 0; r--) {
    for (let c = 0; c < catCount - 1; c++) {
      const v00 = all[r]?.[c] ?? bounds.min;
      const v01 = all[r]?.[c + 1] ?? bounds.min;
      const v11 = all[r + 1]?.[c + 1] ?? bounds.min;
      const v10 = all[r + 1]?.[c] ?? bounds.min;

      const p00 = projectPoint(r, c, v00);
      const p01 = projectPoint(r, c + 1, v01);
      const p11 = projectPoint(r + 1, c + 1, v11);
      const p10 = projectPoint(r + 1, c, v10);

      const avgV = (v00 + v01 + v11 + v10) / 4;
      const bandIdx = Math.min(
        ELEVATION_COLORS.length - 1,
        Math.max(0, Math.floor(((avgV - bounds.min) / span) * ELEVATION_COLORS.length)),
      );
      const color = ELEVATION_COLORS[bandIdx]!;
      const quadPath = `M ${p00[0]} ${p00[1]} L ${p01[0]} ${p01[1]} L ${p11[0]} ${p11[1]} L ${p10[0]} ${p10[1]} Z`;

      if (!model.wireframe) {
        tree.add(
          new LeaferPath({
            path: quadPath,
            fill: color,
            fillOpacity: 0.75,
            stroke: "#404040",
            strokeWidth: 0.75,
          }),
        );
      } else {
        tree.add(
          new LeaferPath({
            path: quadPath,
            stroke: color,
            strokeWidth: 1.2,
          }),
        );
      }

      reg?.({ series: r, point: c }, p00[0] - 4, p00[1] - 4, 8, 8);
    }
  }

  // Category labels along front edge
  for (let c = 0; c < catCount; c++) {
    const text = model.categories[c] ?? String(c + 1);
    const [x] = projectPoint(0, c, bounds.min);
    label(tree, text, x, originY + 12, LABEL_PX - 2, "center");
  }

  // Series labels along depth edge
  for (let r = 0; r < seriesCount; r++) {
    const s = model.series[r];
    const name = str(s?.name) ?? `S${r + 1}`;
    const [x, y] = projectPoint(r, 0, bounds.min);
    label(tree, name, x - 12, y + 6, LABEL_PX - 2, "right");
  }
}

// ── of-pie ──

/** Of-Pie (c:ofPie): primary pie with split slices leading to secondary pie or bar chart. */
function paintOfPie(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const series = model.series[0];
  if (!series) return;
  const vals = valuesOf(series);
  if (vals.length === 0) return;
  const total = vals.reduce((a, b) => a + Math.max(0, b), 0);
  if (total <= 0) return;

  const splitPos = Math.max(1, Math.min(model.splitPosition || 2, vals.length - 1));
  const primaryCount = vals.length - splitPos;
  const isBar = model.ofPieType === "bar";

  const pWidth = plot.width * 0.55;
  const sWidth = plot.width * 0.35;
  const cx1 = plot.x + pWidth * 0.48;
  const cy1 = plot.y + plot.height / 2;
  const r1 = Math.min(pWidth * 0.45, plot.height * 0.42);

  const primaryVals = vals.slice(0, primaryCount);
  const secondaryVals = vals.slice(primaryCount);
  const secondarySum = secondaryVals.reduce((a, b) => a + Math.max(0, b), 0);
  const primaryWithSplit = [...primaryVals, secondarySum];

  let angle = -90;
  let splitStartAngle = 0;
  let splitEndAngle = 0;

  primaryWithSplit.forEach((v, i) => {
    const sweep = (Math.max(0, v) / total) * 360;
    if (sweep <= 0) return;
    const a0 = (angle * Math.PI) / 180;
    const a1 = ((angle + sweep) * Math.PI) / 180;
    angle += sweep;
    const isSplit = i === primaryWithSplit.length - 1;
    if (isSplit) {
      splitStartAngle = a0;
      splitEndAngle = a1;
    }
    const large = sweep > 180 ? 1 : 0;
    const x0 = cx1 + r1 * Math.cos(a0);
    const y0 = cy1 + r1 * Math.sin(a0);
    const x1 = cx1 + r1 * Math.cos(a1);
    const y1 = cy1 + r1 * Math.sin(a1);
    const d = `M ${cx1} ${cy1} L ${x0} ${y0} A ${r1} ${r1} 0 ${large} 1 ${x1} ${y1} Z`;
    const fill = isSplit ? "7F7F7F" : seriesFillOf(series, i);
    tree.add(new LeaferPath({ path: d, fill: `#${fill}`, stroke: "#FFFFFF", strokeWidth: 1 }));

    const dlText = dataLabelTextOf(series, v, i, model.categories, total, model.dataLabels);
    if (dlText) {
      const mid = (a0 + a1) / 2;
      const lr = r1 * 0.7;
      paintDataLabel(tree, dlText, cx1 + lr * Math.cos(mid), cy1 + lr * Math.sin(mid));
    }
    if (!isSplit) {
      reg?.(
        { series: 0, point: i, shape: { kind: "wedge", cx: cx1, cy: cy1, r: r1, a0, a1 } },
        cx1 - r1,
        cy1 - r1,
        r1 * 2,
        r1 * 2,
      );
    }
  });

  if (isBar) {
    const barX = plot.x + plot.width - sWidth + 10;
    const barW = Math.min(sWidth * 0.5, 45);
    const barH = r1 * 1.8;
    const barY = cy1 - barH / 2;

    let curY = barY;
    secondaryVals.forEach((v, si) => {
      const origIdx = primaryCount + si;
      const h = secondarySum > 0 ? (v / secondarySum) * barH : barH / secondaryVals.length;
      const fill = seriesFillOf(series, origIdx);
      tree.add(
        new Rect({
          x: barX,
          y: curY,
          width: barW,
          height: h,
          fill: `#${fill}`,
          stroke: "#FFFFFF",
          strokeWidth: 1,
        }),
      );
      const dlText = dataLabelTextOf(series, v, origIdx, model.categories, total, model.dataLabels);
      if (dlText) paintDataLabel(tree, dlText, barX + barW / 2, curY + h / 2);
      reg?.({ series: 0, point: origIdx }, barX, curY, barW, h);
      curY += h;
    });

    const topWedgeX = cx1 + r1 * Math.cos(splitStartAngle);
    const topWedgeY = cy1 + r1 * Math.sin(splitStartAngle);
    const botWedgeX = cx1 + r1 * Math.cos(splitEndAngle);
    const botWedgeY = cy1 + r1 * Math.sin(splitEndAngle);

    segment(tree, topWedgeX, topWedgeY, barX, barY, "#8C8C8C", 1);
    segment(tree, botWedgeX, botWedgeY, barX, barY + barH, "#8C8C8C", 1);
  } else {
    const cx2 = plot.x + plot.width - sWidth * 0.6;
    const cy2 = cy1;
    const r2 = r1 * 0.75;
    let sAngle = -90;

    secondaryVals.forEach((v, si) => {
      const origIdx = primaryCount + si;
      const sweep = secondarySum > 0 ? (v / secondarySum) * 360 : 360 / secondaryVals.length;
      const a0 = (sAngle * Math.PI) / 180;
      const a1 = ((sAngle + sweep) * Math.PI) / 180;
      sAngle += sweep;
      const large = sweep > 180 ? 1 : 0;
      const x0 = cx2 + r2 * Math.cos(a0);
      const y0 = cy2 + r2 * Math.sin(a0);
      const x1 = cx2 + r2 * Math.cos(a1);
      const y1 = cy2 + r2 * Math.sin(a1);
      const d = `M ${cx2} ${cy2} L ${x0} ${y0} A ${r2} ${r2} 0 ${large} 1 ${x1} ${y1} Z`;
      const fill = seriesFillOf(series, origIdx);
      tree.add(new LeaferPath({ path: d, fill: `#${fill}`, stroke: "#FFFFFF", strokeWidth: 1 }));
      const dlText = dataLabelTextOf(series, v, origIdx, model.categories, total, model.dataLabels);
      if (dlText) {
        const mid = (a0 + a1) / 2;
        paintDataLabel(
          tree,
          dlText,
          cx2 + r2 * 0.6 * Math.cos(mid),
          cy2 + r2 * 0.6 * Math.sin(mid),
        );
      }
      reg?.(
        { series: 0, point: origIdx, shape: { kind: "wedge", cx: cx2, cy: cy2, r: r2, a0, a1 } },
        cx2 - r2,
        cy2 - r2,
        r2 * 2,
        r2 * 2,
      );
    });

    const topWedgeX = cx1 + r1 * Math.cos(splitStartAngle);
    const topWedgeY = cy1 + r1 * Math.sin(splitStartAngle);
    const botWedgeX = cx1 + r1 * Math.cos(splitEndAngle);
    const botWedgeY = cy1 + r1 * Math.sin(splitEndAngle);

    segment(tree, topWedgeX, topWedgeY, cx2, cy2 - r2, "#8C8C8C", 1);
    segment(tree, botWedgeX, botWedgeY, cx2, cy2 + r2, "#8C8C8C", 1);
  }
}

// ── combo ──

/** Combo (mixed types): column + line series on same plot or secondary axis. */
function paintCombo(tree: IGroup, model: ChartModel, plot: PlotBox, reg?: ElementReg): void {
  const all = model.series.map(valuesOf);
  const flat = all.flat();
  if (flat.length === 0) return;

  const catCount = Math.max(model.categories.length, ...all.map((v) => v.length), 1);
  const hasSecondary = model.series.some((s) => s.secondaryAxis === true || s.axis === "secondary");

  const primaryIndices = model.series
    .map((_, i) => i)
    .filter((i) => !model.series[i]?.secondaryAxis && model.series[i]?.axis !== "secondary");
  const secondaryIndices = model.series
    .map((_, i) => i)
    .filter((i) => model.series[i]?.secondaryAxis || model.series[i]?.axis === "secondary");

  const primVals = primaryIndices.flatMap((i) => all[i]!);
  const secVals = secondaryIndices.flatMap((i) => all[i]!);

  const primBounds = niceBounds(
    Math.min(0, ...primVals),
    Math.max(0, ...primVals),
    resolveAxisConfig(model.valueAxis),
  );
  const secBounds = hasSecondary
    ? niceBounds(
        Math.min(0, ...secVals),
        Math.max(0, ...secVals),
        resolveAxisConfig(model.secondaryValueAxis),
      )
    : primBounds;

  const toPxPrim = (v: number): number =>
    plot.y +
    plot.height -
    ((v - primBounds.min) / (primBounds.max - primBounds.min || 1)) * plot.height;
  const toPxSec = (v: number): number =>
    plot.y +
    plot.height -
    ((v - secBounds.min) / (secBounds.max - secBounds.min || 1)) * plot.height;

  for (const t of ticksOf(primBounds)) {
    const p = toPxPrim(t);
    segment(tree, plot.x, p, plot.x + plot.width, p, t === primBounds.min ? AXIS_LINE : GRID_LINE);
    label(tree, fmtTick(t), plot.x - 6, p, LABEL_PX - 1, "right");
  }

  if (hasSecondary) {
    segment(
      tree,
      plot.x + plot.width,
      plot.y,
      plot.x + plot.width,
      plot.y + plot.height,
      AXIS_LINE,
    );
    for (const t of ticksOf(secBounds)) {
      const p = toPxSec(t);
      segment(tree, plot.x + plot.width, p, plot.x + plot.width + 4, p, AXIS_LINE);
      label(tree, fmtTick(t), plot.x + plot.width + 8, p, LABEL_PX - 1, "left");
    }
  }

  const band = plot.width / catCount;
  const basePx = toPxPrim(Math.max(primBounds.min, 0));
  for (let c = 0; c < catCount; c++) {
    const text = model.categories[c] ?? String(c + 1);
    label(tree, text, plot.x + (c + 0.5) * band, basePx + 4, LABEL_PX - 1, "center", band);
  }

  const colSeriesIndices = model.series
    .map((s, i) => ({ s, i }))
    .filter(({ s }) => {
      const t = str(s.type);
      return !t || t === "column" || t === "bar";
    })
    .map(({ i }) => i);

  const colSlot = band / Math.max(colSeriesIndices.length, 1);

  // Pass 1: Area
  model.series.forEach((series, si) => {
    const t = str(series.type);
    if (t !== "area") return;
    const vals = all[si]!;
    const isSec = series.secondaryAxis || series.axis === "secondary";
    const toPx = isSec ? toPxSec : toPxPrim;
    const bMin = isSec ? secBounds.min : primBounds.min;
    const pts = vals.map((v, c) => ({ x: plot.x + (c + 0.5) * band, y: toPx(v) }));
    if (pts.length === 0) return;
    const d = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x} ${p.y}`).join(" ");
    const base = toPx(Math.max(bMin, 0));
    const fill = seriesFillOf(series, si);
    tree.add(
      new LeaferPath({
        path: `${d} L ${pts[pts.length - 1]!.x} ${base} L ${pts[0]!.x} ${base} Z`,
        fill: `#${fill}`,
        opacity: 0.5,
      }),
    );
  });

  // Pass 2: Columns
  colSeriesIndices.forEach((si, colIdx) => {
    const series = model.series[si]!;
    const vals = all[si]!;
    const isSec = series.secondaryAxis || series.axis === "secondary";
    const toPx = isSec ? toPxSec : toPxPrim;
    const bMin = isSec ? secBounds.min : primBounds.min;
    const base = toPx(Math.max(bMin, 0));
    const fill = seriesFillOf(series, si);
    const offset = colIdx * colSlot;

    vals.forEach((v, c) => {
      const p = toPx(v);
      const bar = {
        x: plot.x + c * band + offset + colSlot * 0.1,
        y: Math.min(base, p),
        width: colSlot * 0.8,
        height: Math.abs(p - base),
      };
      tree.add(new Rect({ ...bar, fill: `#${fill}` }));
      const dlText = dataLabelTextOf(series, v, c, model.categories, undefined, model.dataLabels);
      if (dlText) paintDataLabel(tree, dlText, bar.x + bar.width / 2, bar.y - 8);
      reg?.({ series: si, point: c }, bar.x, bar.y, bar.width, bar.height);
    });
  });

  // Pass 3: Lines, markers, trendlines, error bars
  model.series.forEach((series, si) => {
    const t = str(series.type);
    if (t !== "line" && (t || !colSeriesIndices.includes(si))) {
      if (t !== "line") return;
    }
    if (t !== "line" && colSeriesIndices.includes(si)) return;

    const vals = all[si]!;
    const isSec = series.secondaryAxis || series.axis === "secondary";
    const toPx = isSec ? toPxSec : toPxPrim;
    const pts = vals.map((v, c) => ({ x: plot.x + (c + 0.5) * band, y: toPx(v), c, val: v }));
    if (pts.length === 0) return;
    const d = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x} ${p.y}`).join(" ");
    const fill = seriesFillOf(series, si);
    tree.add(new LeaferPath({ path: d, stroke: `#${fill}`, strokeWidth: 2, strokeJoin: "round" }));

    for (const p of pts) {
      tree.add(new Ellipse({ x: p.x - 3, y: p.y - 3, width: 6, height: 6, fill: `#${fill}` }));
      const dlText = dataLabelTextOf(
        series,
        p.val,
        p.c,
        model.categories,
        undefined,
        model.dataLabels,
      );
      if (dlText) paintDataLabel(tree, dlText, p.x, p.y - 8);
      reg?.({ series: si, point: p.c }, p.x - 4, p.y - 4, 8, 8);
    }

    paintTrendlines(tree, series, pts, plot);
    paintErrorBars(tree, series, pts, isSec ? secBounds : primBounds, plot);
  });
}

// ── placeholder (fallback for unknown chart type) ──

function paintPlaceholder(tree: IGroup, model: ChartModel, plot: PlotBox): void {
  tree.add(
    new Rect({
      x: plot.x,
      y: plot.y,
      width: plot.width,
      height: plot.height,
      fill: "#F3F3F3",
      stroke: "#C4C4C4",
      strokeWidth: 1,
      strokeAlign: "center",
    }),
  );
  label(
    tree,
    model.type.toUpperCase(),
    plot.x + plot.width / 2,
    plot.y + plot.height / 2,
    LABEL_PX,
    "center",
  );
}

// ── legend / title ──

function paintLegend(tree: IGroup, model: ChartModel, box: PlotBox, reg?: ElementReg): void {
  const vertical = model.legendPosition === "right" || model.legendPosition === "left";
  // A pie's legend lists its categories (Word colors each point individually),
  // not the single series every pie has.
  const pie = model.type === "pie" || model.type === "doughnut" || model.type === "ofPie";
  const entries = pie
    ? model.categories.map((name, i) => ({
        name,
        fill: model.series[0] ? seriesFillOf(model.series[0]!, i) : ACCENTS[i % ACCENTS.length]!,
        point: i,
      }))
    : model.series.map((series, si) => ({
        name: str(series.name) ?? `系列${si + 1}`,
        fill: seriesFillOf(series, si),
        point: -1,
      }));
  if (vertical) {
    const x = model.legendPosition === "left" ? box.x : box.x + box.width - 84;
    // Word centers the legend block vertically against the plot area.
    const top = box.y + Math.max(0, (box.height - entries.length * 20) / 2);
    entries.forEach((entry, i) => {
      const y = top + i * 20;
      tree.add(new Rect({ x, y: y + 5, width: 10, height: 10, fill: `#${entry.fill}` }));
      label(tree, entry.name, x + 14, y + 10, LABEL_PX - 1, "left", 66);
      reg?.(
        { series: 0, legend: true, ...(entry.point >= 0 ? { point: entry.point } : {}) },
        x,
        y,
        84,
        20,
      );
    });
    return;
  }
  // Word lays a horizontal legend as one row of entries at natural width —
  // swatch then label — centered as a whole in its band, never spread
  // edge-to-edge.
  const y = model.legendPosition === "top" ? box.y : box.y + box.height - 18;
  const widths = entries.map((e) => 14 + measureLabelWidth(e.name, LABEL_PX - 1));
  const total =
    widths.reduce((sum, w) => sum + w, 0) + LEGEND_GAP * Math.max(0, entries.length - 1);
  let x = box.x + Math.max(0, (box.width - total) / 2);
  entries.forEach((entry, i) => {
    const w = widths[i]!;
    tree.add(new Rect({ x, y: y + 6, width: 10, height: 10, fill: `#${entry.fill}` }));
    label(tree, entry.name, x + 14, y + 11, LABEL_PX - 1, "left");
    reg?.(
      { series: 0, legend: true, ...(entry.point >= 0 ? { point: entry.point } : {}) },
      x,
      y,
      w,
      20,
    );
    x += w + LEGEND_GAP;
  });
}

/** A chart sub-element's shape re-based from chart-local to page-local px
 *  (the hit table lives in page coordinates, the painter paints in the
 *  chart's own). */
function offsetShape(shape: ChartPartShape, dx: number, dy: number): ChartPartShape {
  if (shape.kind === "wedge") return { ...shape, cx: shape.cx + dx, cy: shape.cy + dy };
  return { ...shape, pts: shape.pts.map(([px, py]) => [px + dx, py + dy] as [number, number]) };
}

// ── entry ──

/** Paint one chart member: title band, legend strip, then the plot. The
 *  whole chart renders inside the member's extent box. With `hits` the
 *  sub-elements register their click boxes (bars, points, wedges, the
 *  series line, legend entries, the title band) — Word's second-stage chart
 *  selection reads them once the chart is framed. */
export function paintChartMember(
  tree: IGroup,
  m: Extract<LayoutDrawingMember, { kind: "chart" }>,
  hits?: ChartHitContext,
): void {
  const model = isRecord(m.chart) ? readModel(m.chart) : undefined;
  if (!model || model.series.length === 0) {
    // No model — the honest empty frame the picture branch paints.
    tree.add(
      new Rect({
        x: m.x,
        y: m.y,
        width: m.width,
        height: m.height,
        fill: "#F3F3F3",
        stroke: "#C4C4C4",
        strokeWidth: 1,
        strokeAlign: "center",
      }),
    );
    return;
  }
  const chart = new Group({ x: m.x, y: m.y, width: m.width, height: m.height });
  tree.add(chart);
  const reg: ElementReg | undefined = hits
    ? (part, x, y, width, height) => {
        // The value-drag map reads page-local px but its affine coefficients
        // are chart-local: value = a + b·(page − origin), so the origin rides
        // into the intercept as −b·origin — the same fold the box and the
        // exact shapes get here. A radial map is translation-invariant in a/b;
        // its center is a position and folds the origin directly.
        const drag = part.valueDrag;
        let chartPart = part;
        if (drag) {
          const dx = hits.ox + m.x;
          const dy = hits.oy + m.y;
          chartPart = drag.radial
            ? {
                ...part,
                valueDrag: {
                  ...drag,
                  radial: { cx: drag.radial.cx + dx, cy: drag.radial.cy + dy },
                },
              }
            : { ...part, valueDrag: { ...drag, a: drag.a - drag.b * (drag.horizontal ? dx : dy) } };
        }
        hits.ctx.hitBoxes?.push({
          page: hits.ctx.pageIndex,
          x: hits.ox + m.x + x,
          y: hits.oy + m.y + y,
          width,
          height,
          para: hits.para,
          index: hits.index,
          kind: hits.kind,
          chartPart: chartPart.shape
            ? { ...chartPart, shape: offsetShape(chartPart.shape, hits.ox + m.x, hits.oy + m.y) }
            : chartPart,
          ...(hits.ctx.layer === "behind" ? { behind: true } : {}),
        });
      }
    : undefined;
  let top = 0;
  if (model.title) {
    // No width box: the title renders at its natural size anchored center —
    // Word never truncates a chart title, so nothing here may clip it. The
    // label's verticalAlign lift (~6px of ink above the y anchor) is budgeted
    // into the y so the glyphs clear the chart's top clip (inline charts paint
    // inside a clipped holder box — ink above y=0 is cut).
    label(chart, model.title, m.width / 2, top + 10, TITLE_PX, "center");
    reg?.({ title: true }, 0, 0, m.width, TITLE_PX + 8);
    top += TITLE_PX + 8;
  }
  const pie = model.type === "pie" || model.type === "doughnut" || model.type === "ofPie";
  // A pie's legend (its categories) shows on the single series too — Word
  // defaults every pie to a legend; other charts need a second series.
  const legendSize =
    model.legend && (pie || model.series.length > 1)
      ? model.legendPosition === "right" || model.legendPosition === "left"
        ? { w: 84, h: 0 }
        : { w: 0, h: 20 }
      : { w: 0, h: 0 };
  const legendBox: PlotBox | undefined =
    legendSize.w || legendSize.h
      ? {
          // The band anchors to the edge it names — a top band sits below the
          // title — and the cross-axis spans the chart.
          x: model.legendPosition === "right" ? m.width - legendSize.w : 0,
          y: model.legendPosition === "bottom" ? m.height - legendSize.h : top,
          width: legendSize.w || m.width,
          height: legendSize.h || m.height,
        }
      : undefined;
  const plot: PlotBox = {
    x: (legendBox && model.legendPosition === "left" ? legendSize.w : 0) + 44,
    y: top + (legendBox && model.legendPosition === "top" ? legendSize.h : 0) + 6,
    width:
      m.width -
      44 -
      10 -
      (legendBox && model.legendPosition === "right" ? legendSize.w : 0) -
      (legendBox && model.legendPosition === "left" ? legendSize.w : 0),
    height:
      m.height -
      top -
      14 -
      (legendBox && model.legendPosition === "bottom" ? legendSize.h : 0) -
      (legendBox && model.legendPosition === "top" ? legendSize.h : 0),
  };
  if (legendBox) paintLegend(chart, model, legendBox, reg);
  switch (model.type) {
    case "column":
    case "bar":
    case "line":
    case "area":
      paintValueChart(chart, model, plot, reg);
      break;
    case "pie":
    case "doughnut":
      paintPie(chart, model, plot, reg);
      break;
    case "scatter":
      paintScatter(chart, model, plot, reg);
      break;
    case "radar":
      paintRadar(chart, model, plot, reg);
      break;
    case "bubble":
      paintBubble(chart, model, plot, reg);
      break;
    case "stock":
      paintStock(chart, model, plot);
      break;
    case "surface":
      paintSurface(chart, model, plot, reg);
      break;
    case "ofPie":
      paintOfPie(chart, model, plot, reg);
      break;
    case "combo":
      paintCombo(chart, model, plot, reg);
      break;
    default:
      paintPlaceholder(chart, model, plot);
  }
}
