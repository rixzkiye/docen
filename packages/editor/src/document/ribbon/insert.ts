import { presetShapePaths } from "@docen/docx";

import { registerIcon, type RibbonGallery, type RibbonMenuItem, type RibbonTab } from "../../ui";
import { btn, cmd, group, menu, opt, parsedItems, split, tabNode } from "./shared";

export const coverItems = (): string =>
  JSON.stringify([
    { text: opt("cover"), value: "cover", event: "cover-page" },
    { text: opt("blank"), value: "blank", event: "blank-page" },
    { text: cmd("page-break"), value: "page-break", event: "page-break" },
    { text: cmd("section-break"), value: "section-break", event: "section-break" },
  ]);

// Word's Object menu: OLE embedding is not built (greyed), but "Text from
// File" reads a plain-text file in at the caret.
export const objectItems = (): string =>
  JSON.stringify([
    { text: opt("object-dialog"), value: "object", disabled: true },
    { text: opt("file-text"), value: "file", event: "insert-file-text" },
  ]);

export const tableItems = (): string =>
  JSON.stringify([
    // Insert Table opens the classic dialog shape of the table grid entry.
    { text: opt("insert-table"), value: "insert", event: "table-dialog" },
    // Draw Table tool enables freehand grid creation & cell splitting.
    { text: opt("draw-table"), event: "draw-table" },
    { text: opt("convert-text"), value: "convert", disabled: true },
    { text: opt("excel"), value: "excel", disabled: true },
    { text: opt("quick-tables"), value: "quick", disabled: true },
  ]);

/** The Header/Footer split's drop-down: edit (the split's main action),
 *  remove, and the slot-visibility flags. Static seed only — the host
 *  re-stamps the items with live `checked` flags on every transaction
 *  (the flags live in sectionProperties, which the static schema can't
 *  read). */
export const storyMenuItems = (kind: "header" | "footer"): string =>
  JSON.stringify([
    { text: opt(kind === "header" ? "edit-header" : "edit-footer"), value: "edit" },
    {
      text: opt(kind === "header" ? "remove-header" : "remove-footer"),
      value: kind === "header" ? "remove-header" : "remove-footer",
    },
    { text: opt("different-first"), value: "title-page" },
    { text: opt("odd-even"), value: "odd-even" },
  ]);

/** The Page Number split's drop-down — Word's sections as one flat list
 *  (`-` items render as the group rules): top-of-page presets, bottom-of-page
 *  presets, current-position presets, and removal. */
export const pageNumberItems = (): string =>
  JSON.stringify([
    { text: opt("page-top-left"), value: "page-top-left" },
    { text: opt("page-top-center"), value: "page-top-center" },
    { text: opt("page-top-right"), value: "page-top-right" },
    { text: opt("page-top-bar"), value: "page-top-bar" },
    { text: "-" },
    { text: opt("page-bottom-left"), value: "page-bottom-left" },
    { text: opt("page-bottom-center"), value: "page-bottom-center" },
    { text: opt("page-bottom-right"), value: "page-bottom-right" },
    { text: opt("page-bottom-bar"), value: "page-bottom-bar" },
    { text: "-" },
    { text: opt("page-cur-plain"), value: "cur-plain" },
    { text: opt("page-cur-page-of"), value: "cur-page-of" },
    { text: opt("page-cur-page-total"), value: "cur-page-total" },
    { text: opt("page-cur-dash"), value: "cur-dash" },
    { text: opt("page-cur-slash"), value: "cur-slash" },
    { text: "-" },
    { text: opt("page-num-format"), value: "format" },
    { text: opt("remove-page-numbers"), value: "remove-numbers" },
  ]);

/** The Quick Parts split's static seed — the two organizer actions. The host
 *  re-stamps the items on every transaction with the document's building
 *  blocks grouped by gallery (they live in documentExtras, which the static
 *  schema can't read). */
export const quickPartsItems = (): RibbonMenuItem[] => [
  { text: opt("save-quick-part"), value: "save", event: "save-quick-part" },
  {
    text: opt("building-blocks-organizer"),
    value: "organizer",
    event: "building-blocks-organizer",
  },
];

/** The Equation menu: the common OMML structures as empty-argument templates
 *  (Word's equation tool's frequent structures) — the canvas paints each as a
 *  dashed placeholder slot until the math layout engine lands. */
export const equationItems = (): string =>
  JSON.stringify([
    { text: opt("equation-fraction"), value: "fraction" },
    { text: opt("equation-script"), value: "superScript" },
    { text: opt("equation-radical"), value: "radical" },
    { text: opt("equation-sum"), value: "sum" },
    { text: opt("equation-integral"), value: "integral" },
  ]);

/** Word's Shapes picker: category headings + the ST_ShapeType preset each
 *  card inserts, in the picker's order. Only tokens with an ECMA geometry
 *  definition are listed — a definition-less preset (upArrow, lineInv, …)
 *  has no preview and no path projection, so surfacing it would be a blank
 *  card and a silently-degraded insert. */
export const SHAPE_CATEGORIES: readonly { key: string; tokens: readonly string[] }[] = [
  {
    key: "lines",
    tokens: [
      "line",
      "straightConnector1",
      "bentConnector2",
      "bentConnector3",
      "bentConnector4",
      "bentConnector5",
      "curvedConnector2",
      "curvedConnector3",
      "curvedConnector4",
      "curvedConnector5",
    ],
  },
  {
    key: "rectangles",
    tokens: [
      "rect",
      "roundRect",
      "round1Rect",
      "round2SameRect",
      "round2DiagRect",
      "snip1Rect",
      "snip2SameRect",
      "snip2DiagRect",
      "snipRoundRect",
    ],
  },
  {
    key: "basic",
    tokens: [
      "ellipse",
      "triangle",
      "rtTriangle",
      "parallelogram",
      "trapezoid",
      "diamond",
      "pentagon",
      "hexagon",
      "heptagon",
      "octagon",
      "decagon",
      "dodecagon",
      "pie",
      "chord",
      "teardrop",
      "frame",
      "halfFrame",
      "corner",
      "diagStripe",
      "plus",
      "plaque",
      "can",
      "cube",
      "bevel",
      "donut",
      "noSmoking",
      "blockArc",
      "foldedCorner",
      "arc",
      "heart",
      "lightningBolt",
      "sun",
      "moon",
      "smileyFace",
      "cloud",
      "leftBracket",
      "rightBracket",
      "bracketPair",
      "leftBrace",
      "rightBrace",
      "bracePair",
      "funnel",
      "gear6",
      "gear9",
      "chartPlus",
      "chartStar",
      "chartX",
    ],
  },
  {
    key: "block-arrows",
    tokens: [
      "rightArrow",
      "leftArrow",
      "downArrow",
      "leftRightArrow",
      "upDownArrow",
      "quadArrow",
      "leftUpArrow",
      "leftRightUpArrow",
      "bentArrow",
      "bentUpArrow",
      "uturnArrow",
      "circularArrow",
      "leftCircularArrow",
      "leftRightCircularArrow",
      "curvedRightArrow",
      "curvedLeftArrow",
      "curvedUpArrow",
      "curvedDownArrow",
      "homePlate",
      "stripedRightArrow",
      "notchedRightArrow",
      "chevron",
      "swooshArrow",
    ],
  },
  {
    key: "equation",
    tokens: ["mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual", "mathNotEqual"],
  },
  {
    key: "flowchart",
    tokens: [
      "flowChartProcess",
      "flowChartAlternateProcess",
      "flowChartDecision",
      "flowChartInputOutput",
      "flowChartPredefinedProcess",
      "flowChartInternalStorage",
      "flowChartDocument",
      "flowChartMultidocument",
      "flowChartTerminator",
      "flowChartPreparation",
      "flowChartManualInput",
      "flowChartManualOperation",
      "flowChartConnector",
      "flowChartOffpageConnector",
      "flowChartPunchedCard",
      "flowChartPunchedTape",
      "flowChartSummingJunction",
      "flowChartOr",
      "flowChartCollate",
      "flowChartSort",
      "flowChartExtract",
      "flowChartMerge",
      "flowChartOnlineStorage",
      "flowChartOfflineStorage",
      "flowChartMagneticTape",
      "flowChartMagneticDisk",
      "flowChartMagneticDrum",
      "flowChartDisplay",
      "flowChartDelay",
    ],
  },
  {
    key: "stars",
    tokens: [
      "star4",
      "star5",
      "star6",
      "star7",
      "star8",
      "star10",
      "star12",
      "star16",
      "star24",
      "star32",
      "irregularSeal1",
      "irregularSeal2",
      "ribbon",
      "ribbon2",
      "leftRightRibbon",
      "ellipseRibbon",
      "ellipseRibbon2",
      "verticalScroll",
      "horizontalScroll",
      "wave",
      "doubleWave",
    ],
  },
  {
    key: "callouts",
    tokens: [
      "wedgeRectCallout",
      "wedgeRoundRectCallout",
      "wedgeEllipseCallout",
      "cloudCallout",
      "callout1",
      "callout2",
      "callout3",
      "accentCallout1",
      "accentCallout2",
      "accentCallout3",
      "borderCallout1",
      "borderCallout2",
      "borderCallout3",
      "accentBorderCallout1",
      "accentBorderCallout2",
      "accentBorderCallout3",
      "downArrowCallout",
      "leftArrowCallout",
      "rightArrowCallout",
      "upArrowCallout",
      "leftRightArrowCallout",
      "upDownArrowCallout",
      "quadArrowCallout",
    ],
  },
  {
    key: "action-buttons",
    tokens: [
      "actionButtonBackPrevious",
      "actionButtonForwardNext",
      "actionButtonBeginning",
      "actionButtonEnd",
      "actionButtonHome",
      "actionButtonInformation",
      "actionButtonReturn",
      "actionButtonMovie",
      "actionButtonDocument",
      "actionButtonSound",
      "actionButtonHelp",
      "actionButtonBlank",
    ],
  },
];

/** A shape card's thumbnail: the preset's outlines evaluated at 100×100 —
 *  pale-accent fill over a mid-accent stroke (Word's picker colors). */
export function shapePreviewSvg(token: string): string {
  const parts = (presetShapePaths(token, 100, 100) ?? []).map(
    (o) =>
      `<path d="${o.d}" fill="${o.fill ? "#DEEBF7" : "none"}" stroke="${
        o.stroke ? "#2E75B6" : "none"
      }" stroke-width="6" stroke-linejoin="round"/>`,
  );
  return `<svg viewBox="-8 -8 116 116" xmlns="http://www.w3.org/2000/svg">${parts.join("")}</svg>`;
}

export let shapeIconsRegistered = false;

export function ensureShapeIcons(): void {
  if (shapeIconsRegistered) return;
  for (const { tokens } of SHAPE_CATEGORIES) {
    for (const token of tokens) registerIcon(`shape-${token}`, shapePreviewSvg(token));
  }
  shapeIconsRegistered = true;
}

/** Word's Shapes gallery: four pinned cards in the closed strip, the More bar
 *  expanding every category as heading rows + shape cards in the drop-down. */
export const shapeGallery = (): RibbonGallery => {
  ensureShapeIcons();
  const items: RibbonMenuItem[] = [
    { text: opt("shape-rect"), icon: "shape-rect", value: "rect" },
    { text: opt("shape-roundRect"), icon: "shape-roundRect", value: "roundRect" },
    { text: opt("shape-ellipse"), icon: "shape-ellipse", value: "ellipse" },
    { text: opt("shape-line"), icon: "shape-line", value: "line" },
  ];
  for (const { key, tokens } of SHAPE_CATEGORIES) {
    items.push({ text: `ribbon.cat.shapes-${key}`, header: true, value: key });
    items.push(
      ...tokens.map((token) => ({
        text: opt(`shape-${token}`),
        icon: `shape-${token}`,
        value: token,
      })),
    );
  }
  return { type: "gallery", event: "shapes", items, visibleCount: 4 };
};

// --- Tabs --------------------------------------------------------------------

export const insertTab = (): RibbonTab =>
  tabNode("insert", [
    group("pages", [
      split("page-break", "page-break", parsedItems(coverItems()), { size: "large" }),
    ]),
    group("tables", [
      split("table-add", "insert-table", parsedItems(tableItems()), { size: "large" }),
    ]),
    group("illustrations", [
      btn("picture", "insert-picture", { size: "large" }),
      btn("online-picture", "online-picture", { size: "large" }),
      shapeGallery(),
      btn("icon-library", "icons", { size: "large" }),
      btn("3d-model", "3d-model", { size: "large" }),
      btn("smartart", "smartart", { size: "large" }),
      btn("chart", "chart", { size: "large" }),
      btn("insert-picture", "screenshot", { size: "large" }),
    ]),
    group("links", [
      btn("hyperlink", "link", { size: "large" }),
      btn("bookmark", "bookmark", { size: "large" }),
    ]),
    group("comments", [btn("comment-add", "comment", { size: "large" })]),
    group("header-footer", [
      split("header", "header", parsedItems(storyMenuItems("header")), { size: "large" }),
      split("footer", "footer", parsedItems(storyMenuItems("footer")), { size: "large" }),
      split("page-number", "page-number", parsedItems(pageNumberItems()), { size: "large" }),
    ]),
    group("text", [
      btn("text-box", "text-box", { size: "large" }),
      menu("quick-parts", "quick-parts", quickPartsItems(), { size: "large" }),
      btn("wordart", "wordart", { size: "large" }),
      btn("insert-field", "insert-field", { size: "large" }),
      btn("date-time", "date-time", { size: "large" }),
      menu("object", "object", parsedItems(objectItems()), { size: "large" }),
    ]),
    group("symbols", [
      menu("equation", "equation", parsedItems(equationItems()), { size: "large" }),
      btn("symbol", "symbol", { size: "large" }),
    ]),
  ]);
