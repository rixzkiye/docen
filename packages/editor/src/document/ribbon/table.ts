import { registerIcon, resolveLang, type RibbonGallery, type RibbonTab } from "../../ui";
import { TABLE_STYLE_PRESETS, type TableStylePreset } from "../extensions/commands";
import {
  borderItems,
  btn,
  cmd,
  col,
  combo,
  grid,
  group,
  menu,
  opt,
  parsedItems,
  picker,
  row,
  split,
  tab,
} from "./shared";

// Word's Draw Border pen (Table Design → Draw Border): style tokens reuse the
// Borders and Shading style names; the width ladder is Word's point sizes with
// the emitted value in eighth-points (tcBorders @w:sz).
export const PEN_STYLES: readonly string[] = [
  "single",
  "dotted",
  "dashed",
  "dashSmallGap",
  "dotDash",
  "double",
  "thick",
  "wave",
];

export const penStyleItems = (): string =>
  JSON.stringify(
    PEN_STYLES.map((style) => ({
      text: `bordersShading.style-${style}`,
      event: "pen-style",
      value: style,
      ...(style === "single" ? { checked: true } : {}),
    })),
  );

export const PEN_SIZES: ReadonlyArray<readonly [string, number]> = [
  ["pen-quarter-point", 2],
  ["pen-half-point", 4],
  ["pen-three-quarter-point", 6],
  ["pen-one-point", 8],
  ["pen-one-half-point", 12],
  ["pen-two-quarter-point", 18],
  ["pen-three-point", 24],
];

export const penSizeItems = (): string =>
  JSON.stringify(
    PEN_SIZES.map(([key, eighths], i) => ({
      text: opt(key),
      event: "pen-size",
      value: String(eighths),
      ...(i === 1 ? { checked: true } : {}),
    })),
  );

// The border painter split: the face toggles the armed pen; the drop-down
// holds Word's erase half (the same sweep, w:val="nil").
export const borderPainterItems = (): string =>
  JSON.stringify([{ text: opt("border-eraser"), event: "border-painter", value: "eraser" }]);

// Word's Cell Margins presets; "custom" opens the cell options dialog (not
// built yet) and stays greyed.
export const cellMarginItems = (): string =>
  JSON.stringify([
    { text: opt("cell-margin-normal"), value: "default" },
    { text: opt("cell-margin-none"), value: "none" },
    { text: opt("narrow"), value: "narrow" },
    { text: opt("wide"), value: "wide" },
    { text: opt("custom-margins"), value: "custom", disabled: true },
  ]);

// --- Contextual tabs (Word's Table Tools) ------------------------------------
// Values match the commands.ts value spaces: table-style presets, the
// table-borders sides (same as the Home border menu), and the align-cell
// 9-grid keys (top/middle/bottom × left/center/right).

// --- Table Design gallery -----------------------------------------------------
// Word renders each gallery entry as a mini-table thumbnail painted from the
// style's borders and conditional fills. We generate that thumbnail as an SVG
// from the same preset data the command stamps — one source, preview and
// result can't drift.

export function tableStylePreviewSvg(preset: TableStylePreset): string {
  const stroke = `stroke="#595959" stroke-width="1"`;
  const cell = 10;
  const x3 = 1 + cell * 3;
  const y3 = 1 + cell * 3;
  const parts: string[] = [];
  if (preset.headerFill) {
    parts.push(
      `<rect x="1" y="1" width="${cell * 3}" height="${cell}" fill="#${preset.headerFill}"/>`,
    );
  }
  if (preset.bandFill) {
    parts.push(
      `<rect x="1" y="${1 + cell * 2}" width="${cell * 3}" height="${cell}" fill="#${preset.bandFill}"/>`,
    );
  }
  const b = preset.borders ?? {};
  const on = (side: { style: string } | undefined): boolean => !!side && side.style !== "none";
  if (on(b.top)) parts.push(`<line x1="1" y1="1" x2="${x3}" y2="1" ${stroke}/>`);
  if (on(b.bottom)) parts.push(`<line x1="1" y1="${y3}" x2="${x3}" y2="${y3}" ${stroke}/>`);
  if (on(b.left)) parts.push(`<line x1="1" y1="1" x2="1" y2="${y3}" ${stroke}/>`);
  if (on(b.right)) parts.push(`<line x1="${x3}" y1="1" x2="${x3}" y2="${y3}" ${stroke}/>`);
  if (on(b.insideHorizontal)) {
    for (const i of [1, 2]) {
      parts.push(`<line x1="1" y1="${1 + cell * i}" x2="${x3}" y2="${1 + cell * i}" ${stroke}/>`);
    }
  }
  if (on(b.insideVertical)) {
    for (const i of [1, 2]) {
      parts.push(`<line x1="${1 + cell * i}" y1="1" x2="${1 + cell * i}" y2="${y3}" ${stroke}/>`);
    }
  }
  // A borderless preset still shows a faint dashed cell grid, like Word's.
  if (parts.length === 0) {
    parts.push(
      `<rect x="1" y="1" width="${cell * 3}" height="${cell * 3}" fill="none" stroke="#C8C8C8" stroke-dasharray="2,2"/>`,
    );
  }
  return `<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">${parts.join("")}</svg>`;
}

// The tblLook flags behind Word's Table Style Options checkboxes, in Word's
// order, with their checkbox label keys.
export const TABLE_LOOK_OPTIONS: readonly { flag: string; label: string }[] = [
  { flag: "firstRow", label: opt("look-header-row") },
  { flag: "lastRow", label: opt("look-total-row") },
  { flag: "bandRow", label: opt("look-banded-rows") },
  { flag: "firstCol", label: opt("look-first-column") },
  { flag: "lastCol", label: opt("look-last-column") },
  { flag: "bandCol", label: opt("look-banded-columns") },
];

export let tableStyleIconsRegistered = false;

export function ensureTableStyleIcons(): void {
  if (tableStyleIconsRegistered) return;
  for (const [id, preset] of Object.entries(TABLE_STYLE_PRESETS)) {
    registerIcon(`table-style-${id}`, tableStylePreviewSvg(preset));
  }
  tableStyleIconsRegistered = true;
}

/** Word's Table Styles gallery: the presets as icon-over-label thumbnails in a
 *  strip (first four visible), the More bar expanding every preset in the
 *  same compound shape as a drop-down grid. */
export const tableStyleGallery = (): RibbonGallery => ({
  type: "gallery",
  event: "table-style",
  label: opt("more-table-styles"),
  items: Object.keys(TABLE_STYLE_PRESETS).map((id) => ({
    icon: `table-style-${id}`,
    text: opt(`style-${id}`),
    value: id,
  })),
  visibleCount: 4,
});

export const cellAlignItems = (): string =>
  JSON.stringify(
    ["tl", "tc", "tr", "ml", "mc", "mr", "bl", "bc", "br"].map((value) => ({
      text: opt(`cell-align-${value}`),
      value,
    })),
  );

export const tableSelectItems = (): string =>
  JSON.stringify([
    { text: opt("select-table-cell"), value: "cell", event: "select-table-cell" },
    { text: opt("select-table-column"), value: "column", event: "select-table-column" },
    { text: opt("select-table-row"), value: "row", event: "select-table-row" },
    { text: opt("select-table"), value: "table" },
  ]);

export const tableDeleteItems = (): string =>
  JSON.stringify([
    // Word opens the Delete Cells dialog — needs its own dialog (not built).
    { text: opt("delete-cells"), value: "cells", event: "delete-table", disabled: true },
    { text: opt("delete-columns"), value: "columns", event: "delete-column" },
    { text: opt("delete-rows"), value: "rows", event: "delete-row" },
    { text: opt("delete-table"), value: "table" },
  ]);

/** The AutoFit split's drop-down: Word's three AutoFit modes. */
export const autofitItems = (): string =>
  JSON.stringify([
    { text: opt("autofit-contents"), value: "contents", event: "autofit-contents" },
    { text: opt("autofit-window"), value: "window", event: "autofit-window" },
    { text: opt("fixed-column-width"), value: "fixed", event: "fixed-column-width" },
  ]);

/** The locale's Cell Size unit system (Word zh shows cm, en shows inches) —
 *  resolved against the ribbon's i18n scope (the workspace carries the
 *  effective `<docen-document lang>`), shared with the host's live combo
 *  sync (#syncCellSize). */
export const useCmUnits = (scope?: Element): boolean =>
  resolveLang(scope ?? document.documentElement)
    .toLowerCase()
    .startsWith("zh");

/** A length in twips as the Cell Size string ("1.50 厘米" / '1.00"'). */
export function formatMeasureTwip(tw: number, scope?: Element): string {
  return useCmUnits(scope) ? `${(tw / 567).toFixed(2)} 厘米` : `${(tw / 1440).toFixed(2)}"`;
}

/** Word's Cell Size spinner presets — the values are the twips the commands
 *  receive; the labels follow the locale's unit system. Free-typed entries
 *  accept the same UniversalMeasure units ("1.5cm" / "0.5in"). */
export const cellWidthPresets = (scope?: Element): readonly (readonly [string, string])[] =>
  useCmUnits(scope)
    ? [
        ["0.5 厘米", "283"],
        ["1 厘米", "567"],
        ["1.5 厘米", "850"],
        ["2 厘米", "1134"],
        ["3 厘米", "1701"],
        ["4 厘米", "2268"],
      ]
    : [
        ['0.5"', "720"],
        ['0.75"', "1080"],
        ['1"', "1440"],
        ['1.5"', "2160"],
        ['2"', "2880"],
        ['3"', "4320"],
      ];

export const cellHeightPresets = (scope?: Element): readonly (readonly [string, string])[] => [
  ...(useCmUnits(scope)
    ? ([
        ["0.5 厘米", "283"],
        ["1 厘米", "567"],
        ["1.5 厘米", "850"],
        ["2 厘米", "1134"],
      ] as const)
    : ([
        ['0.25"', "360"],
        ['0.5"', "720"],
        ['0.75"', "1080"],
        ['1"', "1440"],
      ] as const)),
  [useCmUnits(scope) ? "自动" : "auto", "0"],
];

export const cellWidthItems = (scope?: Element): string =>
  JSON.stringify(cellWidthPresets(scope).map(([text, value]) => ({ text, value })));

export const cellHeightItems = (scope?: Element): string =>
  JSON.stringify(cellHeightPresets(scope).map(([text, value]) => ({ text, value })));

/** Word's contextual Table Tools — the Table Design / Table Layout tabs that
 *  appear while the caret is inside a table. Marked `contextual` so
 *  {@link ribbonTabs} excludes them from the static render; the host appends
 *  them (via {@link buildContextualTab}) when the selection enters a table and
 *  removes them when it leaves. `scope` is the i18n scope the unit-system
 *  presets (Cell Size) resolve against. */
export function tableContextTabs(scope?: Element): RibbonTab[] {
  ensureTableStyleIcons();
  return [
    {
      id: "table-design",
      label: tab("table-design"),
      contextual: true,
      groups: [
        // Word's two leading groups, in Word's order: the Table Style Options
        // checkboxes (2×3), then the Table Styles gallery.
        group("table-style-options", [
          grid(
            TABLE_LOOK_OPTIONS.map((o) => ({
              type: "checkbox" as const,
              event: "toggle-table-look",
              value: o.flag,
              label: o.label,
            })),
          ),
        ]),
        group("table-styles", [tableStyleGallery()]),
        group("table-shading", [
          {
            type: "color-picker",
            icon: "shading",
            event: "cell-shading",
            label: cmd("cell-shading"),
            defaultColor: "FFFF00",
            size: "large",
          },
        ]),
        group("table-borders", [
          split("border", "table-borders", parsedItems(borderItems()), { size: "large" }),
        ]),
        group("draw-border", [
          // Word's Draw Border tools: the pen pickers stamp the host's pen
          // state; the painter split arms the sweep (face) and holds the
          // eraser (drop-down).
          split("pen", "pen-style", parsedItems(penStyleItems()), { size: "large" }),
          split("font-size", "pen-size", parsedItems(penSizeItems()), { size: "large" }),
          picker("font-color", "pen-color", "000000", { withLabel: true, size: "large" }),
          split("format-painter", "border-painter", parsedItems(borderPainterItems()), {
            size: "large",
          }),
          btn("gridlines", "toggle-gridlines", { size: "large" }),
        ]),
      ],
    },
    {
      id: "table-layout",
      label: tab("table-layout"),
      contextual: true,
      groups: [
        group("table", [
          // Table Properties opens Word's table dialog (read arrives via
          // #onCommand, the commit via table-properties:ok — same pair as the
          // context menu's entry).
          btn("table-properties", "table-properties", { size: "large" }),
          // Word stacks Select over Delete — two small menus, one column.
          col([
            menu("table-cursor", "select-table", parsedItems(tableSelectItems())),
            menu("table-delete", "delete-table", parsedItems(tableDeleteItems())),
          ]),
        ]),
        group("draw", [
          btn("pen", "draw-table", { size: "large" }),
          btn("eraser", "table-eraser", { size: "large" }),
        ]),
        group("rows-columns", [
          row([
            col([
              btn("table-stack-above", "insert-row-above"),
              btn("table-stack-below", "insert-row-below"),
            ]),
            col([
              btn("table-stack-left", "insert-column-left"),
              btn("table-stack-right", "insert-column-right"),
            ]),
          ]),
        ]),
        group("merge", [
          btn("merge-cells", "merge-cells", { size: "large" }),
          btn("split-cells", "split-cell", { size: "large" }),
          btn("table-simple", "split-table", { size: "large" }),
        ]),
        group("cell-size", [
          split("autofit", "autofit", parsedItems(autofitItems()), { size: "large" }),
          col([
            combo("cell-height", "auto", parsedItems(cellHeightItems(scope)), {
              comboboxSize: "short",
            }),
            combo("cell-width", '1"', parsedItems(cellWidthItems(scope)), {
              comboboxSize: "short",
            }),
          ]),
          col([
            grid([
              btn("distribute-rows", "distribute-rows"),
              btn("distribute-columns", "distribute-columns"),
            ]),
          ]),
        ]),
        group("alignment", [
          split("align-center", "align-cell", parsedItems(cellAlignItems()), { size: "large" }),
          btn("text-direction", "text-direction", { size: "large" }),
          // Word's Cell Margins menu button: the presets stamp the caret
          // cell's tcMar; Custom opens the options dialog (not built).
          menu("table-properties", "cell-margins", parsedItems(cellMarginItems()), {
            size: "large",
          }),
        ]),
        group("data", [
          btn("table-repeat-headers", "repeat-header-rows", { size: "large" }),
          btn("list", "convert-to-text", { size: "large" }),
        ]),
      ],
    },
  ];
}
