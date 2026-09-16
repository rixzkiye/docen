import {
  quickStyles,
  resolveFontName,
  type RunStylePropertiesOptions,
  type StylesOptions,
} from "@docen/docx";

import {
  resolveLang,
  t,
  type RibbonButton,
  type RibbonColorPicker,
  type RibbonCombobox,
  type RibbonControlOrLayout,
  type RibbonGroup,
  type RibbonInput,
  type RibbonLayout,
  type RibbonMenu,
  type RibbonMenuItem,
  type RibbonSeparator,
  type RibbonSplit,
  type RibbonTab,
} from "../../ui";

/** Shared ribbon machinery — i18n key shortcuts, schema builders, and the
 *  option sets / groups more than one tab uses. The tab modules below compose
 *  these; `ribbon.ts` re-exports the public surface. */

// Ribbon i18n keys — the schema stores keys (not translated strings); the
// render pass (renderRibbonFromSchema) resolves them via t(). t() returns the
// key itself when no translation is registered, so an addin may also pass a
// plain display string as a label (escape hatch).
export const tab = (id: string): string => `ribbon.tab.${id}`;

export const grp = (id: string): string => `ribbon.group.${id}`;

export const cmd = (event: string): string => `ribbon.cmd.${event}`;

export const opt = (value: string): string => `ribbon.opt.${value}`;

/** Minimal built-in set shown when a document carries no styles.xml (e.g. a
 *  blank editor) so the Styles gallery is never empty — the cards still carry
 *  the document defaults' formatting so they match the rendered text. */
export const FALLBACK_STYLE_ITEMS = (css?: string): RibbonMenuItem[] => {
  const entry = (key: string, value: string): RibbonMenuItem => ({
    text: t(opt(key)),
    preview: css ? { text: t(opt(key)), css } : { text: t(opt(key)) },
    value,
  });
  return [
    entry("normal", "Normal"),
    entry("heading-1", "Heading1"),
    entry("heading-2", "Heading2"),
    entry("heading-3", "Heading3"),
    entry("title", "Title"),
  ];
};

/** Inline CSS previewing a style's own character formatting: the Styles
 *  gallery renders each card's label in the style's font/size/color (Word's
 *  Quick Styles thumbnails) and the Modify Style dialog's preview area renders
 *  its sample text the same way. Sizes clamp so a Title still fits the card
 *  row. Exported for the host, which passes it into the dialog state. */
export const stylePreviewCss = (run?: RunStylePropertiesOptions): string | undefined => {
  if (!run) return undefined;
  const parts: string[] = [];
  const family = resolveFontName(run.font);
  if (family) parts.push(`font-family:${family}`);
  if (typeof run.size === "number" && run.size > 0)
    parts.push(`font-size:${Math.min(Math.max(run.size * 0.85, 8), 15)}pt`);
  if (run.bold) parts.push("font-weight:700");
  if (run.italic) parts.push("font-style:italic");
  // HexColorOrAuto: the hex string, "auto" (default ink), or a theme-color
  // reference whose `val` is the resolved literal — preview it either way.
  const color = typeof run.color === "string" ? run.color : run.color?.val;
  if (typeof color === "string" && color !== "auto") parts.push(`color:#${color}`);
  return parts.length > 0 ? parts.join(";") : undefined;
};

/** Build the Styles gallery (Word's Quick Styles strip): quickFormat paragraph
 *  styles by uiPriority, each card showing the style's own name in its own
 *  character formatting. The value is the pStyle id, which round-trips via the
 *  paragraph `style` attr. Exported so the host can rebuild the items when the
 *  document's styles model changes (the ribbon template bakes one snapshot). */
export const styleGalleryItems = (styles?: StylesOptions | null): RibbonMenuItem[] => {
  const entries = quickStyles(styles);
  if (entries.length === 0)
    return FALLBACK_STYLE_ITEMS(stylePreviewCss(styles?.default?.document?.run ?? undefined));
  return entries.map((e) => ({
    text: e.name,
    preview: { text: e.name, css: stylePreviewCss(e.run) },
    value: e.id,
  }));
};

export const spacingItems = (): string =>
  JSON.stringify([
    { text: "1.0", value: "1.0" },
    { text: "1.15", value: "1.15" },
    { text: "1.5", value: "1.5" },
    { text: "2.0", value: "2.0" },
    { text: opt("add-before"), value: "add-before" },
    { text: opt("add-after"), value: "add-after" },
  ]);

export const borderItems = (): string =>
  JSON.stringify([
    { text: opt("no-border"), value: "none" },
    { text: opt("bottom"), value: "bottom" },
    { text: opt("top"), value: "top" },
    { text: opt("left"), value: "left" },
    { text: opt("right"), value: "right" },
    { text: opt("all"), value: "all" },
    { text: opt("outside"), value: "outside" },
    { text: opt("borders-shading"), value: "borders-shading" },
  ]);

// Word's Group menu. The host event is the group command (the primary item
// inherits it); Ungroup overrides with its own — it needs no multi-selection,
// just a selected floating group.
export const groupItems = (): string =>
  JSON.stringify([
    { text: cmd("group"), value: "group" },
    { text: opt("ungroup"), value: "ungroup", event: "drawing-ungroup" },
  ]);

export const rotateItems = (): string =>
  JSON.stringify([
    { text: opt("rotate-right"), value: "right" },
    { text: opt("rotate-left"), value: "left" },
    { text: opt("flip-vertical"), value: "flip-v" },
    { text: opt("flip-horizontal"), value: "flip-h" },
  ]);

// The Arrange group's floating-drawing menus: Wrap Text (Word's menu — In Line
// with Text first, then the five text-flow styles) and the Position gallery's
// nine-cell grid (option texts reuse the 9-grid keys).
export const wrapItems = (): string =>
  JSON.stringify([
    { text: opt("wrap-inline"), value: "inline" },
    { text: opt("wrap-front"), value: "front" },
    { text: opt("wrap-behind"), value: "behind" },
    { text: opt("wrap-square"), value: "square" },
    { text: opt("wrap-tight"), value: "tight" },
    { text: opt("wrap-through"), value: "through" },
    { text: opt("wrap-top-bottom"), value: "top-bottom" },
  ]);

export const positionItems = (): string =>
  JSON.stringify(
    ["tl", "tc", "tr", "ml", "mc", "mr", "bl", "bc", "br"].map((cell) => ({
      text: opt(`cell-align-${cell}`),
      value: cell,
    })),
  );

// Word's Align menu: both axes' margin alignment plus the two distributes.
// The distributes act on the multi-selection (the primary + Shift+Click set);
// the command declines below two members.
export const alignObjectsItems = (): string =>
  JSON.stringify([
    { text: cmd("align-left"), value: "left" },
    { text: cmd("align-center"), value: "center" },
    { text: cmd("align-right"), value: "right" },
    { text: opt("align-top"), value: "top" },
    { text: opt("align-middle"), value: "middle" },
    { text: opt("align-bottom"), value: "bottom" },
    { text: opt("distribute-h"), value: "h", event: "drawing-distribute" },
    { text: opt("distribute-v"), value: "v", event: "drawing-distribute" },
  ]);

/** Default active tab id. */
export const DEFAULT_RIBBON_TAB = "home";

/** The full ordered set of ribbon tab ids (Home → View). */
export const RIBBON_TAB_IDS = [
  "home",
  "insert",
  "draw",
  "design",
  "layout",
  "references",
  "mailings",
  "review",
  "view",
  "developer",
] as const;

export type RibbonTabId = (typeof RIBBON_TAB_IDS)[number];

/** Parse a legacy items JSON string (the form the *Panel helpers emit) into
 *  RibbonMenuItem data for the data-driven ribbon. */
// --- Data-driven ribbon (RibbonTab tree) -------------------------------------
// The 9 tabs expressed as data (i18n keys, not translated strings);
// renderRibbonFromSchema consumes this tree to build the ribbon DOM and
// resolves the keys via t() — so re-rendering on a locale change re-localizes.

export const parsedItems = (json: string): RibbonMenuItem[] => JSON.parse(json) as RibbonMenuItem[];

export const col = (controls: readonly RibbonControlOrLayout[]): RibbonLayout => ({
  type: "layout",
  layout: "column",
  controls,
});

export const row = (controls: readonly RibbonControlOrLayout[]): RibbonLayout => ({
  type: "layout",
  layout: "row",
  controls,
});

export const grid = (
  controls: readonly RibbonControlOrLayout[],
  columns?: number,
): RibbonLayout => ({
  type: "layout",
  layout: "grid",
  ...(columns ? { columns } : {}),
  controls,
});

export const sep = (): RibbonSeparator => ({ type: "separator" });

export const btn = (
  icon: string,
  event: string,
  o: { size?: "large"; iconOnly?: boolean; toggle?: boolean } = {},
): RibbonButton => ({
  type: "button",
  icon,
  event,
  label: cmd(event),
  ...(o.size ? { size: o.size } : {}),
  ...(o.iconOnly ? { iconOnly: true } : {}),
  ...(o.toggle ? { toggle: true } : {}),
});

export const split = (
  icon: string,
  event: string,
  items: RibbonMenuItem[],
  o: { size?: "large"; iconOnly?: boolean; label?: string; toggle?: boolean } = {},
): RibbonSplit => ({
  type: "split",
  icon,
  event,
  label: o.label ?? cmd(event),
  items,
  ...(o.size ? { size: o.size } : {}),
  ...(o.iconOnly ? { iconOnly: true } : {}),
  ...(o.toggle ? { toggle: true } : {}),
});

export const menu = (
  icon: string,
  event: string,
  items: RibbonMenuItem[],
  o: { label?: string; size?: "large"; iconOnly?: boolean } = {},
): RibbonMenu => ({
  type: "menu",
  icon,
  event,
  label: o.label ?? cmd(event),
  items,
  ...(o.size ? { size: o.size } : {}),
  ...(o.iconOnly ? { iconOnly: true } : {}),
});

export const combo = (
  event: string,
  value: string,
  items: RibbonMenuItem[],
  o: { source?: "local-fonts"; comboboxSize?: "short" } = {},
): RibbonCombobox => ({
  type: "combobox",
  event,
  value,
  items,
  ...(o.source ? { source: o.source } : {}),
  ...(o.comboboxSize ? { comboboxSize: o.comboboxSize } : {}),
});

/** A plain typeable measure box (no drop-down) — Word's Size groups. */
export const input = (event: string, value: string): RibbonInput => ({
  type: "input",
  event,
  value,
});

export const picker = (
  icon: string,
  event: string,
  defaultColor: string,
  o: {
    palette?: "theme" | "highlight";
    panel?: "color" | "outline";
    /** Show the label beside the icon (Word's border/fill buttons in the
     *  drawing tabs carry text; the font/shading pickers stay icon-only). */
    withLabel?: boolean;
    size?: "large";
  } = {},
): RibbonColorPicker => ({
  type: "color-picker",
  icon,
  event,
  label: cmd(event),
  defaultColor,
  iconOnly: !o.withLabel,
  ...(o.size ? { size: o.size } : {}),
  ...(o.palette ? { palette: o.palette } : {}),
  ...(o.panel ? { panel: o.panel } : {}),
});

export const group = (
  id: string,
  controls: readonly RibbonControlOrLayout[],
  launcher?: string,
): RibbonGroup => ({
  id,
  label: grp(id),
  controls,
  ...(launcher ? { launcher } : {}),
});

export const tabNode = (id: RibbonTabId, groups: RibbonGroup[]): RibbonTab => ({
  id,
  label: tab(id),
  groups,
});

/** The Accessibility group — Alt Text through the size-and-position dialog,
 *  shared by every drawing contextual tab. */
export const accessibilityGroup = (): RibbonGroup =>
  group("accessibility", [
    {
      type: "button",
      icon: "alt-text",
      label: "ribbon.cmd.alt-text",
      event: "drawing-properties",
      size: "large",
    },
  ]);

/** The Arrange group — the Layout tab's floating-drawing commands, repeated
 *  verbatim on each drawing contextual tab (Word repeats them across the
 *  Picture/Shape Format tabs). Position/Wrap apply to floating drawings
 *  only, matching Word's inline grey-out at the command layer. */
export const arrangeGroup = (): RibbonGroup =>
  group("arrange", [
    // Word lays these four as a 2×2 grid with shared column tracks — the
    // second column starts at one x even though the top row's dropdowns
    // measure wider than the bottom row's buttons.
    grid(
      [
        menu("orientation", "position", parsedItems(positionItems())),
        menu("wrap", "wrap", parsedItems(wrapItems())),
        btn("orientation", "bring-forward"),
        btn("orientation", "send-backward"),
      ],
      2,
    ),
    menu("align-left", "align-objects", parsedItems(alignObjectsItems()), { size: "large" }),
    menu("group-objects", "drawing-group", parsedItems(groupItems()), { size: "large" }),
    menu("rotate", "rotate", parsedItems(rotateItems()), { size: "large" }),
    btn("selection-pane", "selection-pane", { size: "large" }),
  ]);

/** The Crop split's drop-down — Word's Crop menu flattened: the crop action,
 *  Word's eight Aspect Ratio presets (Word nests them in a Portrait/Landscape
 *  submenu the flat menu cannot express, so they list in Word's order), and
 *  the reset. The ratio texts assemble a localized lead with the plain
 *  "W:H" (the picture-correction pattern). */
export const cropMenuItems = (): RibbonMenuItem[] => {
  const zh = resolveLang().toLowerCase().startsWith("zh");
  const portrait = zh ? "纵向" : "Portrait";
  const landscape = zh ? "横向" : "Landscape";
  const ratio = (lead: string, r: string): RibbonMenuItem => ({
    text: `${lead} ${r}`,
    value: r,
    event: "drawing-crop-aspect",
  });
  return [
    { text: "ribbon.cmd.crop", event: "drawing-crop" },
    { text: "-" },
    ratio(portrait, "2:3"),
    ratio(portrait, "3:4"),
    ratio(portrait, "3:5"),
    ratio(portrait, "4:5"),
    ratio(landscape, "1:1"),
    ratio(landscape, "4:3"),
    ratio(landscape, "5:3"),
    ratio(landscape, "16:9"),
    { text: "-" },
    { text: "context.crop-reset", event: "drawing-crop-reset" },
  ];
};

/** The Reset Picture split's drop-down — Word's two variants: the plain
 *  discard (adjustments/border/effects/crop) and the one that also restores
 *  the source's natural size. */
export const resetPictureItems = (): RibbonMenuItem[] => [
  { text: "ribbon.cmd.reset-picture", event: "reset-picture" },
  { text: "ribbon.cmd.reset-picture-size", event: "reset-picture-size" },
];

/** The Size group — the numeric Height/Width boxes every drawing tab shares;
 *  pictures add the crop split (a shape has no crop). The host mirrors the
 *  selected drawing's extent into the boxes per transaction (#syncDrawingSize);
 *  a typed measure commits the dimension (a bare number reads in the locale's
 *  unit system — #onCommand qualifies it before dispatch). */
export const sizeGroup = (
  id: "picture-size" | "shape-size" | "chart-size",
  withCrop: boolean,
): RibbonGroup =>
  group(id, [
    col([input("drawing-height", ""), input("drawing-width", "")]),
    ...(withCrop
      ? [
          split("crop", "drawing-crop", cropMenuItems(), {
            size: "large",
            label: "ribbon.cmd.crop",
          }),
        ]
      : []),
  ]);
