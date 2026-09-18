// Style cascade — resolving a document's styles.xml model (StylesOptions):
// the id index, the default paragraph style, and the basedOn-chain merge.
// Rendering-neutral: the layout projection, the CSS route, and the editor's
// caret/gallery resolvers all share these primitives, so one cascade runs
// everywhere.

import type {
  ConditionalTableStyleOptions,
  ParagraphStylePropertiesOptions,
  RunStylePropertiesOptions,
  StylesOptions,
  TableBordersOptions,
  TableOptions,
  TablePropertiesOptions,
  TableRowPropertiesOptions,
  TableStyleOptions,
  TableStyleOverrideType,
} from "@office-open/docx";

type TableCellPropertiesOptions = NonNullable<ConditionalTableStyleOptions["cell"]>;

/** Table-level cell margins (w:tblCellMar) — TableCellMarginOptions is not
 *  exported, derive it from the field that carries it. */
type TableCellMargins = NonNullable<TableOptions["margins"]>;

/** A named style entry as office-open models it: BaseParagraphStyleOptions or
 *  BaseCharacterStyleOptions (both extend the internal StyleOptions, carrying
 *  name/uiPriority/quickFormat). Derived from the public StylesOptions — not
 *  imported — because StyleOptions is not a public export of @office-open/docx. */
export type StyleEntry =
  | NonNullable<StylesOptions["paragraphStyles"]>[number]
  | NonNullable<StylesOptions["characterStyles"]>[number];

/** The pStyle val (the style's OOXML id) for a built-in named style nested
 *  under DefaultStylesOptions: the key with its first letter upper-cased
 *  ("heading1" → "Heading1", "title" → "Title", "listParagraph" →
 *  "ListParagraph"). This matches office-open's HeadingLevel literals / pStyle
 *  ids, so we derive the id from the key instead of hard-coding a name table. */
export function pStyleIdFromKey(key: string): string {
  return key.charAt(0).toUpperCase() + key.slice(1);
}

/** The styleId of the document's default paragraph style (`w:default="1"`
 *  type="paragraph") — the implicit style applied to every paragraph WITHOUT an
 *  explicit pStyle. OOXML renders a pStyle-less paragraph as this style (usually
 *  "Normal"). Searched in `paragraphStyles` and the built-in named styles nested
 *  under `default` (key → pStyle id). null when the document declares none.
 *  WeakMap-cached per styles object like the indexes — the projection asks for
 *  the default style of every pStyle-less paragraph, per transaction. */
const defaultParagraphStyleCache = new WeakMap<StylesOptions, string | null>();

export function defaultParagraphStyleId(styles: StylesOptions | null | undefined): string | null {
  if (!styles) return null;
  const cached = defaultParagraphStyleCache.get(styles);
  if (cached !== undefined) return cached;
  let result: string | null = null;
  for (const ps of styles.paragraphStyles ?? []) {
    // `default` (w:default="1") is on the runtime shape but not the public
    // StyleOptions type — read it loosely.
    if ((ps as { default?: boolean }).default) {
      result = ps.id;
      break;
    }
  }
  if (result === null) {
    const defaults = styles.default as unknown as Record<string, StyleEntry | undefined>;
    for (const [key, style] of Object.entries(defaults ?? {})) {
      if (key === "document" || !style) continue;
      if ((style as { default?: boolean }).default) {
        result = pStyleIdFromKey(key);
        break;
      }
    }
  }
  defaultParagraphStyleCache.set(styles, result);
  return result;
}

/** Build an id → style-entry index over every paragraph style: the explicit
 *  `paragraphStyles` plus the built-in named styles nested under `default`
 *  (key → pStyle id via pStyleIdFromKey). `document` is docDefaults, not a
 *  named style, so it is excluded. A built-in that also appears in
 *  paragraphStyles is deduped by id — the built-in wins, being set second. */
// Cache the style index by the styles object reference. A document's styles
// model is stable for its lifetime (set on load, unchanged across edits), yet
// indexParagraphStyles is called per-paragraph (detectHeadingLevel during
// resolve), per-transaction (effectiveRunProps at the caret), and per-render
// (layout projection). The WeakMap memo turns all of those into O(1) lookups after
// the first build and frees the entry when the styles object is GC'd. Callers
// treat the result as read-only (mergeStyleChain only .get()s).
const styleIndexCache = new WeakMap<StylesOptions, Map<string, StyleEntry>>();

export function indexParagraphStyles(styles: StylesOptions): Map<string, StyleEntry> {
  const cached = styleIndexCache.get(styles);
  if (cached) return cached;
  const byId = new Map<string, StyleEntry>();
  for (const ps of styles.paragraphStyles ?? []) byId.set(ps.id, ps);
  const defaults = styles.default as unknown as Record<string, StyleEntry | undefined>;
  for (const [key, style] of Object.entries(defaults ?? {})) {
    if (key === "document" || !style) continue;
    byId.set(pStyleIdFromKey(key), style);
  }
  styleIndexCache.set(styles, byId);
  return byId;
}

/** The `default` keys that carry character styles (DefaultStylesOptions types
 *  them CharacterStyleOptions; every other key is a paragraph style). */
const CHARACTER_DEFAULT_KEYS = [
  "hyperlink",
  "footnoteReference",
  "footnoteTextChar",
  "endnoteReference",
  "endnoteTextChar",
] as const;

/** Build an id → style-entry index over every character style: the explicit
 *  `characterStyles` plus the built-in character styles nested under `default`
 *  (key → style id via pStyleIdFromKey, e.g. "hyperlink" → "Hyperlink"). A
 *  built-in that also appears in characterStyles is deduped by id. WeakMap-
 *  cached per styles object, like indexParagraphStyles — the projection
 *  resolves a run's w:rStyle per run, only a handful of distinct ids exist. */
const characterStyleIndexCache = new WeakMap<StylesOptions, Map<string, StyleEntry>>();

export function indexCharacterStyles(styles: StylesOptions | undefined): Map<string, StyleEntry> {
  if (!styles) return new Map();
  const cached = characterStyleIndexCache.get(styles);
  if (cached) return cached;
  const byId = new Map<string, StyleEntry>();
  for (const cs of styles.characterStyles ?? []) byId.set(cs.id, cs);
  const defaults = styles.default as unknown as Record<string, StyleEntry | undefined>;
  for (const key of CHARACTER_DEFAULT_KEYS) {
    const style = defaults?.[key];
    if (style) byId.set(pStyleIdFromKey(key), style);
  }
  characterStyleIndexCache.set(styles, byId);
  return byId;
}

/** id → table-style index, WeakMap-cached per styles object like the
 *  paragraph/character indexes — the projection resolves a table's w:tblStyle
 *  per table, per transaction. */
const tableStyleIndexCache = new WeakMap<StylesOptions, Map<string, TableStyleOptions>>();

export function indexTableStyles(
  styles: StylesOptions | undefined,
): Map<string, TableStyleOptions> {
  if (!styles) return new Map();
  const cached = tableStyleIndexCache.get(styles);
  if (cached) return cached;
  const byId = new Map<string, TableStyleOptions>();
  for (const ts of styles.tableStyles ?? []) byId.set(ts.id, ts);
  tableStyleIndexCache.set(styles, byId);
  return byId;
}

/** Whether `v` is a plain object — an OOXML property group (spacing/indent/
 *  border/shading/font) that merges key by key — as opposed to an array
 *  (tabStops) or scalar, which replace. */
function isPlainObject(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

/** Deep-merge `source` into `target` (mutates target) per the OOXML `basedOn`
 *  model: nested property groups merge key by key (a child's spacing.before
 *  merges with, not replaces, the parent's spacing.line); arrays and scalars
 *  replace. Nullish source values are skipped so an unset child key doesn't
 *  clobber an inherited value. */
export function deepMergeInto(
  target: Record<string, unknown>,
  source: Record<string, unknown>,
): Record<string, unknown> {
  for (const [key, srcVal] of Object.entries(source)) {
    if (srcVal === null || srcVal === undefined) continue;
    const tgtVal = target[key];
    target[key] =
      isPlainObject(srcVal) && isPlainObject(tgtVal)
        ? deepMergeInto({ ...tgtVal }, srcVal)
        : isPlainObject(srcVal)
          ? { ...srcVal }
          : srcVal;
  }
  return target;
}

/** Merge a paragraph style's run/paragraph properties with its `basedOn` chain
 *  (root first, child overrides per-property) — the OOXML inheritance model.
 *  Nested property groups (spacing/indent/border/font) merge key by key; arrays
 *  and scalars replace. Shared by every style consumer (the layout projection,
 *  the CSS route, the caret resolver) so all of them resolve identical values. */
// Memoize mergeStyleChain per (byId, styleId). byId is itself memoized by
// indexParagraphStyles (one Map per StylesOptions object), so this WeakMap
// frees the cache when the styles object is GC'd. The same styleId is resolved
// for every paragraph/run that carries it (thousands of calls on a large doc,
// only dozens of distinct ids) — the chain walk + deepMergeInto is repeated
// work. Callers treat the result as read-only (consumers only read fields;
// resolveNode spreads `{...paragraph}` before merging its own attrs), so a
// shared cached value is safe.
const styleChainCache = new WeakMap<
  Map<string, StyleEntry>,
  Map<string, { run: Record<string, unknown>; paragraph: Record<string, unknown> }>
>();

export function mergeStyleChain(
  byId: Map<string, StyleEntry>,
  styleId: string | null | undefined,
): { run: Record<string, unknown>; paragraph: Record<string, unknown> } {
  if (!styleId) return { run: {}, paragraph: {} };
  const perId = styleChainCache.get(byId);
  if (perId) {
    const cached = perId.get(styleId);
    if (cached) return cached;
  }
  const chain: StyleEntry[] = [];
  const visited = new Set<string>();
  let curId: string | undefined = styleId;
  while (curId && !visited.has(curId)) {
    visited.add(curId);
    const style = byId.get(curId);
    if (!style) break;
    chain.unshift(style); // root first, so children override
    curId = style.basedOn ?? undefined;
  }
  const run: Record<string, unknown> = {};
  const paragraph: Record<string, unknown> = {};
  for (const style of chain) {
    // StyleEntry is a paragraph|character union; paragraph props live only on
    // the paragraph side, so access via a loose record.
    const s = style as unknown as Record<string, unknown>;
    if (s.run) deepMergeInto(run, s.run as Record<string, unknown>);
    if (s.paragraph) deepMergeInto(paragraph, s.paragraph as Record<string, unknown>);
  }
  const result = { run, paragraph };
  const bucket =
    perId ??
    new Map<string, { run: Record<string, unknown>; paragraph: Record<string, unknown> }>();
  if (!perId) styleChainCache.set(byId, bucket);
  bucket.set(styleId, result);
  return result;
}

/** Resolve a table style's effective table-level properties (tblBorders,
 *  tblCellMar) by walking its basedOn chain (root first, child overrides) —
 *  the table-style counterpart of mergeStyleChain. office-open's parseDocument
 *  does NOT merge the referenced <w:tblStyle> into table.borders/cellMargin
 *  (they reflect only the table's own <w:tblPr>), so a "Table Grid" table whose
 *  borders live in the style needs this to render its grid. Returns empty when
 *  the style is absent or unknown. */
export function mergeTableStyleProps(
  tableStyles: StylesOptions["tableStyles"],
  styleId: string | null | undefined,
): { borders?: TableBordersOptions; margins?: TableCellMargins } {
  if (!styleId || !tableStyles) return {};
  const byId = new Map(tableStyles.map((t) => [t.id, t]));
  const chain: NonNullable<StylesOptions["tableStyles"]> = [];
  const visited = new Set<string>();
  let cur: string | undefined = styleId ?? undefined;
  while (cur && !visited.has(cur)) {
    visited.add(cur);
    const s = byId.get(cur);
    if (!s) break;
    chain.unshift(s); // root first → children override below
    cur = s.basedOn;
  }
  // Per-key inheritance: a child style's tblBorders/tblCellMar overrides only
  // the edges it declares — Word keeps the basedOn chain's remaining edges,
  // so a whole-object swap would lose the parent's other sides.
  let borders: TableBordersOptions | undefined;
  let margins: TableCellMargins | undefined;
  for (const s of chain) {
    const t = s.table;
    if (!t) continue;
    if (t.borders) borders = { ...borders, ...t.borders };
    if (t.margins) margins = { ...margins, ...t.margins };
  }
  const out: { borders?: TableBordersOptions; margins?: TableCellMargins } = {};
  if (borders) out.borders = borders;
  if (margins) out.margins = margins;
  return out;
}

// ── OOXML Table Style & tblLook Conditional Formatting Cascade ──

export interface ResolvedTableStyle {
  table?: TablePropertiesOptions;
  row?: TableRowPropertiesOptions;
  cell?: TableCellPropertiesOptions;
  paragraph?: ParagraphStylePropertiesOptions;
  run?: RunStylePropertiesOptions;
  conditionalFormats: Map<TableStyleOverrideType, ConditionalTableStyleOptions>;
}

export interface ResolvedTableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  bandRow: boolean;
  bandCol: boolean;
}

/** Built-in table style definitions fallback when not present in styles.xml. */
const BUILTIN_TABLE_STYLES: Record<string, TableStyleOptions> = {
  "table-grid": {
    id: "table-grid",
    table: {
      borders: {
        top: { style: "single", size: 4, color: "auto" },
        bottom: { style: "single", size: 4, color: "auto" },
        left: { style: "single", size: 4, color: "auto" },
        right: { style: "single", size: 4, color: "auto" },
        insideHorizontal: { style: "single", size: 4, color: "auto" },
        insideVertical: { style: "single", size: 4, color: "auto" },
      },
    },
  },
  TableGrid: {
    id: "TableGrid",
    table: {
      borders: {
        top: { style: "single", size: 4, color: "auto" },
        bottom: { style: "single", size: 4, color: "auto" },
        left: { style: "single", size: 4, color: "auto" },
        right: { style: "single", size: 4, color: "auto" },
        insideHorizontal: { style: "single", size: 4, color: "auto" },
        insideVertical: { style: "single", size: 4, color: "auto" },
      },
    },
  },
  "no-style-no-grid": {
    id: "no-style-no-grid",
    table: {
      borders: {
        top: { style: "none", size: 0, color: "auto" },
        bottom: { style: "none", size: 0, color: "auto" },
        left: { style: "none", size: 0, color: "auto" },
        right: { style: "none", size: 0, color: "auto" },
        insideHorizontal: { style: "none", size: 0, color: "auto" },
        insideVertical: { style: "none", size: 0, color: "auto" },
      },
    },
  },
  "light-shading": {
    id: "light-shading",
    table: {
      borders: {
        top: { style: "single", size: 4, color: "auto" },
        bottom: { style: "single", size: 4, color: "auto" },
        insideHorizontal: { style: "single", size: 4, color: "auto" },
      },
    },
    conditionalFormats: [
      {
        type: "band1Horz",
        cell: { shading: { fill: "D9E2F3", type: "clear" } },
      },
    ],
  },
  "light-list": {
    id: "light-list",
    table: {
      borders: {
        top: { style: "single", size: 4, color: "auto" },
        bottom: { style: "single", size: 4, color: "auto" },
        insideHorizontal: { style: "single", size: 4, color: "auto" },
      },
    },
    conditionalFormats: [
      {
        type: "firstRow",
        cell: { shading: { fill: "8EAADB", type: "clear" } },
        run: { bold: true },
      },
    ],
  },
  "light-grid": {
    id: "light-grid",
    table: {
      borders: {
        top: { style: "single", size: 4, color: "auto" },
        bottom: { style: "single", size: 4, color: "auto" },
        left: { style: "single", size: 4, color: "auto" },
        right: { style: "single", size: 4, color: "auto" },
        insideHorizontal: { style: "single", size: 4, color: "auto" },
        insideVertical: { style: "single", size: 4, color: "auto" },
      },
    },
    conditionalFormats: [
      {
        type: "firstRow",
        cell: { shading: { fill: "D9E2F3", type: "clear" } },
        run: { bold: true },
      },
    ],
  },
  "grid-table": {
    id: "grid-table",
    table: {
      borders: {
        top: { style: "single", size: 8, color: "auto" },
        bottom: { style: "single", size: 8, color: "auto" },
        left: { style: "single", size: 8, color: "auto" },
        right: { style: "single", size: 8, color: "auto" },
        insideHorizontal: { style: "single", size: 4, color: "auto" },
        insideVertical: { style: "single", size: 4, color: "auto" },
      },
    },
    conditionalFormats: [
      {
        type: "firstRow",
        cell: { shading: { fill: "4472C4", type: "clear" } },
        run: { bold: true, color: "FFFFFF" },
      },
      {
        type: "band1Horz",
        cell: { shading: { fill: "D9E2F3", type: "clear" } },
      },
    ],
  },
};

/** Parse and resolve w:tblLook from raw attributes (hex val mask or explicit flags).
 *  Word defaults: firstRow: true, lastRow: false, bandRow: true, firstCol: false, lastCol: false, bandCol: false. */
export function resolveTableLook(look: unknown): ResolvedTableLook {
  let firstRow = true;
  let lastRow = false;
  let firstCol = false;
  let lastCol = false;
  let bandRow = true;
  let bandCol = false;

  if (isPlainObject(look)) {
    const l = look as Record<string, unknown>;
    if (l.val != null) {
      const mask =
        typeof l.val === "number"
          ? l.val
          : typeof l.val === "string"
            ? parseInt(l.val, 16)
            : Number.NaN;
      if (!Number.isNaN(mask)) {
        // Bit 0x0020: apply first row conditional formatting
        firstRow = (mask & 0x0020) !== 0;
        // Bit 0x0040: apply last row conditional formatting
        lastRow = (mask & 0x0040) !== 0;
        // Bit 0x0080: apply first column conditional formatting
        firstCol = (mask & 0x0080) !== 0;
        // Bit 0x0100: apply last column conditional formatting
        lastCol = (mask & 0x0100) !== 0;
        // Bit 0x0200: do not apply horizontal banding (noHBand)
        bandRow = (mask & 0x0200) === 0;
        // Bit 0x0400: do not apply vertical banding (noVBand)
        bandCol = (mask & 0x0400) === 0;
      }
    }
    // Explicit booleans take precedence over val mask
    if (typeof l.firstRow === "boolean") firstRow = l.firstRow;
    if (typeof l.lastRow === "boolean") lastRow = l.lastRow;
    if (typeof l.firstCol === "boolean") firstCol = l.firstCol;
    if (typeof l.lastCol === "boolean") lastCol = l.lastCol;
    if (typeof l.bandRow === "boolean") bandRow = l.bandRow;
    if (typeof l.bandCol === "boolean") bandCol = l.bandCol;
  }
  return { firstRow, lastRow, firstCol, lastCol, bandRow, bandCol };
}

const resolvedTableStyleCache = new WeakMap<
  NonNullable<StylesOptions["tableStyles"]>,
  Map<string, ResolvedTableStyle>
>();
const standaloneResolvedTableStyleCache = new Map<string, ResolvedTableStyle>();

/** Resolve a table style and its basedOn chain into merged base props and conditional formats. */
export function resolveTableStyle(
  tableStyles: StylesOptions["tableStyles"],
  styleId: string | null | undefined,
): ResolvedTableStyle | undefined {
  if (!styleId) return undefined;
  if (tableStyles && tableStyles.length > 0) {
    let perArray = resolvedTableStyleCache.get(tableStyles);
    if (!perArray) {
      perArray = new Map();
      resolvedTableStyleCache.set(tableStyles, perArray);
    }
    const cached = perArray.get(styleId);
    if (cached) return cached;
  } else {
    const cached = standaloneResolvedTableStyleCache.get(styleId);
    if (cached) return cached;
  }

  const byId = new Map<string, TableStyleOptions>();
  for (const t of tableStyles ?? []) byId.set(t.id, t);
  for (const [id, t] of Object.entries(BUILTIN_TABLE_STYLES)) {
    if (!byId.has(id)) byId.set(id, t);
  }

  const chain: TableStyleOptions[] = [];
  const visited = new Set<string>();
  let cur: string | undefined = styleId;
  while (cur && !visited.has(cur)) {
    visited.add(cur);
    const s = byId.get(cur);
    if (!s) break;
    chain.unshift(s); // root first
    cur = s.basedOn;
  }
  if (chain.length === 0) return undefined;

  let table: Record<string, unknown> | undefined;
  let row: Record<string, unknown> | undefined;
  let cell: Record<string, unknown> | undefined;
  let paragraph: Record<string, unknown> | undefined;
  let run: Record<string, unknown> | undefined;
  const conditionalFormats = new Map<TableStyleOverrideType, ConditionalTableStyleOptions>();

  for (const s of chain) {
    if (s.table) table = deepMergeInto(table ?? {}, s.table as Record<string, unknown>);
    if (s.row) row = deepMergeInto(row ?? {}, s.row as Record<string, unknown>);
    if (s.cell) cell = deepMergeInto(cell ?? {}, s.cell as Record<string, unknown>);
    if (s.paragraph)
      paragraph = deepMergeInto(paragraph ?? {}, s.paragraph as Record<string, unknown>);
    if (s.run) run = deepMergeInto(run ?? {}, s.run as Record<string, unknown>);

    for (const cf of s.conditionalFormats ?? []) {
      const existing = conditionalFormats.get(cf.type);
      if (existing) {
        const merged: ConditionalTableStyleOptions = { type: cf.type };
        if (existing.table || cf.table) {
          merged.table = deepMergeInto(
            { ...(existing.table as Record<string, unknown>) },
            (cf.table ?? {}) as Record<string, unknown>,
          ) as TablePropertiesOptions;
        }
        if (existing.row || cf.row) {
          merged.row = deepMergeInto(
            { ...(existing.row as Record<string, unknown>) },
            (cf.row ?? {}) as Record<string, unknown>,
          ) as TableRowPropertiesOptions;
        }
        if (existing.cell || cf.cell) {
          merged.cell = deepMergeInto(
            { ...(existing.cell as Record<string, unknown>) },
            (cf.cell ?? {}) as Record<string, unknown>,
          ) as TableCellPropertiesOptions;
        }
        if (existing.paragraph || cf.paragraph) {
          merged.paragraph = deepMergeInto(
            { ...(existing.paragraph as Record<string, unknown>) },
            (cf.paragraph ?? {}) as Record<string, unknown>,
          ) as ParagraphStylePropertiesOptions;
        }
        if (existing.run || cf.run) {
          merged.run = deepMergeInto(
            { ...(existing.run as Record<string, unknown>) },
            (cf.run ?? {}) as Record<string, unknown>,
          ) as RunStylePropertiesOptions;
        }
        conditionalFormats.set(cf.type, merged);
      } else {
        conditionalFormats.set(cf.type, {
          type: cf.type,
          table: cf.table ? { ...cf.table } : undefined,
          row: cf.row ? { ...cf.row } : undefined,
          cell: cf.cell ? { ...cf.cell } : undefined,
          paragraph: cf.paragraph ? { ...cf.paragraph } : undefined,
          run: cf.run ? { ...cf.run } : undefined,
        });
      }
    }
  }

  const result: ResolvedTableStyle = {
    table: table as TablePropertiesOptions | undefined,
    row: row as TableRowPropertiesOptions | undefined,
    cell: cell as TableCellPropertiesOptions | undefined,
    paragraph: paragraph as ParagraphStylePropertiesOptions | undefined,
    run: run as RunStylePropertiesOptions | undefined,
    conditionalFormats,
  };

  if (tableStyles && tableStyles.length > 0) {
    resolvedTableStyleCache.get(tableStyles)?.set(styleId, result);
  } else {
    standaloneResolvedTableStyleCache.set(styleId, result);
  }
  return result;
}

export interface TableCellPosition {
  rowIndex: number;
  colIndex: number;
  rowSpan?: number;
  colSpan?: number;
  totalRows: number;
  totalCols: number;
}

/** ECMA-376 Part 1 §17.7.2 priority order for conditional table formatting.
 *  Earlier entries have lower precedence; later entries override earlier ones. */
export const CONDITIONAL_FORMAT_PRIORITY: readonly TableStyleOverrideType[] = [
  "wholeTable",
  "band1Horz",
  "band2Horz",
  "band1Vert",
  "band2Vert",
  "firstCol",
  "lastCol",
  "firstRow",
  "lastRow",
  "neCell",
  "nwCell",
  "seCell",
  "swCell",
];

/** Determine which conditional formatting types apply to a cell at a given position. */
export function activeConditionalTypes(
  pos: TableCellPosition,
  look: ResolvedTableLook,
  styleRowBandSize?: number,
  styleColBandSize?: number,
): Set<TableStyleOverrideType> {
  const active = new Set<TableStyleOverrideType>();
  active.add("wholeTable");

  const r = pos.rowIndex;
  const col = pos.colIndex;
  const rSpan = pos.rowSpan ?? 1;
  const cSpan = pos.colSpan ?? 1;
  const rEnd = r + rSpan - 1;
  const cEnd = col + cSpan - 1;

  const isFirstRow = r === 0;
  const isLastRow = rEnd === pos.totalRows - 1;
  const isFirstCol = col === 0;
  const isLastCol = cEnd === pos.totalCols - 1;

  // Horizontal banding (alternating row bands)
  if (look.bandRow) {
    const isExempt = (look.firstRow && isFirstRow) || (look.lastRow && isLastRow);
    if (!isExempt) {
      const bandSize = Math.max(1, styleRowBandSize ?? 1);
      const rowOffset = r - (look.firstRow ? 1 : 0);
      if (rowOffset >= 0) {
        const bandIdx = Math.floor(rowOffset / bandSize);
        if (bandIdx % 2 === 0) {
          active.add("band1Horz");
        } else {
          active.add("band2Horz");
        }
      }
    }
  }

  // Vertical banding (alternating column bands)
  if (look.bandCol) {
    const isExempt = (look.firstCol && isFirstCol) || (look.lastCol && isLastCol);
    if (!isExempt) {
      const bandSize = Math.max(1, styleColBandSize ?? 1);
      const colOffset = col - (look.firstCol ? 1 : 0);
      if (colOffset >= 0) {
        const bandIdx = Math.floor(colOffset / bandSize);
        if (bandIdx % 2 === 0) {
          active.add("band1Vert");
        } else {
          active.add("band2Vert");
        }
      }
    }
  }

  if (look.firstCol && isFirstCol) active.add("firstCol");
  if (look.lastCol && isLastCol) active.add("lastCol");
  if (look.firstRow && isFirstRow) active.add("firstRow");
  if (look.lastRow && isLastRow) active.add("lastRow");

  // Four corner cells (active only when both respective flags are enabled)
  if (look.firstRow && look.lastCol && isFirstRow && isLastCol) active.add("neCell");
  if (look.firstRow && look.firstCol && isFirstRow && isFirstCol) active.add("nwCell");
  if (look.lastRow && look.lastCol && isLastRow && isLastCol) active.add("seCell");
  if (look.lastRow && look.firstCol && isLastRow && isFirstCol) active.add("swCell");

  return active;
}

export interface EffectiveTableCellStyle {
  cell?: TableCellPropertiesOptions;
  row?: TableRowPropertiesOptions;
  paragraph?: ParagraphStylePropertiesOptions;
  run?: RunStylePropertiesOptions;
}

/** Resolve effective cell, row, paragraph, and run properties for a cell by cascading
 *  active conditional formats in ECMA-376 priority order. */
export function resolveTableCellStyle(
  style: ResolvedTableStyle,
  pos: TableCellPosition,
  look: ResolvedTableLook,
  styleRowBandSize?: number,
  styleColBandSize?: number,
): EffectiveTableCellStyle {
  const active = activeConditionalTypes(pos, look, styleRowBandSize, styleColBandSize);

  let cell: Record<string, unknown> = style.cell
    ? deepMergeInto({}, style.cell as Record<string, unknown>)
    : {};
  let row: Record<string, unknown> = style.row
    ? deepMergeInto({}, style.row as Record<string, unknown>)
    : {};
  let paragraph: Record<string, unknown> = style.paragraph
    ? deepMergeInto({}, style.paragraph as Record<string, unknown>)
    : {};
  let run: Record<string, unknown> = style.run
    ? deepMergeInto({}, style.run as Record<string, unknown>)
    : {};

  for (const type of CONDITIONAL_FORMAT_PRIORITY) {
    if (!active.has(type)) continue;
    const cf = style.conditionalFormats.get(type);
    if (!cf) continue;

    if (cf.cell) cell = deepMergeInto(cell, cf.cell as Record<string, unknown>);
    if (cf.row) row = deepMergeInto(row, cf.row as Record<string, unknown>);
    if (cf.paragraph) paragraph = deepMergeInto(paragraph, cf.paragraph as Record<string, unknown>);
    if (cf.run) run = deepMergeInto(run, cf.run as Record<string, unknown>);
  }

  return {
    cell: Object.keys(cell).length > 0 ? (cell as TableCellPropertiesOptions) : undefined,
    row: Object.keys(row).length > 0 ? (row as TableRowPropertiesOptions) : undefined,
    paragraph:
      Object.keys(paragraph).length > 0
        ? (paragraph as ParagraphStylePropertiesOptions)
        : undefined,
    run: Object.keys(run).length > 0 ? (run as RunStylePropertiesOptions) : undefined,
  };
}
