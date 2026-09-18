/**
 * Drawing and VML Shape Layout Modeling.
 * Converts between raw VML shape style strings / dictionaries and structured layout properties.
 */

export type ShapeWrapping =
  | "inline"
  | "square"
  | "tight"
  | "through"
  | "topAndBottom"
  | "behind"
  | "inFront";

export interface DrawingShapeLayout {
  width?: number; // In twips (1 pt = 20 twips)
  height?: number; // In twips
  position?: "relative" | "absolute" | "static";
  left?: number; // In twips
  top?: number; // In twips
  marginLeft?: number; // In twips
  marginTop?: number; // In twips
  marginRight?: number; // In twips
  marginBottom?: number; // In twips
  zIndex?: number;
  rotation?: number; // Degrees
  wrapping?: ShapeWrapping;
  horizontalPosition?: string; // e.g. "absolute", "left", "center", "right"
  verticalPosition?: string; // e.g. "absolute", "top", "center", "bottom"
  horizontalRelative?: string; // e.g. "page", "margin", "column"
  verticalRelative?: string; // e.g. "page", "margin", "paragraph", "line"
  rawStyle?: Record<string, string | number>;
}

// ── Unit conversion helpers ──

export function ptToTwip(pt: number): number {
  return Math.round(pt * 20);
}

export function twipToPt(twip: number): number {
  return twip / 20;
}

export function emuToTwip(emu: number): number {
  return Math.round(emu / 635);
}

export function twipToEmu(twip: number): number {
  return Math.round(twip * 635);
}

export function ptToEmu(pt: number): number {
  return Math.round(pt * 12700);
}

export function emuToPt(emu: number): number {
  return emu / 12700;
}

const LENGTH_RE = /^(-?[\d.]+)\s*(pt|pc|in|mm|cm|px|emu|twip|twips)?$/i;

/**
 * Parses length units into twips.
 * If no unit is specified, defaultUnit ("pt" for VML styles) is assumed.
 */
export function parseLengthToTwips(
  val: string | number | undefined | null,
  defaultUnit: "pt" | "twip" = "pt",
): number | undefined {
  if (val == null) return undefined;
  if (typeof val === "number") {
    return defaultUnit === "twip" ? Math.round(val) : ptToTwip(val);
  }
  const s = String(val).trim();
  if (!s) return undefined;
  const match = LENGTH_RE.exec(s);
  if (!match) return undefined;
  const num = Number(match[1]);
  if (Number.isNaN(num)) return undefined;
  const unit = (match[2] ?? defaultUnit).toLowerCase();

  switch (unit) {
    case "pt":
      return ptToTwip(num);
    case "in":
      return Math.round(num * 1440);
    case "cm":
      return Math.round(num * (1440 / 2.54));
    case "mm":
      return Math.round(num * (1440 / 25.4));
    case "pc":
      return Math.round(num * 240);
    case "px":
      return Math.round(num * 15);
    case "emu":
      return emuToTwip(num);
    case "twip":
    case "twips":
      return Math.round(num);
    default:
      return defaultUnit === "twip" ? Math.round(num) : ptToTwip(num);
  }
}

/**
 * Normalizes CSS-like style strings or objects into a key-value dictionary.
 */
export function normalizeVmlShapeStyle(
  style: string | Record<string, unknown> | undefined | null,
): Record<string, string | number> {
  if (!style) return {};
  if (typeof style === "object") {
    const res: Record<string, string | number> = {};
    for (const [k, v] of Object.entries(style)) {
      if (v != null && (typeof v === "string" || typeof v === "number")) {
        res[k] = v;
      }
    }
    return res;
  }

  const res: Record<string, string | number> = {};
  const pairs = style.split(";");
  for (const pair of pairs) {
    const idx = pair.indexOf(":");
    if (idx === -1) continue;
    const key = pair.slice(0, idx).trim().toLowerCase();
    const val = pair.slice(idx + 1).trim();
    if (!key || !val) continue;
    const num = Number(val);
    res[key] = !Number.isNaN(num) && /^-?\d+(\.\d+)?$/.test(val) ? num : val;
  }
  return res;
}

function parseWrapping(val: string | undefined): ShapeWrapping | undefined {
  if (!val) return undefined;
  const lower = val.toLowerCase().replace(/[-_]/g, "");
  switch (lower) {
    case "inline":
    case "none":
      return "inline";
    case "square":
      return "square";
    case "tight":
      return "tight";
    case "through":
      return "through";
    case "topandbottom":
    case "topbottom":
      return "topAndBottom";
    case "behind":
      return "behind";
    case "infront":
      return "inFront";
    default:
      return undefined;
  }
}

/**
 * Parses raw VML style (string or object) into structured DrawingShapeLayout.
 */
export function parseVmlShapeLayout(
  raw: string | Record<string, unknown> | undefined | null,
): DrawingShapeLayout | undefined {
  if (!raw) return undefined;
  const style = normalizeVmlShapeStyle(raw);
  if (Object.keys(style).length === 0) return undefined;

  const pick = (...keys: string[]): string | number | undefined => {
    for (const k of keys) {
      if (k in style) return style[k];
      const kebab = k.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`);
      if (kebab in style) return style[kebab];
    }
    return undefined;
  };

  const width = parseLengthToTwips(pick("width"));
  const height = parseLengthToTwips(pick("height"));
  const left = parseLengthToTwips(pick("left"));
  const top = parseLengthToTwips(pick("top"));
  const marginLeft = parseLengthToTwips(pick("margin-left", "marginLeft"));
  const marginTop = parseLengthToTwips(pick("margin-top", "marginTop"));
  const marginRight = parseLengthToTwips(pick("margin-right", "marginRight"));
  const marginBottom = parseLengthToTwips(pick("margin-bottom", "marginBottom"));

  const rawPos = pick("position");
  const position =
    rawPos === "absolute" || rawPos === "relative" || rawPos === "static"
      ? (rawPos as "relative" | "absolute" | "static")
      : undefined;

  const zVal = pick("z-index", "zIndex");
  const zIndex = zVal != null && !Number.isNaN(Number(zVal)) ? Number(zVal) : undefined;

  const rVal = pick("rotation");
  const rotation = rVal != null && !Number.isNaN(Number(rVal)) ? Number(rVal) : undefined;

  const wrapVal = pick("mso-wrap-style", "msoWrapStyle", "wrap", "wrapping");
  const wrapping = parseWrapping(wrapVal != null ? String(wrapVal) : undefined);

  const horizontalPosition = pick(
    "mso-position-horizontal",
    "msoPositionHorizontal",
    "horizontalPosition",
  );
  const verticalPosition = pick("mso-position-vertical", "msoPositionVertical", "verticalPosition");
  const horizontalRelative = pick(
    "mso-position-horizontal-relative",
    "msoPositionHorizontalRelative",
    "horizontalRelative",
  );
  const verticalRelative = pick(
    "mso-position-vertical-relative",
    "msoPositionVerticalRelative",
    "verticalRelative",
  );

  return {
    width,
    height,
    position,
    left,
    top,
    marginLeft,
    marginTop,
    marginRight,
    marginBottom,
    zIndex,
    rotation,
    wrapping,
    horizontalPosition: horizontalPosition != null ? String(horizontalPosition) : undefined,
    verticalPosition: verticalPosition != null ? String(verticalPosition) : undefined,
    horizontalRelative: horizontalRelative != null ? String(horizontalRelative) : undefined,
    verticalRelative: verticalRelative != null ? String(verticalRelative) : undefined,
    rawStyle: style,
  };
}

/**
 * Serializes DrawingShapeLayout into a VML style string.
 */
export function stringifyVmlShapeLayout(layout: DrawingShapeLayout): string {
  const parts: string[] = [];

  if (layout.position) parts.push(`position:${layout.position}`);
  if (layout.left != null) parts.push(`left:${twipToPt(layout.left)}pt`);
  if (layout.top != null) parts.push(`top:${twipToPt(layout.top)}pt`);
  if (layout.marginLeft != null) parts.push(`margin-left:${twipToPt(layout.marginLeft)}pt`);
  if (layout.marginTop != null) parts.push(`margin-top:${twipToPt(layout.marginTop)}pt`);
  if (layout.marginRight != null) parts.push(`margin-right:${twipToPt(layout.marginRight)}pt`);
  if (layout.marginBottom != null) parts.push(`margin-bottom:${twipToPt(layout.marginBottom)}pt`);
  if (layout.width != null) parts.push(`width:${twipToPt(layout.width)}pt`);
  if (layout.height != null) parts.push(`height:${twipToPt(layout.height)}pt`);
  if (layout.zIndex != null) parts.push(`z-index:${layout.zIndex}`);
  if (layout.rotation != null) parts.push(`rotation:${layout.rotation}`);

  if (layout.wrapping) {
    const wrapMap: Record<ShapeWrapping, string> = {
      inline: "none",
      square: "square",
      tight: "tight",
      through: "through",
      topAndBottom: "top-and-bottom",
      behind: "none",
      inFront: "none",
    };
    parts.push(`mso-wrap-style:${wrapMap[layout.wrapping] ?? layout.wrapping}`);
  }

  if (layout.horizontalPosition) parts.push(`mso-position-horizontal:${layout.horizontalPosition}`);
  if (layout.verticalPosition) parts.push(`mso-position-vertical:${layout.verticalPosition}`);
  if (layout.horizontalRelative)
    parts.push(`mso-position-horizontal-relative:${layout.horizontalRelative}`);
  if (layout.verticalRelative)
    parts.push(`mso-position-vertical-relative:${layout.verticalRelative}`);

  if (layout.rawStyle) {
    for (const [k, v] of Object.entries(layout.rawStyle)) {
      const kebab = k.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`);
      if (v != null && !parts.some((p) => p.startsWith(`${k}:`) || p.startsWith(`${kebab}:`))) {
        parts.push(`${kebab}:${v}`);
      }
    }
  }

  return parts.join(";");
}
