// PDF vector scene — the serializable bridge between the live LeaferJS scene
// graph (the canvas's own vector IR: Path/Rect/Line/Text/Image/Group/Box) and
// the PDF content-stream writer in export-pdf.ts.
//
// The canvas scene is walked once per page at export time (canvas/scene-export)
// and flattened into these plain nodes: every shape carries an absolute
// page-space matrix plus its local geometry — the same contract Leafer renders
// from — so this module stays pure (no DOM, no Leafer import) and is unit
// testable in Node. Page coordinates are CSS px with y growing downward; the
// emitter wraps the page in `0.75 0 0 -0.75 0 H cm` (px → pt, top-left →
// bottom-left), so every node below is written verbatim in px.
//
// This module owns: path-data → PDF operators (with quadratic→cubic and
// arc→cubic conversion), color parsing, opacity/blend ExtGStates, image
// placement (JPEG passthrough / RGBA+SMask), and visible Text rows (each row
// is a laid-out baseline from Leafer's own text layout — no re-measuring).

/** A 2-D affine matrix [a b c d e f] — PDF's `cm` operand order. */
export interface PdfMatrix {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

/** One decoded bitmap the scene references. `jpeg` is passed through with
 *  /DCTDecode; otherwise `rgba` is Flate-encoded (and gets an /SMask when it
 *  carries alpha). `key` dedupes the same bytes across pages. */
export interface PdfSceneImageData {
  readonly key: string;
  readonly width: number;
  readonly height: number;
  readonly jpeg?: Uint8Array;
  readonly rgba?: Uint8Array;
  /** True when `rgba` has at least one non-opaque pixel (→ /SMask). */
  readonly hasAlpha?: boolean;
}

interface PdfSceneNodeBase {
  readonly opacity?: number;
  readonly blendMode?: string;
}

/** A group: children already carry absolute matrices; `clip` (a Box with
 *  overflow hidden) clips them to the box path. */
export interface PdfSceneGroupNode extends PdfSceneNodeBase {
  readonly type: "group";
  readonly clip?: { readonly path: string; readonly matrix: PdfMatrix };
  readonly children: readonly PdfSceneNode[];
}

/** One filled/stroked vector path. `glyph` marks text outlines painted by
 *  the glyph painter — the font-text measurement mode drops them. */
export interface PdfSceneShapeNode extends PdfSceneNodeBase {
  readonly type: "shape";
  readonly matrix: PdfMatrix;
  readonly path: string;
  readonly fill?: string;
  readonly stroke?: string;
  readonly strokeWidth?: number;
  readonly dash?: readonly number[];
  readonly cap?: "butt" | "round" | "square";
  readonly join?: "miter" | "round" | "bevel";
  readonly winding?: "nonzero" | "evenodd";
  readonly glyph?: boolean;
}

export interface PdfSceneImageNode extends PdfSceneNodeBase {
  readonly type: "image";
  readonly matrix: PdfMatrix;
  readonly image: PdfSceneImageData;
  /** The element's display box in local px (the image stretches to it). */
  readonly width: number;
  readonly height: number;
}

/** One text row: the baseline origin `(x, y)` in the Text element's local
 *  space (from Leafer's own row layout — alignment is already applied). */
export interface PdfSceneTextRow {
  readonly x: number;
  readonly y: number;
  readonly text: string;
  readonly width: number;
  /** Justified/squeezed rows only: the extra advance (px in the element's
   *  text space, may be negative) the row inserts after each grapheme of
   *  `text` — Leafer CharLayout's uniform per-letter share or per-word-gap
   *  share. Absent = natural advances (plain `Tj`). */
  readonly extras?: readonly number[];
}

export interface PdfSceneTextNode extends PdfSceneNodeBase {
  readonly type: "text";
  readonly matrix: PdfMatrix;
  readonly rows: readonly PdfSceneTextRow[];
  readonly fontSize: number;
  readonly fontFamily?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly letterSpacing?: number;
  readonly fill?: string;
  readonly stroke?: string;
  readonly strokeWidth?: number;
}

export type PdfSceneNode =
  | PdfSceneGroupNode
  | PdfSceneShapeNode
  | PdfSceneImageNode
  | PdfSceneTextNode;

/** One row from Leafer's own text layout (`textDrawData.rows`), as read by
 *  the scene walkers. Plain rows carry `text`; char-mode rows (wrapping,
 *  letter spacing, an explicit box) carry their glyphs in `data` instead —
 *  Leafer drops the spaces from it (they are advances, not glyphs), so the
 *  element's source string has to be walked in parallel to re-insert them. */
export interface LeaferTextRow {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly text?: string;
  readonly extras?: readonly number[];
  readonly data?: readonly { readonly char?: string }[];
  readonly words?: readonly { readonly data?: readonly { readonly char?: string }[] }[];
}

/** Rebuild one element's row strings. Plain rows carry their own `text`; in
 *  char mode Leafer strips spaces from `data`, so the element's source string
 *  is walked in parallel to re-insert them (spaces are gaps in char mode, and
 *  a PDF Tj must carry them explicitly — the row's glyph positions are the
 *  face's own advances, see the /W widths the exporter emits). `data` entries
 *  are per word when Leafer justified the row (each entry spans the word's
 *  chars), so the cursor advances by the entry's length, not one code unit. */
export function rowTexts(elementText: string, rows: readonly LeaferTextRow[]): string[] {
  const out: string[] = [];
  let cursor = 0;
  for (const row of rows) {
    if (typeof row.text === "string" && row.text.length > 0) {
      out.push(row.text);
      cursor += row.text.length;
      if (elementText[cursor] === "\n") cursor++;
      continue;
    }
    const chars = row.data ?? [];
    let text = "";
    for (const entry of chars) {
      const char = entry?.char;
      if (char === " ") {
        // Overflow char-mode rows keep the space entry itself; its source
        // slot is the same space.
        if (elementText[cursor] === " ") cursor++;
        text += " ";
        continue;
      }
      // Spaces between entries are the gaps Leafer dropped from `data`.
      while (cursor < elementText.length && elementText[cursor] === " ") {
        text += " ";
        cursor++;
      }
      if (elementText[cursor] === "\n") {
        cursor++;
        break;
      }
      if (typeof char !== "string" || char.length === 0) continue;
      text += char;
      cursor += Math.min(char.length, elementText.length - cursor);
    }
    // Trailing source spaces belong to this row (Leafer trims them from the
    // glyph run, but they are part of the row's text).
    while (cursor < elementText.length && elementText[cursor] === " ") {
      text += " ";
      cursor++;
    }
    if (elementText[cursor] === "\n") cursor++;
    out.push(text);
  }
  return out;
}

/** One page's serialized scene — CSS px, y down. */
export interface PdfScenePage {
  readonly width: number;
  readonly height: number;
  readonly nodes: readonly PdfSceneNode[];
  /** Coverage diagnostics: Leafer tags the walk skipped (should stay empty). */
  readonly skipped?: Readonly<Record<string, number>>;
}

/** One ExtGState the page's content uses: fill/stroke alpha + blend mode. */
export interface PdfExtGStateSpec {
  readonly fill: number;
  readonly stroke: number;
  readonly blend?: string;
}

/** The font resource one visible text node draws with. `isUnicode` picks
 *  hex UTF-16 CIDs (Identity-H) over a WinAnsi literal string. */
export interface PdfTextFontRef {
  readonly resource: string;
  readonly isUnicode: boolean;
}

export interface PdfScenePlan {
  /** The scene's content-stream body (no page-level wrapper). */
  readonly content: string;
  /** Page-local image resources, `/Im0`, `/Im1`, … */
  readonly images: readonly PdfSceneImageData[];
  /** Page-local ExtGStates, `/GS0`, `/GS1`, … */
  readonly extGStates: readonly PdfExtGStateSpec[];
  /** Font resource names the visible text referenced (informational). */
  readonly fonts: readonly string[];
}

/** Format a number for a content stream: fixed precision, trailing zeros
 *  trimmed — vector scenes carry a lot of coordinates, and every byte counts. */
export function pdfNum(value: number, precision = 4): string {
  if (!Number.isFinite(value)) return "0";
  const fixed = value.toFixed(precision);
  if (!fixed.includes(".")) return fixed;
  return fixed.replace(/0+$/, "").replace(/\.$/, "");
}

const matrixOperands = (m: PdfMatrix): string =>
  `${pdfNum(m.a, 6)} ${pdfNum(m.b, 6)} ${pdfNum(m.c, 6)} ${pdfNum(m.d, 6)} ${pdfNum(m.e, 4)} ${pdfNum(m.f, 4)}`;

/** Identity matrix constant. */
export const PDF_IDENTITY: PdfMatrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

/** Compose two affine matrices: applying `second` then `first` (PDF `cm`
 *  semantics — the new CTM is the operand pre-multiplied into the current). */
export function composeMatrix(first: PdfMatrix, second: PdfMatrix): PdfMatrix {
  return {
    a: first.a * second.a + first.c * second.b,
    b: first.b * second.a + first.d * second.b,
    c: first.a * second.c + first.c * second.d,
    d: first.b * second.c + first.d * second.d,
    e: first.a * second.e + first.c * second.f + first.e,
    f: first.b * second.e + first.d * second.f + first.f,
  };
}

/** Invert a matrix; returns identity for a degenerate one. */
export function invertMatrix(m: PdfMatrix): PdfMatrix {
  const det = m.a * m.d - m.b * m.c;
  if (det === 0 || !Number.isFinite(det)) return PDF_IDENTITY;
  return {
    a: m.d / det,
    b: -m.b / det,
    c: -m.c / det,
    d: m.a / det,
    e: (m.c * m.f - m.d * m.e) / det,
    f: (m.b * m.e - m.a * m.f) / det,
  };
}

/** An sRGB color with straight-alpha coverage. */
export interface PdfRgba {
  r: number;
  g: number;
  b: number;
  a: number;
}

/** Parse the color strings the paint layer emits: #rgb/#rgba/#rrggbb/#rrggbbaa,
 *  rgb()/rgba(), "transparent" and the handful of CSS names the painters use.
 *  Unknown strings return undefined (the caller skips that paint). */
export function parseColor(input: string | undefined): PdfRgba | undefined {
  if (!input) return undefined;
  const value = input.trim().toLowerCase();
  if (value === "none" || value === "transparent") return undefined;
  if (value.startsWith("#")) {
    const hex = value.slice(1);
    if (hex.length === 3 || hex.length === 4) {
      const r = parseInt(hex[0]! + hex[0]!, 16);
      const g = parseInt(hex[1]! + hex[1]!, 16);
      const b = parseInt(hex[2]! + hex[2]!, 16);
      const a = hex.length === 4 ? parseInt(hex[3]! + hex[3]!, 16) / 255 : 1;
      return { r, g, b, a };
    }
    if (hex.length === 6 || hex.length === 8) {
      const r = parseInt(hex.slice(0, 2), 16);
      const g = parseInt(hex.slice(2, 4), 16);
      const b = parseInt(hex.slice(4, 6), 16);
      const a = hex.length === 8 ? parseInt(hex.slice(6, 8), 16) / 255 : 1;
      if ([r, g, b, a].every((n) => Number.isFinite(n))) return { r, g, b, a };
      return undefined;
    }
    return undefined;
  }
  if (
    /^[0-9a-f]{3,8}$/i.test(value) &&
    (value.length === 3 || value.length === 4 || value.length === 6 || value.length === 8)
  ) {
    return parseColor("#" + value);
  }
  const fn = /^rgba?\(([^)]+)\)$/.exec(value);
  if (fn) {
    const parts = fn[1]!
      .split(/[,\s/]+/)
      .filter(Boolean)
      .map(Number);
    if (parts.length >= 3 && parts.slice(0, 3).every((n) => Number.isFinite(n))) {
      return {
        r: parts[0]!,
        g: parts[1]!,
        b: parts[2]!,
        a: parts.length >= 4 && Number.isFinite(parts[3]) ? parts[3]! : 1,
      };
    }
    return undefined;
  }
  const named: Record<string, string> = {
    black: "#000000",
    white: "#ffffff",
    red: "#ff0000",
    green: "#008000",
    blue: "#0000ff",
    gray: "#808080",
    grey: "#808080",
    silver: "#c0c0c0",
    yellow: "#ffff00",
    orange: "#ffa500",
    purple: "#800080",
  };
  const mapped = named[value];
  return mapped ? parseColor(mapped) : undefined;
}

/** PDF color operator operands: channels normalized to 0..1. */
function colorOperands(color: PdfRgba): string {
  return `${pdfNum(color.r / 255)} ${pdfNum(color.g / 255)} ${pdfNum(color.b / 255)}`;
}

/** Leafer blend modes → PDF blend mode names. */
const PDF_BLEND_MODES: Readonly<Record<string, string>> = {
  normal: "Normal",
  multiply: "Multiply",
  screen: "Screen",
  overlay: "Overlay",
  darken: "Darken",
  lighten: "Lighten",
  "color-dodge": "ColorDodge",
  "color-burn": "ColorBurn",
  "hard-light": "HardLight",
  "soft-light": "SoftLight",
  difference: "Difference",
  exclusion: "Exclusion",
  hue: "Hue",
  saturation: "Saturation",
  color: "Color",
  luminosity: "Luminosity",
};

/** The `<< … >>` body of one ExtGState. */
export function extGStateDict(spec: PdfExtGStateSpec): string {
  const parts = [`/Type /ExtGState`, `/ca ${pdfNum(spec.fill)}`, `/CA ${pdfNum(spec.stroke)}`];
  const blend = spec.blend ? PDF_BLEND_MODES[spec.blend] : undefined;
  parts.push(`/BM /${blend ?? "Normal"}`);
  return `<< ${parts.join(" ")} >>`;
}

// ── SVG path data → PDF operators ────────────────────────────────────────────

const PATH_TOKEN = /[MmLlHhVvCcSsQqTtAaZz]|[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?/g;

/** Convert SVG path data (the format both Leafer's `getPathString` and the
 *  glyph painter emit) into PDF path operators. Supports the full command set:
 *  quadratics elevate to cubics, elliptical arcs convert to cubic segments. */
export function svgPathToPdf(path: string): string {
  const tokens = path.match(PATH_TOKEN);
  if (!tokens || tokens.length === 0) return "";
  const out: string[] = [];
  let i = 0;
  let x = 0;
  let y = 0;
  let startX = 0;
  let startY = 0;
  // Reflection memory for S/T: the second control point of the previous
  // cubic (or the control point of the previous quadratic).
  let ctrlX = 0;
  let ctrlY = 0;
  let lastCubic = false;
  let lastQuad = false;

  const num = (): number => Number(tokens[i++]);
  const hasNum = (): boolean => i < tokens.length && !/[MmLlHhVvCcSsQqTtAaZz]/.test(tokens[i]!);
  const moveTo = (nx: number, ny: number): void => {
    out.push(`${pdfNum(nx, 3)} ${pdfNum(ny, 3)} m`);
    x = startX = nx;
    y = startY = ny;
  };
  const lineTo = (nx: number, ny: number): void => {
    out.push(`${pdfNum(nx, 3)} ${pdfNum(ny, 3)} l`);
    x = nx;
    y = ny;
  };
  const curveTo = (
    c1x: number,
    c1y: number,
    c2x: number,
    c2y: number,
    nx: number,
    ny: number,
  ): void => {
    out.push(
      `${pdfNum(c1x, 3)} ${pdfNum(c1y, 3)} ${pdfNum(c2x, 3)} ${pdfNum(c2y, 3)} ` +
        `${pdfNum(nx, 3)} ${pdfNum(ny, 3)} c`,
    );
    ctrlX = c2x;
    ctrlY = c2y;
    x = nx;
    y = ny;
  };
  const quadTo = (qx: number, qy: number, nx: number, ny: number): void => {
    // Elevate the quadratic to a cubic: c1 = p0 + 2/3(q - p0), c2 = p1 + 2/3(q - p1).
    const c1x = x + (2 / 3) * (qx - x);
    const c1y = y + (2 / 3) * (qy - y);
    const c2x = nx + (2 / 3) * (qx - nx);
    const c2y = ny + (2 / 3) * (qy - ny);
    curveTo(c1x, c1y, c2x, c2y, nx, ny);
    // The reflection memory keeps the QUADRATIC control (used by T/S).
    ctrlX = qx;
    ctrlY = qy;
  };

  while (i < tokens.length) {
    const cmd = tokens[i++]!;
    switch (cmd) {
      case "M":
      case "m":
        if (!hasNum()) break;
        moveTo(num() + (cmd === "m" ? x : 0), num() + (cmd === "m" ? y : 0));
        // implicit lineto for the remaining pairs
        while (hasNum()) lineTo(num() + (cmd === "m" ? x : 0), num() + (cmd === "m" ? y : 0));
        break;
      case "L":
      case "l":
        while (hasNum()) lineTo(num() + (cmd === "l" ? x : 0), num() + (cmd === "l" ? y : 0));
        break;
      case "H":
      case "h":
        while (hasNum()) lineTo(num() + (cmd === "h" ? x : 0), y);
        break;
      case "V":
      case "v":
        while (hasNum()) lineTo(x, num() + (cmd === "v" ? y : 0));
        break;
      case "C":
      case "c": {
        while (hasNum()) {
          const rel = cmd === "c";
          const x1 = num() + (rel ? x : 0);
          const y1 = num() + (rel ? y : 0);
          const x2 = num() + (rel ? x : 0);
          const y2 = num() + (rel ? y : 0);
          const nx = num() + (rel ? x : 0);
          const ny = num() + (rel ? y : 0);
          curveTo(x1, y1, x2, y2, nx, ny);
        }
        lastCubic = true;
        lastQuad = false;
        continue;
      }
      case "S":
      case "s": {
        while (hasNum()) {
          const rel = cmd === "s";
          const x2 = num() + (rel ? x : 0);
          const y2 = num() + (rel ? y : 0);
          const nx = num() + (rel ? x : 0);
          const ny = num() + (rel ? y : 0);
          const x1 = lastCubic ? 2 * x - ctrlX : x;
          const y1 = lastCubic ? 2 * y - ctrlY : y;
          curveTo(x1, y1, x2, y2, nx, ny);
        }
        lastCubic = true;
        lastQuad = false;
        continue;
      }
      case "Q":
      case "q": {
        while (hasNum()) {
          const rel = cmd === "q";
          const qx = num() + (rel ? x : 0);
          const qy = num() + (rel ? y : 0);
          const nx = num() + (rel ? x : 0);
          const ny = num() + (rel ? y : 0);
          quadTo(qx, qy, nx, ny);
        }
        lastQuad = true;
        lastCubic = false;
        continue;
      }
      case "T":
      case "t": {
        while (hasNum()) {
          const rel = cmd === "t";
          const qx = lastQuad ? 2 * x - ctrlX : x;
          const qy = lastQuad ? 2 * y - ctrlY : y;
          const nx = num() + (rel ? x : 0);
          const ny = num() + (rel ? y : 0);
          quadTo(qx, qy, nx, ny);
        }
        lastQuad = true;
        lastCubic = false;
        continue;
      }
      case "A":
      case "a": {
        while (hasNum()) {
          const rel = cmd === "a";
          const rx = Math.abs(num());
          const ry = Math.abs(num());
          const rotation = num();
          const largeArc = num() !== 0;
          const sweep = num() !== 0;
          const nx = num() + (rel ? x : 0);
          const ny = num() + (rel ? y : 0);
          if (rx === 0 || ry === 0) {
            lineTo(nx, ny);
          } else {
            arcToCubics(x, y, rx, ry, rotation, largeArc, sweep, nx, ny, out);
            x = nx;
            y = ny;
          }
        }
        break;
      }
      case "Z":
      case "z":
        out.push("h");
        x = startX;
        y = startY;
        break;
      default:
        break;
    }
    lastCubic = false;
    lastQuad = false;
  }
  return out.join(" ");
}

/** SVG endpoint arc → up to four cubic segments (F.6.5 of the SVG spec). */
function arcToCubics(
  x0: number,
  y0: number,
  rx: number,
  ry: number,
  rotationDeg: number,
  largeArc: boolean,
  sweep: boolean,
  x1: number,
  y1: number,
  out: string[],
): void {
  if (x0 === x1 && y0 === y1) return;
  const phi = (rotationDeg * Math.PI) / 180;
  const cosPhi = Math.cos(phi);
  const sinPhi = Math.sin(phi);
  const dx2 = (x0 - x1) / 2;
  const dy2 = (y0 - y1) / 2;
  const x1p = cosPhi * dx2 + sinPhi * dy2;
  const y1p = -sinPhi * dx2 + cosPhi * dy2;
  let rxSq = rx * rx;
  let rySq = ry * ry;
  const x1pSq = x1p * x1p;
  const y1pSq = y1p * y1p;
  const lambda = x1pSq / rxSq + y1pSq / rySq;
  if (lambda > 1) {
    const scale = Math.sqrt(lambda);
    rx *= scale;
    ry *= scale;
    rxSq = rx * rx;
    rySq = ry * ry;
  }
  const sign = largeArc !== sweep ? 1 : -1;
  const numerator = Math.max(0, rxSq * rySq - rxSq * y1pSq - rySq * x1pSq);
  const denominator = rxSq * y1pSq + rySq * x1pSq;
  const coef = denominator === 0 ? 0 : sign * Math.sqrt(numerator / denominator);
  const cxp = (coef * (rx * y1p)) / ry;
  const cyp = (coef * (-ry * x1p)) / rx;
  const cx = cosPhi * cxp - sinPhi * cyp + (x0 + x1) / 2;
  const cy = sinPhi * cxp + cosPhi * cyp + (y0 + y1) / 2;
  const theta1 = Math.atan2((y1p - cyp) / ry, (x1p - cxp) / rx);
  const theta2 = Math.atan2((-y1p - cyp) / ry, (-x1p - cxp) / rx);
  let delta = theta2 - theta1;
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  if (sweep && delta < 0) delta += 2 * Math.PI;
  const segments = Math.max(1, Math.ceil(Math.abs(delta) / (Math.PI / 2)));
  const step = delta / segments;
  const k = (4 / 3) * Math.tan(step / 4);
  let theta = theta1;
  const point = (t: number): [number, number] => [
    cx + rx * Math.cos(t) * cosPhi - ry * Math.sin(t) * sinPhi,
    cy + rx * Math.cos(t) * sinPhi + ry * Math.sin(t) * cosPhi,
  ];
  const derivative = (t: number): [number, number] => [
    -rx * Math.sin(t) * cosPhi - ry * Math.cos(t) * sinPhi,
    -rx * Math.sin(t) * sinPhi + ry * Math.cos(t) * cosPhi,
  ];
  for (let s = 0; s < segments; s++) {
    const next = theta + step;
    const p1 = point(theta);
    const p2 = point(next);
    const d1 = derivative(theta);
    const d2 = derivative(next);
    const c1x = p1[0] + k * d1[0];
    const c1y = p1[1] + k * d1[1];
    const c2x = p2[0] - k * d2[0];
    const c2y = p2[1] - k * d2[1];
    out.push(
      `${pdfNum(c1x, 3)} ${pdfNum(c1y, 3)} ${pdfNum(c2x, 3)} ${pdfNum(c2y, 3)} ` +
        `${pdfNum(p2[0], 3)} ${pdfNum(p2[1], 3)} c`,
    );
    theta = next;
  }
}

// ── Content planning ─────────────────────────────────────────────────────────

/** Standard-14 resource selection for a family/weight/slant — the same
 *  heuristics the invisible text layer uses (kept here so the visible path and
 *  the text layer can never drift). */
export function standardFontFor(
  family: string | undefined,
  bold?: boolean,
  italic?: boolean,
): string {
  const fam = (family ?? "").toLowerCase();
  if (
    fam.includes("times") ||
    fam.includes("serif") ||
    fam.includes("georgia") ||
    fam.includes("garamond") ||
    fam.includes("cambria") ||
    fam.includes("caladea")
  ) {
    if (bold && italic) return "F8";
    if (bold) return "F6";
    if (italic) return "F7";
    return "F5";
  }
  if (fam.includes("courier") || fam.includes("mono") || fam.includes("consolas")) {
    if (bold) return "F10";
    return "F9";
  }
  if (bold && italic) return "F4";
  if (bold) return "F2";
  if (italic) return "F3";
  return "F1";
}

/** ASCII-printable test shared by the text encoders. */
function isAsciiPrintable(str: string): boolean {
  return /^[\x20-\x7E]*$/.test(str);
}

/** Escape special characters for PDF literal strings ( ... ). */
function escapeLiteral(str: string): string {
  return str
    .replace(/\\/g, "\\\\")
    .replace(/\(/g, "\\(")
    .replace(/\)/g, "\\)")
    .replace(/\r/g, "\\r")
    .replace(/\n/g, "\\n");
}

/** 2-byte hexadecimal CIDs for /Identity-H fonts. */
function encodeHexUtf16(str: string): string {
  let hex = "";
  for (let i = 0; i < str.length; i++) hex += str.charCodeAt(i).toString(16).padStart(4, "0");
  return hex;
}

/** Grapheme segmentation for a justified row's TJ runs — the same unit the
 *  node kit's row `extras` are keyed by. */
const GRAPHEME_SEGMENTER = new Intl.Segmenter(undefined, { granularity: "grapheme" });

/** Emit one text row — `row.x/y` is the baseline in the node's local space,
 *  the node matrix maps local → page, and the page wrapper flips page px to
 *  PDF pt; the row's own `1 0 0 -1` counter-flip keeps glyphs upright.
 *  Fill-only is `Tr 0`, stroke-only leafer outlines are `Tr 1`, both `Tr 2`. */
function emitTextRow(
  out: string[],
  node: PdfSceneTextNode,
  row: PdfSceneTextRow,
  font: PdfTextFontRef,
  fill: PdfRgba | undefined,
  stroke: PdfRgba | undefined,
): void {
  out.push("BT");
  if (node.letterSpacing) out.push(`${pdfNum(node.letterSpacing, 3)} Tc`);
  if (fill && stroke) out.push("2 Tr");
  else if (stroke) out.push("1 Tr");
  if (fill) out.push(`${colorOperands(fill)} rg`);
  if (stroke) {
    out.push(`${colorOperands(stroke)} RG`);
    const strokeWidth = node.strokeWidth ?? 1;
    out.push(`${pdfNum(strokeWidth, 3)} w`);
  }
  out.push(`/${font.resource} ${pdfNum(node.fontSize, 3)} Tf`);
  out.push(`1 0 0 -1 ${pdfNum(row.x, 3)} ${pdfNum(row.y, 3)} Tm`);
  const extras = row.extras;
  const parts: string[] = [];
  if (extras) {
    for (const { segment } of GRAPHEME_SEGMENTER.segment(row.text)) parts.push(segment);
  }
  if (extras && extras.length === parts.length && parts.length > 1) {
    // Justified/squeezed row: the pen advances by the natural glyph widths
    // plus each grapheme's extra share (Leafer's char/word distribution),
    // expressed as TJ displacement adjustments in thousandths of the font
    // size. Consecutive zero-extra graphemes stay one string.
    const items: string[] = [];
    let run = "";
    const flush = (): void => {
      if (!run) return;
      items.push(font.isUnicode ? `<${encodeHexUtf16(run)}>` : `(${escapeLiteral(run)})`);
      run = "";
    };
    for (let i = 0; i < parts.length; i++) {
      run += parts[i]!;
      const extra = extras[i] ?? 0;
      if (extra !== 0) {
        flush();
        items.push(pdfNum((-extra * 1000) / node.fontSize, 4));
      }
    }
    flush();
    out.push(`[ ${items.join(" ")} ] TJ`);
  } else if (font.isUnicode) {
    out.push(`<${encodeHexUtf16(row.text)}> Tj`);
  } else {
    out.push(`(${escapeLiteral(row.text)}) Tj`);
  }
  out.push("ET");
}

/** Build one page's content stream and resource lists. `fontForText` resolves
 *  a visible text node to a page font resource (embedded subset when the host
 *  has the face, standard-14 otherwise). `includeGlyphs` false drops the
 *  glyph-outline shapes — the embedded-font text measurement mode, which lets
 *  the invisible span layer render visibly instead. */
export function planSceneContent(
  page: PdfScenePage,
  options: {
    readonly fontForText: (node: PdfSceneTextNode) => PdfTextFontRef;
    readonly includeGlyphs?: boolean;
  },
): PdfScenePlan {
  const out: string[] = [];
  const images: PdfSceneImageData[] = [];
  const imageIndex = new Map<string, number>();
  const extGStates: PdfExtGStateSpec[] = [];
  const stateIndex = new Map<string, number>();
  const fonts = new Set<string>();

  const stateFor = (node: PdfSceneNodeBase, fillA: number, strokeA: number): string => {
    const blend = node.blendMode && node.blendMode !== "pass-through" ? node.blendMode : undefined;
    const key = `${pdfNum(fillA)}|${pdfNum(strokeA)}|${blend ?? ""}`;
    const existing = stateIndex.get(key);
    if (existing !== undefined) return `/GS${existing}`;
    extGStates.push({
      fill: fillA,
      stroke: strokeA,
      ...(blend ? { blend } : {}),
    });
    const index = extGStates.length - 1;
    stateIndex.set(key, index);
    return `/GS${index}`;
  };

  const imageRef = (image: PdfSceneImageData): string => {
    let index = imageIndex.get(image.key);
    if (index === undefined) {
      images.push(image);
      index = images.length - 1;
      imageIndex.set(image.key, index);
    }
    return `/Im${index}`;
  };

  const emitNode = (node: PdfSceneNode): void => {
    const opacity = node.opacity ?? 1;
    if (opacity <= 0) return;
    switch (node.type) {
      case "group": {
        out.push("q");
        if (node.clip) {
          out.push(`${matrixOperands(node.clip.matrix)} cm`);
          out.push(`${svgPathToPdf(node.clip.path)} W n`);
          // Children carry absolute page matrices: undo the clip's transform
          // (the clip itself stays in device space).
          out.push(`${matrixOperands(invertMatrix(node.clip.matrix))} cm`);
        }
        for (const child of node.children) emitNode(child);
        out.push("Q");
        return;
      }
      case "shape": {
        if (node.glyph && options.includeGlyphs === false) return;
        const fill = parseColor(node.fill);
        const stroke = parseColor(node.stroke);
        const fillAlpha = (fill?.a ?? 1) * opacity;
        const strokeAlpha = (stroke?.a ?? 1) * opacity;
        if (!fill && !stroke) return;
        out.push("q");
        const state = stateFor(node, fillAlpha, strokeAlpha);
        out.push(`${state} gs`);
        if (
          node.matrix &&
          (node.matrix.a !== 1 ||
            node.matrix.b !== 0 ||
            node.matrix.c !== 0 ||
            node.matrix.d !== 1 ||
            node.matrix.e !== 0 ||
            node.matrix.f !== 0)
        ) {
          out.push(`${matrixOperands(node.matrix)} cm`);
        }
        const ops = svgPathToPdf(node.path);
        if (ops) {
          out.push(ops);
          if (fill) out.push(`${colorOperands(fill)} rg`);
          if (stroke) {
            out.push(`${colorOperands(stroke)} RG`);
            const width = node.strokeWidth ?? 1;
            out.push(`${pdfNum(width, 3)} w`);
            if (node.join)
              out.push(`${node.join === "round" ? 1 : node.join === "bevel" ? 2 : 0} j`);
            if (node.cap) out.push(`${node.cap === "round" ? 1 : node.cap === "square" ? 2 : 0} J`);
            if (node.dash && node.dash.length > 0) {
              out.push(`[ ${node.dash.map((d) => pdfNum(d, 3)).join(" ")} ] 0 d`);
            }
          }
          if (fill && stroke) out.push("B");
          else if (fill) out.push(node.winding === "evenodd" ? "f*" : "f");
          else out.push("S");
        }
        out.push("Q");
        return;
      }
      case "image": {
        out.push("q");
        const state = stateFor(node, opacity, opacity);
        out.push(`${state} gs`);
        out.push(`${matrixOperands(node.matrix)} cm`);
        // The element box (display size) — the image data's own pixel size
        // lives on the XObject; here the box maps the unit square onto it.
        const width = node.width;
        const height = node.height;
        out.push(
          `${pdfNum(width, 3)} 0 0 ${pdfNum(-height, 3)} 0 ${pdfNum(height, 3)} cm ${imageRef(node.image)} Do`,
        );
        out.push("Q");
        return;
      }
      case "text": {
        const fill = parseColor(node.fill);
        const stroke = parseColor(node.stroke);
        if (!fill && !stroke) return;
        const fillAlpha = (fill?.a ?? 1) * opacity;
        const strokeAlpha = (stroke?.a ?? 1) * opacity;
        let font = options.fontForText(node);
        // A non-ASCII row routed to a standard WinAnsi face would drop its
        // glyphs; the universal Identity-H face carries the Unicode mapping.
        for (const row of node.rows) {
          if (!font.isUnicode && !isAsciiPrintable(row.text)) {
            font = { resource: "F_Uni", isUnicode: true };
            break;
          }
        }
        out.push("q");
        const state = stateFor(node, fillAlpha, strokeAlpha);
        out.push(`${state} gs`);
        if (
          node.matrix &&
          (node.matrix.a !== 1 ||
            node.matrix.b !== 0 ||
            node.matrix.c !== 0 ||
            node.matrix.d !== 1 ||
            node.matrix.e !== 0 ||
            node.matrix.f !== 0)
        ) {
          out.push(`${matrixOperands(node.matrix)} cm`);
        }
        for (const row of node.rows) {
          if (!row.text) continue;
          emitTextRow(out, node, row, font, fill, stroke);
        }
        out.push("Q");
        fonts.add(font.resource);
        return;
      }
      default:
        return;
    }
  };

  for (const node of page.nodes) emitNode(node);
  return { content: out.join("\n"), images, extGStates, fonts: [...fonts] };
}

/** Wrap a scene's content with the page-space transform (px → pt, y flip)
 *  and an artifact marker — the painted scene is decorative; the invisible
 *  text layer carries the document's structure. */
export function wrapSceneContent(content: string, pageHeightPx: number): string {
  const height = pdfNum(pageHeightPx * 0.75, 2);
  return `q\n/Artifact BMC\n` + `0.75 0 0 -0.75 0 ${height} cm\n` + content + `\nEMC\nQ\n`;
}
