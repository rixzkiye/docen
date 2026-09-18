import { EMU_PER_PX } from "@docen/layout";

/**
 * Word's Text Effects (Home → Font → Text Effects and Typography) — the w14
 * extended run properties (MS-DOCX §5.1, namespace
 * `http://schemas.microsoft.com/office/word/2010/wordml`).
 *
 * office-open carries every `w14:*` child of `<w:rPr>` verbatim as
 * `RunOptions.w14RawXml`; TextStyle mirrors it, so round-trip is already
 * lossless. This module adds the structured layer: it parses the raw XML into
 * the effect set the layout/painter consume (outline, shadow, glow,
 * reflection, 3-D bevel, 3-D rotation) and serializes Word-shaped chunks when
 * the editor applies or clears one effect. Untouched effects always survive
 * byte-for-byte — {@link setTextEffectChunk} only swaps the addressed chunk.
 */

/** The effect families the gallery edits. */
export type TextEffectKind =
  | "outline"
  | "shadow"
  | "glow"
  | "reflection"
  | "bevel"
  | "rotation"
  | "fill";

/** The w14 element carrying each family. */
export const TEXT_EFFECT_ELEMENT: Readonly<Record<TextEffectKind, string>> = {
  outline: "textOutline",
  shadow: "shadow",
  glow: "glow",
  reflection: "reflection",
  bevel: "props3d",
  rotation: "scene3d",
  fill: "textFill",
};

/** `w14:textOutline` — the character outline stroke. */
export interface TextEffectOutline {
  /** Hex RRGGBB; absent with `hollow` = the run's own text color. */
  color?: string;
  /** Stroke width in px (w14:w EMU resolved). */
  widthPx: number;
  /** The fill inside the stroke is dropped (noFill) — hollow letters. */
  hollow?: boolean;
  /** ST_PresetLineDashVal (solid/dot/dash/…); absent = solid. */
  dash?: string;
}

/** `w14:shadow` — the drop shadow (CT_Shadow attrs). */
export interface TextEffectShadow {
  color: string;
  /** 0-1 (alpha resolved). */
  opacity: number;
  blurPx: number;
  distPx: number;
  /** Degrees, 0 = east, clockwise (DrawingML positive fixed angle). */
  dirDeg: number;
}

/** `w14:glow` — the halo. */
export interface TextEffectGlow {
  color: string;
  opacity: number;
  radiusPx: number;
}

/** `w14:reflection` — the mirrored copy under the run. */
export interface TextEffectReflection {
  blurPx: number;
  /** 0-1 starting opacity (stA). */
  startOpacity: number;
  /** 0-1 ending opacity (endA). */
  endOpacity: number;
  distPx: number;
}

/** One `w14:bevelT`/`w14:bevelB` edge profile. */
export interface TextEffectBevel {
  widthPx: number;
  heightPx: number;
  /** ST_BevelPresetType (circle/cross/angle/softRound/…). */
  preset: string;
}

/** `w14:props3d` — the 3-D Format bevels. */
export interface TextEffectBevels {
  top?: TextEffectBevel;
  bottom?: TextEffectBevel;
}

/** `w14:scene3d` — the 3-D rotation (camera preset + light-rig sphere). */
export interface TextEffectRotation {
  /** Degrees about the horizontal axis (lightRig/rot @lat). */
  x: number;
  /** Degrees about the vertical axis (@lon). */
  y: number;
  /** Degrees about the view axis (@rev). */
  z: number;
  /** The camera preset Word re-applies (e.g. "orthographicFront"). */
  camera?: string;
}

/** The structured effect set of one run. */
export interface TextEffects {
  outline?: TextEffectOutline;
  shadow?: TextEffectShadow;
  glow?: TextEffectGlow;
  reflection?: TextEffectReflection;
  bevel?: TextEffectBevels;
  rotation?: TextEffectRotation;
}

const DEGREES = 60000;
const PERCENT = 100000;

/** Top-level `w14:*` element chunks of a raw-XML fragment, in document order.
 *  Attribute values in this schema never contain `<`/`>` and elements never
 *  nest themselves, so the non-greedy tag scan is exact. */
export function splitW14Elements(xml: string | null | undefined): { name: string; xml: string }[] {
  if (!xml) return [];
  const out: { name: string; xml: string }[] = [];
  const re = /<(w14:[A-Za-z0-9]+)(?:\s[^<>]*?)?(?:\/>|>[\s\S]*?<\/\1>)/g;
  for (const match of xml.matchAll(re)) {
    out.push({ name: match[1]!.slice("w14:".length), xml: match[0] });
  }
  return out;
}

/** The raw chunk of one effect family (null when the run has none). */
export function getTextEffectChunk(
  xml: string | null | undefined,
  kind: TextEffectKind,
): string | null {
  const element = TEXT_EFFECT_ELEMENT[kind];
  return splitW14Elements(xml).find((chunk) => chunk.name === element)?.xml ?? null;
}

/** Replace (or clear) one effect family's chunk, keeping every other chunk —
 *  including families this module does not model — byte-for-byte. */
export function setTextEffectChunk(
  xml: string | null | undefined,
  kind: TextEffectKind,
  chunk: string | null,
): string | null {
  const element = TEXT_EFFECT_ELEMENT[kind];
  const kept = splitW14Elements(xml)
    .filter((entry) => entry.name !== element)
    .map((entry) => entry.xml);
  if (chunk) kept.push(chunk);
  return kept.length ? kept.join("") : null;
}

// ── Parsing ────────────────────────────────────────────────────────────────

function attrOf(xml: string, name: string): string | undefined {
  const match = new RegExp(`\\bw14:${name}="([^"]*)"`).exec(xml);
  return match?.[1];
}

function numberAttr(xml: string, name: string): number | undefined {
  const raw = attrOf(xml, name);
  if (raw == null) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
}

/** A color child (`w14:srgbClr`/`w14:schemeClr`) plus its alpha transform. */
function colorOf(xml: string, schemeFallback?: string): { color?: string; opacity?: number } {
  const srgb = /\bw14:srgbClr\s+w14:val="([0-9A-Fa-f]{6})"/.exec(xml);
  const scheme = /\bw14:schemeClr\s+w14:val="([A-Za-z0-9]+)"/.exec(xml);
  const alpha = /\bw14:alpha\s+w14:val="(\d+)"/.exec(xml);
  return {
    color: srgb?.[1]?.toUpperCase() ?? scheme?.[1] ?? schemeFallback,
    ...(alpha ? { opacity: Number(alpha[1]) / PERCENT } : {}),
  };
}

function parseBevel(xml: string): TextEffectBevel | undefined {
  const w = numberAttr(xml, "w");
  const h = numberAttr(xml, "h");
  const preset = attrOf(xml, "prst");
  if (w == null && h == null && !preset) return undefined;
  return {
    widthPx: w != null ? w / EMU_PER_PX : 0,
    heightPx: h != null ? h / EMU_PER_PX : 0,
    preset: preset ?? "circle",
  };
}

/** Parse the w14 raw XML into the structured effect set. */
export function parseTextEffects(xml: string | null | undefined): TextEffects {
  const effects: TextEffects = {};
  for (const chunk of splitW14Elements(xml)) {
    const body = chunk.xml;
    switch (chunk.name) {
      case "textOutline": {
        const width = numberAttr(body, "w");
        const { color } = colorOf(body);
        const dash = /\bw14:prstDash\s+w14:val="([A-Za-z0-9]+)"/.exec(body)?.[1];
        effects.outline = {
          widthPx: width != null ? width / EMU_PER_PX : 1,
          ...(color ? { color } : {}),
          ...(/\bw14:noFill\b/.test(body) ? { hollow: true } : {}),
          ...(dash ? { dash } : {}),
        };
        break;
      }
      case "shadow": {
        const { color, opacity } = colorOf(body, "000000");
        const dirDeg = (numberAttr(body, "dir") ?? 0) / DEGREES;
        const distPx = (numberAttr(body, "dist") ?? 0) / EMU_PER_PX;
        effects.shadow = {
          color: color ?? "000000",
          opacity: opacity ?? 1,
          blurPx: (numberAttr(body, "blurRad") ?? 0) / EMU_PER_PX,
          distPx,
          dirDeg,
        };
        break;
      }
      case "glow": {
        const { color, opacity } = colorOf(body, "4A90E2");
        effects.glow = {
          color: color ?? "4A90E2",
          opacity: opacity ?? 1,
          radiusPx: (numberAttr(body, "rad") ?? 0) / EMU_PER_PX,
        };
        break;
      }
      case "reflection": {
        effects.reflection = {
          blurPx: (numberAttr(body, "blurRad") ?? 0) / EMU_PER_PX,
          startOpacity: (numberAttr(body, "stA") ?? 100000) / PERCENT,
          endOpacity: (numberAttr(body, "endA") ?? 0) / PERCENT,
          distPx: (numberAttr(body, "dist") ?? 0) / EMU_PER_PX,
        };
        break;
      }
      case "props3d": {
        const top = /<w14:bevelT\b[^<>]*\/?>/.exec(body)?.[0];
        const bottom = /<w14:bevelB\b[^<>]*\/?>/.exec(body)?.[0];
        const bevel: TextEffectBevels = {
          ...(top ? { top: parseBevel(top) } : {}),
          ...(bottom ? { bottom: parseBevel(bottom) } : {}),
        };
        if (bevel.top || bevel.bottom) effects.bevel = bevel;
        break;
      }
      case "scene3d": {
        const camera = /\bw14:camera\s+w14:prst="([A-Za-z0-9]+)"/.exec(body)?.[1];
        const lat = numberAttr(body, "lat");
        const lon = numberAttr(body, "lon");
        const rev = numberAttr(body, "rev");
        if (camera || lat != null || lon != null || rev != null) {
          effects.rotation = {
            x: (lat ?? 0) / DEGREES,
            y: (lon ?? 0) / DEGREES,
            z: (rev ?? 0) / DEGREES,
            ...(camera ? { camera } : {}),
          };
        }
        break;
      }
      default:
        break;
    }
  }
  return effects;
}

// ── Serialization ──────────────────────────────────────────────────────────

const hex = (color: string): string => color.replace(/^#/, "").toUpperCase().slice(0, 6);

const emu = (px: number): number => Math.max(0, Math.round(px * EMU_PER_PX));

/** DrawingML positive fixed angles are 0-360° clockwise from east. */
const angle = (deg: number): number => Math.round((((deg % 360) + 360) % 360) * DEGREES);

const alphaVal = (opacity: number): number =>
  Math.round(Math.min(Math.max(opacity, 0), 1) * PERCENT);

function colorXml(color: string | undefined, opacity?: number): string {
  const inner = `<w14:srgbClr w14:val="${hex(color ?? "000000")}"${
    opacity != null && opacity < 1
      ? `><w14:alpha w14:val="${alphaVal(opacity)}"/></w14:srgbClr`
      : "/"
  }>`;
  return inner;
}

/** Serialize one structured effect into its Word-shaped w14 chunk. */
export function serializeTextEffect(kind: TextEffectKind, effect: TextEffects): string | null {
  switch (kind) {
    case "outline": {
      const outline = effect.outline;
      if (!outline) return null;
      const fill = outline.hollow
        ? "<w14:noFill/>"
        : `<w14:solidFill>${colorXml(outline.color)}</w14:solidFill>`;
      return (
        `<w14:textOutline w14:w="${emu(outline.widthPx)}" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr">` +
        fill +
        `<w14:prstDash w14:val="${outline.dash ?? "solid"}"/><w14:round/></w14:textOutline>`
      );
    }
    case "shadow": {
      const shadow = effect.shadow;
      if (!shadow) return null;
      return (
        `<w14:shadow w14:blurRad="${emu(shadow.blurPx)}" w14:dist="${emu(shadow.distPx)}" ` +
        `w14:dir="${angle(shadow.dirDeg)}" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl">` +
        `<w14:srgbClr w14:val="${hex(shadow.color)}"><w14:alpha w14:val="${alphaVal(shadow.opacity)}"/></w14:srgbClr>` +
        "</w14:shadow>"
      );
    }
    case "glow": {
      const glow = effect.glow;
      if (!glow) return null;
      return (
        `<w14:glow w14:rad="${emu(glow.radiusPx)}">` +
        `<w14:srgbClr w14:val="${hex(glow.color)}"><w14:alpha w14:val="${alphaVal(glow.opacity)}"/></w14:srgbClr>` +
        "</w14:glow>"
      );
    }
    case "reflection": {
      const reflection = effect.reflection;
      if (!reflection) return null;
      return (
        `<w14:reflection w14:blurRad="${emu(reflection.blurPx)}" w14:stA="${alphaVal(reflection.startOpacity)}" ` +
        `w14:stPos="0" w14:endA="${alphaVal(reflection.endOpacity)}" w14:endPos="55000" ` +
        `w14:dist="${emu(reflection.distPx)}" w14:dir="5400000" w14:fadeDir="5400000" ` +
        `w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="bl"/>`
      );
    }
    case "bevel": {
      const bevel = effect.bevel;
      if (!bevel?.top && !bevel?.bottom) return null;
      const edge = (tag: "bevelT" | "bevelB", value: TextEffectBevel): string =>
        `<w14:${tag} w14:w="${emu(value.widthPx)}" w14:h="${emu(value.heightPx)}" w14:prst="${value.preset}"/>`;
      return (
        `<w14:props3d w14:prstMaterial="plastic">` +
        (bevel.top ? edge("bevelT", bevel.top) : "") +
        (bevel.bottom ? edge("bevelB", bevel.bottom) : "") +
        "</w14:props3d>"
      );
    }
    case "rotation": {
      const rotation = effect.rotation;
      if (!rotation) return null;
      return (
        `<w14:scene3d><w14:camera w14:prst="${rotation.camera ?? "orthographicFront"}"/>` +
        `<w14:lightRig w14:rig="threePt" w14:dir="t"><w14:rot w14:lat="${angle(rotation.x)}" ` +
        `w14:lon="${angle(rotation.y)}" w14:rev="${angle(rotation.z)}"/></w14:lightRig></w14:scene3d>`
      );
    }
    case "fill":
      // Text fill presets (gradient fills) have no structured model yet; the
      // chunk-only API leaves them untouched.
      return null;
  }
}

/** Apply (or clear with `effect === null`) one effect family on a run's raw
 *  XML, preserving every other w14 chunk verbatim. */
export function applyTextEffect(
  xml: string | null | undefined,
  kind: TextEffectKind,
  effect: TextEffects | null,
): string | null {
  const chunk = effect ? serializeTextEffect(kind, effect) : null;
  return setTextEffectChunk(xml, kind, chunk);
}
