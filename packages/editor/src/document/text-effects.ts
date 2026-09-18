import {
  applyTextEffect,
  serializeTextEffect,
  type TextEffectKind,
  type TextEffects,
} from "@docen/docx";

/**
 * Word's Text Effects and Typography presets (Home → Font → Text Effects):
 * the gallery's named combinations plus the per-effect clear entries. The
 * `value` tokens are `<kind>:<preset>` (`outline:1`, `rotation:3`) and the
 * `:0` token clears that family; `:options` opens the custom dialog. Each
 * preset serializes through the shared docx text-effects module, so applied
 * effects round-trip exactly as Word's own w14 XML.
 */

/** The preset table — `kind:token` → the structured effect to apply. */
export const TEXT_EFFECT_PRESETS: Readonly<Record<string, TextEffects>> = {
  // Outline — Word's colored outline gallery (1 pt = 4/3 px).
  "outline:1": { outline: { color: "000000", widthPx: 4 / 3 } },
  "outline:2": { outline: { color: "2E75B5", widthPx: 4 / 3 } },
  "outline:3": { outline: { color: "C00000", widthPx: 4 / 3 } },
  "outline:4": { outline: { color: "FFFFFF", widthPx: 4 / 3 } },
  "outline:5": { outline: { color: "BF8F00", widthPx: 2 } },
  // Shadow — the offset presets.
  "shadow:1": { shadow: { color: "000000", opacity: 0.4, blurPx: 2, distPx: 3, dirDeg: 45 } },
  "shadow:2": { shadow: { color: "000000", opacity: 0.4, blurPx: 2, distPx: 3, dirDeg: 225 } },
  "shadow:3": { shadow: { color: "000000", opacity: 0.4, blurPx: 2, distPx: 3, dirDeg: 90 } },
  "shadow:4": { shadow: { color: "000000", opacity: 0.4, blurPx: 2, distPx: 3, dirDeg: 0 } },
  // Reflection.
  "reflection:1": {
    reflection: { blurPx: 0.5, startOpacity: 0.5, endOpacity: 0.1, distPx: 0 },
  },
  "reflection:2": {
    reflection: { blurPx: 1, startOpacity: 0.35, endOpacity: 0.05, distPx: 1 },
  },
  "reflection:3": {
    reflection: { blurPx: 2, startOpacity: 0.25, endOpacity: 0, distPx: 2 },
  },
  // Glow — the theme color halos.
  "glow:1": { glow: { color: "4472C4", opacity: 0.6, radiusPx: 6.67 } },
  "glow:2": { glow: { color: "FF0000", opacity: 0.6, radiusPx: 6.67 } },
  "glow:3": { glow: { color: "00B050", opacity: 0.6, radiusPx: 6.67 } },
  "glow:4": { glow: { color: "FFC000", opacity: 0.6, radiusPx: 6.67 } },
  // 3-D Format bevels.
  "bevel:1": { bevel: { top: { widthPx: 5.3, heightPx: 5.3, preset: "circle" } } },
  "bevel:2": { bevel: { top: { widthPx: 5.3, heightPx: 5.3, preset: "relaxedInset" } } },
  "bevel:3": { bevel: { top: { widthPx: 5.3, heightPx: 5.3, preset: "cross" } } },
  "bevel:4": { bevel: { top: { widthPx: 5.3, heightPx: 5.3, preset: "coolSlant" } } },
  "bevel:5": { bevel: { top: { widthPx: 5.3, heightPx: 5.3, preset: "angle" } } },
  "bevel:6": { bevel: { top: { widthPx: 5.3, heightPx: 5.3, preset: "softRound" } } },
  // 3-D Rotation.
  "rotation:1": { rotation: { x: 0, y: 0, z: 90, camera: "perspectiveFront" } },
  "rotation:2": { rotation: { x: 0, y: 0, z: 270, camera: "perspectiveFront" } },
  "rotation:3": { rotation: { x: 0, y: 0, z: 180, camera: "perspectiveFront" } },
  "rotation:4": { rotation: { x: 20, y: 340, z: 0, camera: "obliqueTopLeft" } },
  "rotation:5": { rotation: { x: 30, y: 315, z: 0, camera: "isometricTopUp" } },
  "rotation:6": { rotation: { x: 330, y: 30, z: 0, camera: "obliqueTopRight" } },
};

/** The effect families the dialog may patch. */
const KINDS: ReadonlySet<string> = new Set([
  "outline",
  "shadow",
  "glow",
  "reflection",
  "bevel",
  "rotation",
]);

/** The XML chunk for one gallery value; undefined for unknown tokens. */
export function presetXml(value: string): string | null | undefined {
  const preset = TEXT_EFFECT_PRESETS[value];
  if (!preset) return undefined;
  const kind = value.slice(0, value.indexOf(":")) as TextEffectKind;
  if (!KINDS.has(kind)) return undefined;
  return serializeTextEffect(kind, preset);
}

/** Apply (or clear, at `:0`) one gallery value onto a run's raw XML. Null when
 *  the value is unknown — callers decline instead of clearing. */
export function applyPreset(
  current: string | null | undefined,
  value: string,
): string | null | undefined {
  const separator = value.indexOf(":");
  if (separator <= 0) return undefined;
  const kind = value.slice(0, separator) as TextEffectKind;
  if (!KINDS.has(kind)) return undefined;
  const token = value.slice(separator + 1);
  if (token === "0") return applyTextEffect(current, kind, null);
  const preset = TEXT_EFFECT_PRESETS[value];
  if (!preset) return undefined;
  return applyTextEffect(current, kind, preset);
}

/** Apply the custom dialog's patch (one entry per family; null clears) onto a
 *  run's raw XML — the families not named stay untouched. */
export function applyEffectsPatch(
  current: string | null | undefined,
  patch: Record<string, TextEffects[keyof TextEffects] | null | undefined>,
): string | null {
  let xml = current ?? null;
  for (const [kind, effect] of Object.entries(patch)) {
    if (!KINDS.has(kind)) continue;
    if (effect == null) {
      xml = applyTextEffect(xml, kind as TextEffectKind, null);
    } else {
      xml = applyTextEffect(xml, kind as TextEffectKind, { [kind]: effect } as TextEffects);
    }
  }
  return xml;
}

/**
 * Design → Document Formatting → Text Effects (the document-wide theme): one
 * effect stamped onto the Title / Heading 1-3 styles (Word applies a text
 * effect theme to the heading styles, not to the body text). `"none"` clears
 * the heading effect; an unknown id returns undefined (declined).
 */
export const TEXT_EFFECT_THEMES: Readonly<
  Record<string, { kind: TextEffectKind; effects: TextEffects }>
> = {
  outline: {
    kind: "outline",
    effects: { outline: { color: "2E75B5", widthPx: 1.33 } },
  },
  shadow: {
    kind: "shadow",
    effects: { shadow: { color: "000000", opacity: 0.35, blurPx: 3, distPx: 3, dirDeg: 45 } },
  },
  reflection: {
    kind: "reflection",
    effects: { reflection: { blurPx: 1, startOpacity: 0.35, endOpacity: 0.05, distPx: 1 } },
  },
  glow: {
    kind: "glow",
    effects: { glow: { color: "4472C4", opacity: 0.6, radiusPx: 6.67 } },
  },
  bevel: {
    kind: "bevel",
    effects: { bevel: { top: { widthPx: 5.3, heightPx: 5.3, preset: "circle" } } },
  },
  rotation: {
    kind: "rotation",
    effects: { rotation: { x: 0, y: 0, z: 0, camera: "orthographicFront" } },
  },
};

/** The raw w14 chunk for a document text-effect theme; null clears, undefined
 *  declines an unknown id. */
export function textEffectThemeXml(value?: string): string | null | undefined {
  if (!value || value === "none") return null;
  const theme = TEXT_EFFECT_THEMES[value];
  if (!theme) return undefined;
  return serializeTextEffect(theme.kind, theme.effects);
}
