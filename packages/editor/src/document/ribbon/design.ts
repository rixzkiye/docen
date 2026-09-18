import type { RibbonTab } from "../../ui";
import { btn, cmd, col, grid, group, menu, opt, parsedItems, split, tabNode } from "./shared";

/** The Page Borders split's presets — Word's Borders and Shading gallery in
 *  menu form: none, a plain box, Word's shadow box (thick bottom/right), a
 *  double rule, and a dashed rule. */
export const pageBorderItems = (): string =>
  JSON.stringify([
    { text: opt("page-border-none"), value: "none" },
    { text: opt("page-border-box"), value: "box" },
    { text: opt("page-border-shadow"), value: "shadow" },
    { text: opt("page-border-double"), value: "double" },
    { text: opt("page-border-dashed"), value: "dashed" },
    { text: opt("borders-shading"), value: "borders-shading" },
  ]);

/** The Design tab's style-set gallery (Word's Document Formatting group): the
 *  document's opening model ("default") plus three font-family presets. The
 *  presets stamp through the style-set command; "default" restores from the
 *  host's open snapshot. */
export const styleSetItems = (): string =>
  JSON.stringify([
    { text: opt("style-set-default"), value: "default" },
    { text: opt("style-set-modern"), value: "modern" },
    { text: opt("style-set-classic"), value: "classic" },
    { text: opt("style-set-elegant"), value: "elegant" },
  ]);

/** The Design tab's Text Effects gallery — the document-wide effect theme
 *  stamped onto the Title/Heading styles (Word's Document Formatting → Text
 *  Effects). Labels reuse the Font gallery's effect names. */
export const textEffectThemeItems = (): string =>
  JSON.stringify([
    { text: opt("te-none"), value: "none" },
    { text: opt("te-outline"), value: "outline" },
    { text: opt("te-shadow"), value: "shadow" },
    { text: opt("te-reflection"), value: "reflection" },
    { text: opt("te-glow"), value: "glow" },
    { text: opt("te-bevel"), value: "bevel" },
    { text: opt("te-rotation"), value: "rotation" },
  ]);

/** The Watermark split's drop-down: Word's preset gallery plus Remove. */
export const watermarkItems = (): string =>
  JSON.stringify([
    { text: opt("watermark-confidential-1"), value: "confidential-1" },
    { text: opt("watermark-confidential-2"), value: "confidential-2" },
    { text: opt("watermark-confidential-3"), value: "confidential-3" },
    { text: opt("watermark-urgent"), value: "urgent" },
    { text: opt("watermark-asap"), value: "asap" },
    { text: opt("watermark-draft"), value: "draft" },
    { text: opt("watermark-sample"), value: "sample" },
    { text: opt("watermark-remove"), value: "remove" },
    { text: opt("watermark-custom"), value: "custom" },
  ]);

/** The Paragraph Spacing menu (Word's Design → Paragraph Spacing): document-
 *  default spacing presets stamped onto the styles' docDefaults — Word's
 *  "default" restores its factory 8pt-after / 1.08-line spacing. */
export const paragraphSpacingItems = (): string =>
  JSON.stringify([
    { text: opt("paragraph-spacing-default"), value: "default" },
    { text: opt("paragraph-spacing-none"), value: "none" },
    { text: opt("paragraph-spacing-compact"), value: "compact" },
    { text: opt("narrow"), value: "narrow" },
    { text: opt("wide"), value: "wide" },
  ]);

export const themeItems = (): string =>
  JSON.stringify([
    { text: "Office", value: "office" },
    { text: "Facet", value: "facet" },
    { text: "Integral", value: "integral" },
    { text: "Ion", value: "ion" },
    { text: "Organic", value: "organic" },
    { text: "Retrospect", value: "retrospect" },
    { text: "Slice", value: "slice" },
    { text: "Wisp", value: "wisp" },
  ]);

export const themeColorItems = (): string =>
  JSON.stringify([
    { text: "Office", value: "office" },
    { text: "Facet", value: "facet" },
    { text: "Integral", value: "integral" },
    { text: "Ion", value: "ion" },
    { text: "Organic", value: "organic" },
    { text: "Retrospect", value: "retrospect" },
    { text: "Slice", value: "slice" },
    { text: "Wisp", value: "wisp" },
  ]);

export const themeFontItems = (): string =>
  JSON.stringify([
    { text: "Office (Calibri / Calibri Light)", value: "office" },
    { text: "Facet (Trebuchet MS / Garamond)", value: "facet" },
    { text: "Integral (Century Gothic)", value: "integral" },
    { text: "Ion (Century Gothic)", value: "ion" },
    { text: "Organic (Calibri / Arial)", value: "organic" },
    { text: "Retrospect (Georgia / Arial)", value: "retrospect" },
    { text: "Slice (Georgia / Garamond)", value: "slice" },
    { text: "Wisp (Georgia / Segoe UI)", value: "wisp" },
  ]);

export const designTab = (): RibbonTab =>
  tabNode("design", [
    group("document-formatting", [
      // The paint-brush glyph doubles here: a style set is Word's "apply a
      // formatting theme to the document" action.
      split("format-painter", "style-set", parsedItems(styleSetItems()), { size: "large" }),
      split("theme", "theme", parsedItems(themeItems()), { size: "large" }),
      split("font-color", "theme-color", parsedItems(themeColorItems()), { size: "large" }),
      split("text-font", "theme-font", parsedItems(themeFontItems()), { size: "large" }),
      btn("text-effects", "effects", { size: "large" }),
      // Document text-effect theme (stamped onto the heading styles).
      split("text-effects", "effects", parsedItems(textEffectThemeItems()), { size: "large" }),
      col([
        grid([
          menu("line-spacing", "paragraph-spacing", parsedItems(paragraphSpacingItems())),
          btn("page-border", "set-default"),
        ]),
      ]),
    ]),
    group("page-background", [
      split("watermark", "watermark", parsedItems(watermarkItems()), { size: "large" }),
      {
        type: "color-picker",
        icon: "page-color",
        event: "page-color",
        label: cmd("page-color"),
        defaultColor: "FFFFFF",
        size: "large",
      },
      btn("fill-effects", "fill-effects", { size: "large" }),
      split("page-border", "page-border", parsedItems(pageBorderItems()), { size: "large" }),
    ]),
  ]);
