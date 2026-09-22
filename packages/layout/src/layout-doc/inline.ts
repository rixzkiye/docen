import type { FontSlots } from "../font";
import type { LayoutMathData } from "../math/math-layout";
import { formatNumber } from "../numbering-format";
import type { LayoutDrawingLine, LayoutDrawingMember, LayoutDrawingShadow } from "./drawing";

/** A run's character border (w:rPr/w:bdr) — the box Word paints around the
 *  run's text with its own padding. */
export interface LayoutCharBorder {
  /** The ST_Border token (single/double/dashed/dotted/…) — absent = single. */
  style?: string;
  /** Stroke width in px (w:sz eighths of a point resolved). */
  px?: number;
  /** Padding between the run's glyphs and the box, px (w:space resolved). */
  spacePx?: number;
  /** Stroke color, hex RRGGBB — absent = the run's ink. */
  color?: string;
}

export interface LayoutTextStyle {
  family: string | FontSlots;
  sizePx: number;
  bold?: boolean;
  italic?: boolean;
  /** Text color, hex RRGGBB (w:color; absent/auto → the renderer's ink). */
  color?: string;
  /** Underlined (w:u, any non-none pattern). */
  underline?: boolean;
  /** The w:u pattern token (ST_Underline minus "none") — absent = single. */
  underlineStyle?: string;
  /** w:u color, hex RRGGBB — absent = the text color. */
  underlineColor?: string;
  /** Struck through (w:strike / w:dstrike). */
  strikethrough?: boolean;
  /** Extra per-character spacing in px (OOXML w:spacing resolved). */
  letterSpacingPx?: number;
  /** Raised/lowered run (w:vertAlign): glyphs paint at the scaled size on a
   *  shifted baseline (Word's FootnoteReference look). Measuring applies the
   *  same scaling — see vertAlignedSizePx. */
  verticalAlign?: "superscript" | "subscript";
  /** Character highlight (w:highlight), the ST_HighlightColor token — the
   *  painter maps it to Word's highlight palette and fills the run's box
   *  beneath the glyphs. */
  highlight?: string;
  /** Character shading (w:shd) fill, #RRGGBB — the arbitrary-color highlight
   *  Word's shading button paints on selected text. A highlight (same box,
   *  Word's palette) wins when both are present, per OOXML precedence. */
  shadingFill?: string;
  /** Display transform for the run's cased glyphs (w:caps / w:smallCaps):
   *  "all" uppercases every glyph at the run's size; "small" uppercases the
   *  lowercase letters at SMALL_CAPS_SCALE of the size while the run's own
   *  capitals stay full size. The stored text never changes — only the
   *  measured and painted forms (see displayTextOf). */
  caps?: "all" | "small";
  /** Horizontal glyph scale in percent (w:w, 1–600; absent = 100) — the
   *  run's advances measure and its glyphs paint scaled by px/100. */
  scalePct?: number;
  /** Baseline offset in px from w:position (negative = raised) — the glyphs
   *  move without changing the line box, Word's raise/lower semantics. */
  baselineShiftPx?: number;
  /** Hidden formatting (w:vanish): hidden runs are not measured or painted
   *  while the host's Show Hidden Text setting is off (the projection marks
   *  the atom `suppressed`); when shown, the painter adds its dotted marker. */
  hidden?: boolean;
  /** Kerning threshold in points (w:kern / 2): kerning applies when the run's
   *  font size is at least this — see kerningActive. Absent/0 = no kerning. */
  kernPt?: number;
  /** Text direction: ltr, rtl, or auto (BiDi / script-dependent). */
  direction?: "ltr" | "rtl" | "auto";
  /** BCP 47 or OpenType language tag. */
  language?: string;
  /** ISO 15924 script tag (e.g. Latn, Arab, Hebr, Thai, Deva, Khmr, Hani). */
  script?: string;
  /** Vertical text flow (e.g. vertical CJK). */
  vertical?: boolean;
  /** Character border (w:bdr) — a box around the run's glyphs. */
  border?: LayoutCharBorder;
  /** Emphasis mark (w:em): a small mark drawn above every glyph (dot / comma
   *  / circle) or below it (underDot). */
  emphasisMark?: "dot" | "comma" | "circle" | "underDot";
  /** w:outline — character outline / hollow stroke effect. */
  outline?: boolean | { color?: string; widthPx?: number };
  /** w:shadow — drop shadow effect. */
  shadow?: boolean | { x?: number; y?: number; blur?: number; color?: string };
  /** w:emboss — raised 3D appearance. */
  emboss?: boolean;
  /** w:imprint — engraved 3D appearance. */
  imprint?: boolean;
  /** Glow halo effect (Word 2010+ / DrawingML). */
  glow?: { radiusPx?: number; color?: string };
  /** Reflection effect (Word 2010+ / DrawingML). */
  reflection?: { blur?: number; distancePx?: number; opacity?: number };
  /** 3-D Format bevels (w14:props3d bevelT/bevelB) — the painter raises the
   *  glyphs with a light top edge and a dark bottom edge. */
  bevel?: {
    top?: { widthPx?: number; heightPx?: number; preset?: string };
    bottom?: { widthPx?: number; heightPx?: number; preset?: string };
  };
  /** 3-D rotation (w14:scene3d) in degrees: x = lat, y = lon, z = rev. The
   *  painter approximates the projection (z rotates, x/y squash the run). */
  rotation3d?: { x?: number; y?: number; z?: number };
  /** OpenType ligatures setting (w:ligatures). */
  ligatures?: "none" | "standard" | "contextual" | "historical" | "discretionary" | "all";
  /** OpenType number form (w:numForm). */
  numForm?: "default" | "lining" | "oldStyle";
  /** OpenType number spacing (w:numSpacing). */
  numSpacing?: "default" | "proportional" | "tabular";
  /** OpenType stylistic set index (1-20, w:stylisticSet). */
  stylisticSet?: number;
  /** Explicit font feature settings tag -> value. */
  fontFeatures?: readonly { tag: string; value?: number }[];
  /** Explicit font variation settings tag -> value (for variable fonts). */
  fontVariations?: readonly { tag: string; value: number }[];
  /** Numeric font weight (100-900) for variable fonts. */
  fontWeight?: number;
}

/** a:srcRect crop as fractions of the image edge (0-1, each side inward);
 *  the painted region is the remainder. */
export interface LayoutPictureCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** Phonetic-guide annotation riding a text item (w:ruby — Word's 拼音指南).
 *  The annotation paints above the base glyphs at fontSizePx; the base text
 *  is the item's own text (the editor models the guide as a character mark,
 *  so this is pure paint/layout metadata). */
export interface LayoutRuby {
  /** The annotation text (w:rt, flattened). */
  text: string;
  /** The annotation's distribution over the base (ST_RubyAlign token).
   *  center/left/right shift the annotation within the base's width — the
   *  distribute variants fall to center. */
  alignment?: string;
  /** Annotation font size in px (w:hps resolved; Word's default is half the
   *  base size). */
  fontSizePx: number;
}

/** Two-lines-in-one / combined characters (w:eastAsianLayout @w:combine —
 *  Word's 双行合一 / 合并字符). The run's text packs into two half-size lines
 *  inside a normal line box, optionally wrapped in bracket glyphs. */
export interface LayoutCombine {
  /** The upper line's text. */
  first: string;
  /** The lower line's text. */
  second: string;
  /** The bracket pair drawn around both lines (ST_CombineBrackets minus
   *  "none"); the painter draws the matching glyph characters. */
  bracket?: "round" | "square" | "angle" | "curly";
}

/** One inline item of a paragraph. Text, hard breaks, tabs, and inline
 *  pictures are all line-box content — one box packer handles the four (the
 *  unified text+picture breaker the DOM route never had). A picture's `src`
 *  (data URL or object URL) is renderer-only passthrough — the engine measures
 *  the box and never loads it. A tab advances to the next stop: the explicit
 *  `toPx` (a numbering bullet's hop to the body text), the paragraph's
 *  `tabStops`, or the default grid (720 twips). */

export interface LayoutInlineNoteRef {
  kind: "footnote" | "endnote";
  id: number;
  ordinal: number;
}

export interface LayoutInlineFormField {
  name: string;
  type: "text" | "checkbox" | "dropdown";
  value?: string | boolean;
  readOnly?: boolean;
  options?: string[];
}

export type LayoutInline =
  /** A `field` marker makes the text a dynamic page-number atom (w:fldSimple /
   *  complexField PAGE / NUMPAGES): the value only exists after pagination, so
   *  `text` is a single-digit placeholder for measuring and the painter swaps
   *  in the real page number.
   *
   *  Every field atom also carries its descriptor — `instruction` verbatim
   *  plus the `result` the document cached and the `resolved` value the render
   *  pass derived from the paginated document. Paint prefers `resolved`, then
   *  `text` (the cache); the dynamic page markers keep their placeholder text
   *  so pagination never feeds back into their own measurement. */
  | {
      kind: "text";
      text: string;
      style: LayoutTextStyle;
      /** The hidden run is not displayed (Show Hidden Text off): the layout
       *  charges no advance for it and the painter draws nothing, while the
       *  atom keeps its `text` so the editor's caret lattice stays aligned
       *  with the document model (Word's hidden-text behavior). */
      suppressed?: boolean;
      field?: "page" | "numPages";
      /** The field instruction verbatim (`PAGE \* MERGEFORMAT`, `AUTHOR`,
       *  `REF _Ref123 \h`) — set on every w:fldSimple / complexField atom. */
      instruction?: string;
      /** The cached result the document stored (simpleField @w:cachedValue /
       *  complexField result). `text` mirrors it except on the dynamic page
       *  markers, whose text is the measuring placeholder. */
      result?: string;
      /** The live value the render pass resolved against the paginated
       *  document (real page/section numbers) — paint prefers it over `text`.
       *  Absent = the field was not resolved this render (cached value shows). */
      resolved?: string;
      /** Ids of the comments whose range covers this atom (w:commentRangeStart
       *  /commentRangeEnd): the painter tints the text's box (sorted, unique).
       *  Pure paint metadata — measuring and wrapping ignore it. */
      commentIds?: number[];
      /** Synthesized paint content with no document-model character behind it
       *  (a numbering bullet's glyph): the editor's caret/selection mapping
       *  must not count it against the paragraph's text positions. */
      synthetic?: boolean;
      /** Note reference metadata (Word's FootnoteReference / EndnoteReference).
       *  The paginator reads this to anchor footnotes to the page where this run lands. */
      noteRef?: LayoutInlineNoteRef;
      /** Phonetic guide (w:ruby): the annotation paints above the base glyphs
       *  and the line's natural height reserves space for it. */
      ruby?: LayoutRuby;
      /** Word field shading: atom carries a calculated field result. */
      fieldShading?: boolean;
      /** Two-lines-in-one (w:eastAsianLayout): the atom packs its whole text
       *  into two half-size lines — an unbreakable box of the combined width. */
      combine?: LayoutCombine;
      /** Tracked format change (w:rPrChange): the painter draws Word's change
       *  bar beside the line this run appears on, in the author's revision
       *  color. Pure paint metadata — measuring and wrapping ignore it. */
      formatChange?: { color: string };
      /** Hyperlink destination (external URL or internal #bookmark anchor). */
      link?: { url?: string; anchor?: string; tooltip?: string };
      /** Form field metadata (Word legacy form fields: textInput, checkBox, dropDownList). */
      formField?: LayoutInlineFormField;
    }
  | { kind: "break" }
  | { kind: "tab"; toPx?: number }
  | {
      /** An OMML formula (m:oMath): structured math layout data (with element
       *  geometries, fraction lines, surd bars, slots), or fallback label box. */
      kind: "math";
      label: string;
      widthPx: number;
      heightPx: number;
      data?: LayoutMathData;
    }
  | {
      kind: "picture";
      widthPx: number;
      heightPx: number;
      src?: string;
      /** a:srcRect crop: the flat `src` paints only the visible remainder
       *  (Leafer paints whole sources, so the renderer sub-regions it). */
      crop?: LayoutPictureCrop;
      /** Vector replay members (a WMF/EMF metafile source): when present, the
       * renderer paints these instead of loading `src` — same renderer-only
       * passthrough contract; the engine never reads beyond widthPx/heightPx. */
      members?: LayoutDrawingMember[];
      /** Clockwise spin of the box about its center, degrees (a:xfrm @rot) —
       * the extent stays put, the painted content tilts inside it. */
      rotation?: number;
      /** Mirrored content (a:xfrm @flipH/@flipV) — the glyphs of the picture
       * mirror inside the extent box, which stays put. */
      flipH?: boolean;
      flipV?: boolean;
      /** Pixel-adjustment filter, CSS filter syntax (`brightness(1.2)
       * saturate(0.4)`) — the projection mapped the picture's blip effects
       * into it; the renderer composites the source through it. */
      filter?: string;
      /** Fill opacity 0-1 (the blip alpha modulate); absent → opaque. */
      opacity?: number;
      /** The shape's outer shadow (pic:spPr a:effectLst a:outerShdw). */
      shadow?: LayoutDrawingShadow;
      /** The picture's outline stroke (pic:spPr a:ln) — Word's picture border. */
      line?: LayoutDrawingLine;
      /** The docPr name (Word's shape name) — the tagged PDF's /T. */
      title?: string;
      /** The docPr description (Word's alt text) — the tagged PDF's /Alt. */
      altText?: string;
    };

/** A text atom's painted label — the shared field-display contract the
 *  projection, the render resolve pass and the painter all read: the render
 *  pass's `resolved` value wins, then the live page-number value for the
 *  dynamic markers (furniture atoms the page flow never resolved), else the
 *  measured text (the cached result, or the instruction under Alt+F9 — code
 *  atoms carry neither `resolved` nor a marker). */
export function fieldLabelOf(
  inline: LayoutInline & { kind: "text" },
  page: {
    pageIndex: number;
    pageCount: number;
    pageNumber?: { fmt?: string; offset?: number };
  },
  measured: string,
): string {
  return (
    inline.resolved ??
    (inline.field === "page"
      ? formatNumber(page.pageNumber?.fmt, page.pageIndex + 1 + (page.pageNumber?.offset ?? 0))
      : inline.field === "numPages"
        ? String(page.pageCount)
        : measured)
  );
}

export type LaidOutInlineText = Extract<LayoutInline, { kind: "text" }>;
