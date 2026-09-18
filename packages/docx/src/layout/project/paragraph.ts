// Paragraph projection: one LayoutParagraph out of the full style cascade
// (direct → style chain → docDefaults) — spacing, indent, tab stops,
// borders, shading, the numbering marker, and the inline run walk.

import {
  ptToPx,
  twipToPx,
  type LayoutBalloonAnchor,
  type LayoutInline,
  type LayoutLineHeight,
  type LayoutParagraph,
  type LayoutParagraphBorderEdge,
  type LayoutSpacing,
  type LayoutTabStop,
  type LayoutTextStyle,
} from "@docen/layout";

import type { ProjectContext } from "./context";
import { markStateful } from "./context";
import { projectDrawings } from "./drawing";
import {
  childRunsOf,
  eighthPtToPx,
  isRecord,
  measureTwip,
  num,
  str,
  type BodyParagraph,
  type Rec,
} from "./guards";
import { BUILTIN_BULLET_LEVEL, formatListNumber } from "./numbering";
import { formatIndicatorOf, projectRuns, balloonKinds } from "./runs";
import {
  alignOf,
  docDefaultsOf,
  fontAttr,
  normalizeScalePct,
  pick,
  runStyleOf,
  styleChainOf,
  toFamily,
  type FontAttr,
} from "./styles";

// ── paragraph projection ──

function toLineHeight(line: number | undefined, rule: unknown): LayoutLineHeight | undefined {
  if (line == null) return undefined;
  if (rule === "exact") return { rule: "exact", px: twipToPx(line) };
  if (rule === "atLeast") return { rule: "atLeast", px: twipToPx(line) };
  return { rule: "multiple", factor: line / 240 };
}

export function projectParagraph(p: BodyParagraph, ctx: ProjectContext): LayoutParagraph {
  const pPr: Rec = isRecord(p) ? p : {};
  const styleId = str(pPr.style) ?? str(pPr.heading);
  const chain = styleChainOf(ctx.styles, styleId);
  const chainPPr: Rec = chain.paragraph;
  const chainRPr: Rec = chain.run;
  const docDefaults = docDefaultsOf(ctx.styles);
  const docPPr: Rec = isRecord(docDefaults.paragraph) ? docDefaults.paragraph : {};
  const docRPr: Rec = isRecord(docDefaults.run) ? docDefaults.run : {};
  const cellDefaults = ctx.tableCellDefaults;
  const cellPPr: Rec = isRecord(cellDefaults?.paragraph) ? cellDefaults.paragraph : {};
  const cellRPr: Rec = isRecord(cellDefaults?.run) ? cellDefaults.run : {};

  // ¶-mark strut: direct rPr, else style chain run over docDefaults.
  const markRun: Rec = isRecord(pPr.run) ? pPr.run : {};
  const markSize = num(markRun.size);
  const isNoteStyle = styleId === "FootnoteText" || styleId === "EndnoteText";
  const chainSizeRaw = num(chainRPr.size);
  const chainSize = isNoteStyle && chainSizeRaw === 20 ? 10 : chainSizeRaw;
  const markSizePt =
    markSize ?? chainSize ?? num(cellRPr.size) ?? num(docRPr.size) ?? (isNoteStyle ? 10 : 12);
  const defFont: FontAttr =
    fontAttr(chainRPr.font) ?? fontAttr(cellRPr.font) ?? fontAttr(docRPr.font) ?? null;
  // The paragraph's default run style — every field cascades from the style
  // chain over tableCellDefaults and docDefaults (Word's effective-rPr resolution),
  // so a style's color/underline/strikethrough reach runs that carry no rPr of their own.
  const chainDefRun = runStyleOf({ ...docRPr, ...cellRPr, ...chainRPr });
  // The document theme's font pair supplies the family for text with no
  // explicit font (body → minor, heading styles → major). This is the minimal
  // theme-font consumer; a full `w:rFonts w:*Theme` resolution against a
  // parsed theme1.xml is the follow-up.
  const themeFamily = styleId?.startsWith("Heading")
    ? ctx.themeFonts?.majorFont
    : ctx.themeFonts?.minorFont;
  // The style chain's character effects, resolved once into the paragraph's
  // default run style — a run with no own rPr inherits them per field (the
  // same cascade rule bold/italic already follow). w:caps wins over
  // w:smallCaps; an explicit false on either leaves no token.
  // w:bidi drives the paragraph direction; runs inherit it unless they carry
  // their own w:rtl (direct beats inherited).
  const bidi = pPr.bidirectional === true || chainPPr.bidirectional === true;
  const defaultTextStyle: LayoutTextStyle = {
    family: toFamily(null, defFont) ?? themeFamily ?? {},
    sizePx: ptToPx(markSizePt),
    ...(bidi ? { direction: "rtl" as const } : {}),
    bold: chainDefRun.bold,
    italic: chainDefRun.italic,
    color: chainDefRun.color,
    underline: chainDefRun.underline,
    strikethrough: chainDefRun.strikethrough,
    letterSpacingPx:
      chainDefRun.characterSpacingTw != null ? twipToPx(chainDefRun.characterSpacingTw) : undefined,
    caps:
      chainDefRun.allCaps === true ? "all" : chainDefRun.smallCaps === true ? "small" : undefined,
    scalePct: normalizeScalePct(chainDefRun.scalePct),
    baselineShiftPx: chainDefRun.positionPt != null ? -ptToPx(chainDefRun.positionPt) : undefined,
    hidden: chainDefRun.vanish,
    kernPt: chainDefRun.kernPt,
    border: chainDefRun.border,
    emphasisMark: chainDefRun.emphasisMark,
  };

  // Spacing/indent cascade: direct attr wins per-field, else chain, else cellDefaults, else docDefaults.
  const direct: Rec = isRecord(pPr.spacing) ? pPr.spacing : {};
  const styleSp: Rec = isRecord(chainPPr.spacing) ? chainPPr.spacing : {};
  const cellSp: Rec = isRecord(cellPPr.spacing) ? cellPPr.spacing : {};
  const docSp: Rec = isRecord(docPPr.spacing) ? docPPr.spacing : {};
  // Word's *Lines spacing unit (hundredths of a line) beats its twip twin and
  // resolves against one line — the grid pitch on a gridded page, else the
  // single-spaced line height. The pitch lives in the section projection, so
  // the approximation here rides the engine's empirical single-line factor
  // (the DengXian word ratio) over the paragraph's default run size.
  const spacingPx = (linesKey: string, twipKey: string): number => {
    const lines = num(pick([direct, styleSp, cellSp, docSp], linesKey));
    if (lines != null) return (lines / 100) * defaultTextStyle.sizePx * 1.4;
    return twipToPx(measureTwip(pick([direct, styleSp, cellSp, docSp], twipKey)) ?? 0);
  };
  const spacing: LayoutSpacing = {
    beforePx: spacingPx("beforeLines", "before"),
    afterPx: spacingPx("afterLines", "after"),
    lineHeight: toLineHeight(
      measureTwip(pick([direct, styleSp, cellSp, docSp], "line")),
      pick([direct, styleSp, cellSp, docSp], "lineRule"),
    ),
  };

  const dInd: Rec = isRecord(pPr.indent) ? pPr.indent : {};
  const sInd: Rec = isRecord(chainPPr.indent) ? chainPPr.indent : {};
  const cellInd: Rec = isRecord(cellPPr.indent) ? cellPPr.indent : {};
  const docInd: Rec = isRecord(docPPr.indent) ? docPPr.indent : {};
  const ind = (key: string): unknown => pick([dInd, sInd, cellInd, docInd], key);

  // Numbering: the paragraph's own numPr wins, else the style chain's. The
  // level's indent fills gaps the paragraph left unset (Word: direct w:ind
  // overrides w:lvl's); the level's marker (bullet glyph or its live counter)
  // prepends + a tab hop to the body-text start. A `bullet {level}` paragraph
  // (no numbering definition) resolves against the built-in bullet table.
  const numRef: Rec | null = isRecord(pPr.numbering)
    ? pPr.numbering
    : isRecord(chainPPr.numbering)
      ? chainPPr.numbering
      : null;
  const numReference = numRef ? str(numRef.reference) : undefined;
  const numLevelIndex = num(numRef?.level) ?? 0;
  const levels = numReference ? ctx.numberings.get(numReference) : undefined;
  const bulletLevel = isRecord(pPr.bullet) ? (num(pPr.bullet.level) ?? 0) : undefined;
  const level =
    levels?.[numLevelIndex] ??
    (bulletLevel != null && !numRef ? BUILTIN_BULLET_LEVEL(bulletLevel) : undefined);

  // Indent cascade: direct w:ind > the numbering level's w:ind > style chain
  // > docDefaults. The level beating the style is Word's rule — applying a
  // list re-indents styled paragraphs (ListParagraph's 720tw must not pin
  // every level to level 0's indent). Char-unit attributes beat their twip
  // twins, w:start/w:end are the modern names of w:left/w:right (Word reads
  // each pair as one slot), and w:hanging is firstLine's negative twin
  // winning the pair within a tier — the engine takes it as a negative first
  // line (line 0 starts LEFT of the indent, where a list marker sits).
  const charsPx = (v: unknown): number | undefined => {
    const n = num(v);
    return n != null && n > 0 ? (n / 100) * defaultTextStyle.sizePx : undefined;
  };
  const twPx = (v: unknown): number | undefined => {
    const t = measureTwip(v);
    return t != null ? twipToPx(t) : undefined;
  };
  const leftPx =
    charsPx(dInd.leftChars ?? dInd.startChars) ??
    twPx(dInd.left ?? dInd.start) ??
    (level?.leftTw != null ? twipToPx(level.leftTw) : undefined) ??
    charsPx(pick([sInd, docInd], "leftChars") ?? pick([sInd, docInd], "startChars")) ??
    twPx(pick([sInd, docInd], "left") ?? pick([sInd, docInd], "start"));
  const firstLinePx = (() => {
    const directHanging = charsPx(dInd.hangingChars) ?? twPx(dInd.hanging);
    if (directHanging != null) return -directHanging;
    const directTw = twPx(dInd.firstLine);
    if (directTw != null) return directTw;
    const directChars = charsPx(dInd.firstLineChars);
    if (directChars != null) return directChars;
    if (level?.hangingTw != null && level.hangingTw > 0) return -twipToPx(level.hangingTw);
    const styleHanging =
      charsPx(pick([sInd, docInd], "hangingChars")) ?? twPx(pick([sInd, docInd], "hanging"));
    if (styleHanging != null) return -styleHanging;
    const styleTw = twPx(pick([sInd, docInd], "firstLine"));
    if (styleTw != null) return styleTw;
    return charsPx(pick([sInd, docInd], "firstLineChars"));
  })();
  const rightPx =
    charsPx(dInd.rightChars ?? dInd.endChars) ??
    twipToPx(measureTwip(ind("right") ?? ind("end")) ?? 0);
  const indent = {
    leftPx: leftPx || undefined,
    rightPx: rightPx || undefined,
    firstLinePx,
  };

  // Tab stops: twips from the content-box left edge → px from the TEXT-box
  // edge (the engine measures x from the left indent). "decimal" renders as
  // left for now; the exotic bar/clear/end kinds carry no box.
  const tabStops: LayoutTabStop[] | undefined = Array.isArray(pPr.tabStops)
    ? pPr.tabStops.flatMap((ts) => {
        if (!isRecord(ts)) return [];
        const positionPx = measureTwip(ts.position);
        if (positionPx == null) return [];
        const type =
          ts.type === "right"
            ? "right"
            : ts.type === "center"
              ? "center"
              : ts.type === "decimal"
                ? "decimal"
                : ts.type === "bar"
                  ? "bar"
                  : ("left" as const);
        const leader =
          ts.leader === "dot" ||
          ts.leader === "heavy" ||
          ts.leader === "hyphen" ||
          ts.leader === "middleDot" ||
          ts.leader === "underscore"
            ? ts.leader
            : undefined;
        return [{ positionPx: twipToPx(positionPx) - (indent.leftPx ?? 0), type, leader }];
      })
    : undefined;

  // Paragraph borders (w:pBdr): direct, else the style chain's.
  const bRec: Rec = isRecord(pPr.border)
    ? pPr.border
    : isRecord(chainPPr.border)
      ? chainPPr.border
      : {};
  const borderEdge = (v: unknown): LayoutParagraphBorderEdge | undefined => {
    if (!isRecord(v)) return undefined;
    const size = num(v.size);
    const space = num(v.space);
    return {
      style: typeof v.style === "string" ? v.style : undefined,
      px: size != null ? eighthPtToPx(size) : undefined,
      spacePx: space != null ? ptToPx(space) : undefined,
    };
  };
  const borders = {
    top: borderEdge(bRec.top),
    right: borderEdge(bRec.right),
    bottom: borderEdge(bRec.bottom),
    left: borderEdge(bRec.left),
  };
  // Paragraph shading (w:shd): direct, else the style chain's — same
  // direct-else-chain as borders; the fill gate mirrors the cell projection.
  const shdRec: Rec = isRecord(pPr.shading)
    ? pPr.shading
    : isRecord(chainPPr.shading)
      ? chainPPr.shading
      : {};
  const shadingFill =
    typeof shdRec.fill === "string" && shdRec.fill !== "auto" && shdRec.type !== "nil"
      ? shdRec.fill
      : undefined;

  // The list marker: a bullet emits its glyph; a numbered level advances its
  // counter (resetting deeper levels) and substitutes %k in w:lvlText with the
  // formatted counter of level k-1 — "%1.%2" at level 1 → "2.3".
  const markerInline: LayoutInline[] = (() => {
    if (!level) return [];
    if (level.format === "bullet") {
      return level.text
        ? [
            { kind: "text", text: level.text, style: defaultTextStyle, synthetic: true },
            { kind: "tab", toPx: 0 },
          ]
        : [];
    }
    if (level.format === "none") return [];
    if (!numReference) return [];
    markStateful(ctx);
    const counters = ctx.listCounters.get(numReference) ?? [];
    ctx.listCounters.set(numReference, counters);
    counters[numLevelIndex] = (counters[numLevelIndex] ?? 0) + 1;
    counters.length = numLevelIndex + 1;
    const marker = (level.text ?? "%1.").replace(/%([1-9])/g, (_, k: string) => {
      const idx = Number(k) - 1;
      const lvl = levels?.[idx];
      return formatListNumber(lvl?.format ?? level.format, counters[idx] ?? 1);
    });
    return [
      { kind: "text", text: marker, style: defaultTextStyle, synthetic: true },
      { kind: "tab", toPx: 0 },
    ];
  })();

  const runs: readonly unknown[] = childRunsOf(p);
  const drawings = projectDrawings(runs, ctx);
  // Margin-balloon anchors: run-level ones come from the inline walk,
  // paragraph-level ones from a pPrChange on this node.
  const anchors: LayoutBalloonAnchor[] = [];
  const suppressAutoHyphens =
    pPr.suppressAutoHyphens === true || chainPPr.suppressAutoHyphens === true;
  const runCtx: ProjectContext =
    suppressAutoHyphens !== ctx.suppressAutoHyphens ? { ...ctx, suppressAutoHyphens } : ctx;
  const inline = projectRuns(runs, chainRPr, docRPr, defaultTextStyle, runCtx, anchors);
  // A tracked pPr change (w:pPrChange) marks the paragraph for the painter's
  // change bar — author-colored, or neutral in "By change type" mode.
  const revision = isRecord(pPr.revision) ? pPr.revision : undefined;
  const indicator = revision ? formatIndicatorOf(ctx, revision) : {};
  const change = (indicator as { formatChange?: { color: string } }).formatChange;
  if (revision && change && balloonKinds(ctx).revisions) {
    anchors.push({
      id: num(revision.id) ?? 0,
      kind: "revision",
      color: change.color,
      label: str(revision.author) ?? "",
      inlineIndex: -1,
    });
  }
  const textDir = str(pPr.textDirection ?? chainPPr.textDirection);
  const textDirection =
    textDir === "tbRl" || textDir === "btLr" || textDir === "lrTb" ? textDir : undefined;
  const kinsokuRaw = pick([pPr, chainPPr], "kinsoku");
  const overflowRaw = pick([pPr, chainPPr], "overflowPunctuation");
  const charSpacing = pick([pPr, chainPPr], "characterSpacingControl");
  const wordWrapRaw = pick([pPr, chainPPr], "wordWrap");
  const autoSpaceRaw = pick([pPr, chainPPr], "autoSpaceDE");
  const textAlignment = (str(pick([pPr, chainPPr], "textAlignment")) ?? undefined) as
    | "auto"
    | "baseline"
    | "bottom"
    | "center"
    | "top"
    | undefined;

  const frame = (pPr.framePr as Rec) ?? (pPr.frame as Rec);
  const dropCapVal = str((pPr.dropCap as Rec)?.val ?? pPr.dropCap ?? frame?.dropCap);
  const dropCapLines = num((pPr.dropCap as Rec)?.lines ?? frame?.lines) ?? 3;
  // office-open parses w:hSpace/w:vSpace into frame.space.{horizontal,vertical};
  // older/legacy mirrors may carry hSpace directly.
  const frameSpace = isRecord(frame?.space) ? frame.space : undefined;
  const dropCapDist = num(
    (pPr.dropCap as Rec)?.distance ?? frameSpace?.horizontal ?? frame?.hSpace,
  );
  const dropCap =
    dropCapVal === "drop" || dropCapVal === "dropped"
      ? {
          type: "dropped" as const,
          lines: dropCapLines,
          distancePx: dropCapDist ? twipToPx(dropCapDist) : undefined,
        }
      : dropCapVal === "margin"
        ? {
            type: "margin" as const,
            lines: dropCapLines,
            distancePx: dropCapDist ? twipToPx(dropCapDist) : undefined,
          }
        : undefined;

  return {
    kind: "paragraph",
    inline: markerInline.length ? markerInline.concat(inline) : inline,
    drawings: drawings.length > 0 ? drawings : undefined,
    ...(dropCap ? { dropCap } : {}),
    spacing,
    indent,
    tabStops: tabStops && tabStops.length > 0 ? tabStops : undefined,
    defaultTabStopPx: ctx.defaultTabStopPx,
    borders: borders.top || borders.right || borders.bottom || borders.left ? borders : undefined,
    shadingFill,
    markSizePx: markSize != null ? ptToPx(markSize) : undefined,
    defaultTextStyle,
    snapToGrid: typeof pPr.snapToGrid === "boolean" ? pPr.snapToGrid : null,
    align: alignOf(pick([pPr, chainPPr, cellPPr, docPPr], "alignment")),
    keepLines: pPr.keepLines === true || chainPPr.keepLines === true,
    keepNext: pPr.keepNext === true || chainPPr.keepNext === true,
    widowControl: pick([pPr, chainPPr], "widowControl") !== false,
    pageBreakBefore: pPr.pageBreakBefore === true || chainPPr.pageBreakBefore === true,
    suppressLineNumbers: pPr.suppressLineNumbers === true || chainPPr.suppressLineNumbers === true,
    bidi: bidi || undefined,
    textDirection,
    ...(typeof kinsokuRaw === "boolean" ? { kinsoku: kinsokuRaw } : {}),
    ...(typeof overflowRaw === "boolean" ? { overflowPunctuation: overflowRaw } : {}),
    ...(charSpacing === "dontCompress"
      ? { compressPunctuation: false }
      : charSpacing != null
        ? { compressPunctuation: true }
        : {}),
    ...(typeof wordWrapRaw === "boolean" ? { wordWrap: wordWrapRaw } : {}),
    ...(typeof autoSpaceRaw === "boolean" ? { autoSpaceDE: autoSpaceRaw } : {}),
    textAlignment,
    ...indicator,
    ...(anchors.length > 0 ? { balloons: anchors } : {}),
  };
}
