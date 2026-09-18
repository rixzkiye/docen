import type { ParagraphChild } from "@office-open/docx";

import { Node } from "../core";
import { convertLinearToOMML as convertLinear, mathInputToLatex } from "./math-tex";
import type { ParseInlineRule } from "./types";
import { attrNative } from "./utils";

/**
 * Math representation: `MathInline` (`mathInline`) is THE single math node.
 * DOCX `w:oMath`/`w:oMathPara` content resolves into it (the `math` attr keeps
 * the office-open MathInput verbatim, `linear` is the editable LaTeX-like
 * projection), compile emits the MathInput back into the office-open model,
 * and the Insert → Equation path seeds the same node. `convertLinearToOMML` /
 * `convertOMMLToLinear` back the linear ↔ structured conversion (HTML
 * clipboard spans and authoring labels); the parser/serializer pair lives in
 * math-tex.ts and covers fractions, roots, scripts, Greek, operators, n-ary
 * limits, functions, delimiters, matrices and aligned environments.
 */

/** DOCX `{ math: MathInput }` run branch → the mathInline atom. The run-level
 *  `math` boolean flag (rPr) is NOT this branch — the discriminator is the
 *  object shape. */
type MathDocxBranch = Extract<ParagraphChild, { math: object }>;

export const parseDocxInline: ParseInlineRule<MathDocxBranch> = {
  match: (child): child is MathDocxBranch => {
    const math = (child as { math?: unknown }).math;
    return typeof math === "object" && math !== null;
  },
  convert: (child) => ({
    type: "mathInline",
    attrs: {
      math: child.math,
      linear: convertOMMLToLinear(child.math),
    },
  }),
};

/**
 * Convert linear math text (`\frac{a}{b}`, `\sqrt[3]{x}`, `x^2`, `x_i`,
 * `\sum_{i=1}^{n}`, `\alpha`, `\begin{pmatrix}…\end{pmatrix}`, …) into
 * structured OMML MathInput. The parser covers the LaTeX subset listed in
 * math-tex.ts; unknown commands degrade to their literal text, never to an
 * empty structure.
 */
export function convertLinearToOMML(linear: string): Record<string, unknown> {
  return convertLinear(linear);
}

/**
 * Convert structured OMML MathInput back to linear representation — the
 * inverse of {@link convertLinearToOMML} for every structure it emits, plus
 * the office-open shapes a parsed DOCX carries (box/phant/borderBox, `{text,
 * properties}` runs, pre/sub/superscripts).
 */
export function convertOMMLToLinear(math: unknown): string {
  return mathInputToLatex(math);
}

export const MathInline = Node.create({
  name: "mathInline",
  group: "inline",
  inline: true,
  atom: true,

  parseDocxInline,

  addAttributes() {
    return {
      math: attrNative(),
      linear: { default: "" },
    };
  },

  parseHTML() {
    return [
      {
        tag: "span.docx-math",
        getAttrs: (element) => {
          const el = element as HTMLElement;
          const linear = el.getAttribute("data-linear") ?? el.textContent ?? "";
          const converted = convertLinearToOMML(linear);
          // The math attr is the office-open paragraph shape
          // ({ children }); a single-construct result is wrapped so paste
          // never produces an empty oMath.
          return {
            linear,
            math: Array.isArray(converted.children) ? converted : { children: [converted] },
          };
        },
      },
    ];
  },

  renderHTML({ HTMLAttributes }) {
    return [
      "span",
      {
        class: "docx-math",
        "data-linear": HTMLAttributes.linear ?? "",
      },
      HTMLAttributes.linear ?? "",
    ];
  },
});
