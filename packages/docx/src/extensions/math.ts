import { Node } from "../core";
import { attrNative } from "./utils";

/**
 * Math Architecture Decision:
 * - `inlinePassthrough.math` (registered in coverage.ts) is canonical for DOCX/OOXML
 *   lossless roundtrip and byte fidelity.
 * - `MathInline` (`mathInline`) and convertLinearToOMML / convertOMMLToLinear provide
 *   the linear LaTeX-like representation (`\frac{a}{b}`, `\sqrt{x}`, etc.) for authoring,
 *   HTML clipboard roundtrips, and UI interaction.
 */

/**
 * Convert simple linear math text (e.g. `\frac{a}{b}`, `\sqrt{x}`, `x^2`, `x_i`)
 * into structured OMML MathInput for @office-open/docx.
 */
export function convertLinearToOMML(linear: string): Record<string, unknown> {
  const trimmed = linear.trim();
  // Fraction: \frac{num}{den}
  const fracMatch = /^\\frac\{([^}]+)\}\{([^}]+)\}$/.exec(trimmed);
  if (fracMatch) {
    return {
      fraction: {
        numerator: [fracMatch[1]],
        denominator: [fracMatch[2]],
      },
    };
  }

  // Radical: \sqrt{rad} or \sqrt[deg]{rad}
  const radMatch = /^\\sqrt(?:\[([^\]]+)\])?\{([^}]+)\}$/.exec(trimmed);
  if (radMatch) {
    return {
      radical: {
        children: [radMatch[2]],
        ...(radMatch[1] ? { degree: [radMatch[1]] } : {}),
      },
    };
  }

  // Superscript: base^exp
  const supMatch = /^([^^]+)\^\{?([^}]+)\}?$/.exec(trimmed);
  if (supMatch) {
    return {
      superScript: {
        children: [supMatch[1]],
        superScript: [supMatch[2]],
      },
    };
  }

  // Subscript: base_sub
  const subMatch = /^([^_]+)_\{?([^}]+)\}?$/.exec(trimmed);
  if (subMatch) {
    return {
      subScript: {
        children: [subMatch[1]],
        subScript: [subMatch[2]],
      },
    };
  }

  // Fallback to plain run
  return {
    run: {
      text: trimmed,
    },
  };
}

/**
 * Convert structured OMML MathInput back to linear representation.
 */
export function convertOMMLToLinear(math: unknown): string {
  if (typeof math === "string") return math;
  if (!math || typeof math !== "object") return "";

  const rec = math as Record<string, unknown>;
  if (rec.fraction && typeof rec.fraction === "object") {
    const f = rec.fraction as Record<string, unknown>;
    const num = Array.isArray(f.numerator) ? f.numerator.map(convertOMMLToLinear).join("") : "";
    const den = Array.isArray(f.denominator) ? f.denominator.map(convertOMMLToLinear).join("") : "";
    return `\\frac{${num}}{${den}}`;
  }

  if (rec.radical && typeof rec.radical === "object") {
    const r = rec.radical as Record<string, unknown>;
    const base = Array.isArray(r.children) ? r.children.map(convertOMMLToLinear).join("") : "";
    const deg = Array.isArray(r.degree) ? r.degree.map(convertOMMLToLinear).join("") : "";
    return deg ? `\\sqrt[${deg}]{${base}}` : `\\sqrt{${base}}`;
  }

  if (rec.superScript && typeof rec.superScript === "object") {
    const s = rec.superScript as Record<string, unknown>;
    const base = Array.isArray(s.children) ? s.children.map(convertOMMLToLinear).join("") : "";
    const sup = Array.isArray(s.superScript) ? s.superScript.map(convertOMMLToLinear).join("") : "";
    return `${base}^{${sup}}`;
  }

  if (rec.subScript && typeof rec.subScript === "object") {
    const s = rec.subScript as Record<string, unknown>;
    const base = Array.isArray(s.children) ? s.children.map(convertOMMLToLinear).join("") : "";
    const sub = Array.isArray(s.subScript) ? s.subScript.map(convertOMMLToLinear).join("") : "";
    return `${base}_{${sub}}`;
  }

  if (rec.run && typeof rec.run === "object") {
    const r = rec.run as Record<string, unknown>;
    return typeof r.text === "string" ? r.text : "";
  }

  if (Array.isArray(rec.children)) {
    return rec.children.map(convertOMMLToLinear).join("");
  }

  return "";
}

export const MathInline = Node.create({
  name: "mathInline",
  group: "inline",
  inline: true,
  atom: true,

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
          return {
            linear,
            math: convertLinearToOMML(linear),
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
