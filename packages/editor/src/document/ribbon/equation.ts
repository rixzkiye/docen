import type { RibbonTab } from "../../ui";
import { grid, group, opt, tab } from "./shared";

/** The math symbols of the equation context tab's Symbols grid — Word's Basic
 *  Math panel's frequent picks (operators, Greek letters, arrows). */
export const EQUATION_SYMBOLS =
  "± ∞ ≠ ≈ ≤ ≥ ∝ × ÷ ∑ ∫ √ ∂ π ∇ α β γ δ ε θ λ σ μ ω Δ Ω → ⇒ ∈ ∀ ∃".split(" ");

/** Word's Equation Tools — the contextual tab while the selection sits on a
 *  math atom. Structures replays the Insert tab's equation templates (the
 *  engine's five OMML shapes plus the blank Alt+= insert), Symbols is a grid
 *  of characters inserted at the caret. Marked `contextual` so
 *  {@link ribbonTabs} excludes it from the static render; the host appends it
 *  via {@link buildContextualTab} as the selection enters/leaves a math atom. */
export function equationContextTab(): RibbonTab {
  return {
    id: "equation",
    label: tab("equation"),
    contextual: true,
    groups: [
      group("structures", [
        grid([
          { type: "button", label: opt("equation-blank"), event: "equation", value: "plain" },
          {
            type: "button",
            label: opt("equation-fraction"),
            event: "equation",
            value: "fraction",
          },
          {
            type: "button",
            label: opt("equation-script"),
            event: "equation",
            value: "superScript",
          },
          {
            type: "button",
            label: opt("equation-radical"),
            event: "equation",
            value: "radical",
          },
          { type: "button", label: opt("equation-sum"), event: "equation", value: "sum" },
          {
            type: "button",
            label: opt("equation-integral"),
            event: "equation",
            value: "integral",
          },
        ]),
      ]),
      group("symbols", [
        grid(
          EQUATION_SYMBOLS.map((char) => ({
            type: "button" as const,
            label: char,
            event: "insert-symbol",
            value: char,
          })),
        ),
      ]),
    ],
  };
}
