import { convertOMMLToLinear, type JSONContent } from "@docen/docx";

/** The Insert → Equation gallery templates (Word's frequent structures) as
 *  mathInline seeds: each argument is an empty run — the □ slot; the radical's
 *  absent degree reads as the square root (degHide follows). Shared by the
 *  host's insert path and the tests. */
export function equationSeed(template: string): JSONContent | null {
  const slot = (): object => ({ text: "" });
  const templates: Record<string, object> = {
    // Alt+= inserts a blank equation — one empty run to type into.
    plain: { text: "" },
    fraction: { fraction: { numerator: [slot()], denominator: [slot()] } },
    superScript: { superScript: { children: [slot()], superScript: [slot()] } },
    radical: { radical: { children: [slot()] } },
    sum: {
      sum: {
        children: [slot()],
        subScript: [slot()],
        superScript: [slot()],
        properties: { limitLocation: "undOvr" },
      },
    },
    integral: {
      integral: {
        children: [slot()],
        subScript: [slot()],
        superScript: [slot()],
        properties: { limitLocation: "subSup" },
      },
    },
  };
  const shape = templates[template];
  if (!shape) return null;
  return {
    type: "mathInline",
    attrs: {
      math: { children: [shape] },
      linear: convertOMMLToLinear(shape),
    },
  };
}
