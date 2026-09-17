// @vitest-environment happy-dom
// Insert → Equation seeds the single math representation (mathInline) and the
// node round-trips through the real DOCX pipeline.
import { generateDOCXSync, parseDOCXSync } from "@docen/docx";
import { describe, expect, it } from "vitest";

import { equationSeed } from "./equation";

describe("equationSeed", () => {
  it("seeds a mathInline node with the MathInput payload and a linear label", () => {
    const seed = equationSeed("fraction")!;
    expect(seed.type).toBe("mathInline");
    expect(seed.attrs?.math).toEqual({
      children: [{ fraction: { numerator: [{ text: "" }], denominator: [{ text: "" }] } }],
    });
    expect(seed.attrs?.linear).toBe("\\frac{}{}");
    expect(equationSeed("nope")).toBeNull();
  });

  it("round-trips through generateDOCXSync → parseDOCXSync as mathInline", () => {
    const seed = equationSeed("fraction")!;
    const json = {
      type: "doc",
      content: [{ type: "paragraph", content: [seed] }],
    };
    const bytes = generateDOCXSync(json as never);
    const reparsed = parseDOCXSync(bytes) as never as {
      content: Array<{ content?: Array<{ type?: string; attrs?: Record<string, unknown> }> }>;
    };
    const inline = reparsed.content[0]?.content?.[0];
    expect(inline?.type).toBe("mathInline");
    // office-open normalizes the empty slot runs ({text:""} → ""); the
    // fraction structure itself must survive verbatim.
    const math = inline?.attrs?.math as {
      children?: Array<{ fraction?: { numerator?: unknown[]; denominator?: unknown[] } }>;
    };
    const fraction = math?.children?.[0]?.fraction;
    expect(fraction).toBeTruthy();
    expect(fraction?.numerator).toHaveLength(1);
    expect(fraction?.denominator).toHaveLength(1);
  });
});
