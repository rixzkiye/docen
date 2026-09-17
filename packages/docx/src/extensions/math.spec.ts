import { describe, expect, it } from "vitest";

import { MathInline, convertLinearToOMML, convertOMMLToLinear, parseDocxInline } from "./math";

describe("Math extension and conversion", () => {
  it("converts fraction between linear and OMML format", () => {
    const omml = convertLinearToOMML("\\frac{x}{y}");
    expect(omml).toEqual({
      fraction: {
        numerator: ["x"],
        denominator: ["y"],
      },
    });

    const linear = convertOMMLToLinear(omml);
    expect(linear).toBe("\\frac{x}{y}");
  });

  it("converts radical between linear and OMML format", () => {
    const omml = convertLinearToOMML("\\sqrt{x+1}");
    expect(omml).toEqual({
      radical: {
        children: ["x+1"],
      },
    });
    expect(convertOMMLToLinear(omml)).toBe("\\sqrt{x+1}");

    const degOmml = convertLinearToOMML("\\sqrt[3]{8}");
    expect(degOmml).toEqual({
      radical: {
        children: ["8"],
        degree: ["3"],
      },
    });
    expect(convertOMMLToLinear(degOmml)).toBe("\\sqrt[3]{8}");
  });

  it("converts superscripts and subscripts", () => {
    const supOmml = convertLinearToOMML("x^{2}");
    expect(supOmml).toEqual({
      superScript: {
        children: ["x"],
        superScript: ["2"],
      },
    });
    expect(convertOMMLToLinear(supOmml)).toBe("x^{2}");

    const subOmml = convertLinearToOMML("a_{i}");
    expect(subOmml).toEqual({
      subScript: {
        children: ["a"],
        subScript: ["i"],
      },
    });
    expect(convertOMMLToLinear(subOmml)).toBe("a_{i}");
  });

  it("defines MathInline extension with proper attributes", () => {
    expect(MathInline.name).toBe("mathInline");
    const attrs = (MathInline as any).config.addAttributes();
    expect(attrs).toHaveProperty("math");
    expect(attrs).toHaveProperty("linear");
  });

  it("claims the DOCX math branch but not the run-level math flag", () => {
    const branch = { math: { fraction: { numerator: ["a"], denominator: ["b"] } } };
    expect(parseDocxInline.match(branch as never, {} as never)).toBe(true);
    // A plain run carrying the rPr `math: true` flag is not the branch.
    expect(parseDocxInline.match({ math: true } as never, {} as never)).toBe(false);
    expect(parseDocxInline.convert(branch as never, {} as never)).toEqual({
      type: "mathInline",
      attrs: {
        math: { fraction: { numerator: ["a"], denominator: ["b"] } },
        linear: "\\frac{a}{b}",
      },
    });
  });
});
