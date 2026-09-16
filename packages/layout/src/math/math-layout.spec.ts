import { describe, expect, it } from "vitest";

import { layoutMathTree } from "./math-layout";

describe("Math Layout Engine", () => {
  it("lays out stacked fractions with horizontal fraction line", () => {
    const fractionMath = {
      fraction: {
        numerator: "1",
        denominator: "2",
      },
    };
    const layout = layoutMathTree(fractionMath, 14);
    expect(layout.elements.length).toBe(1);
    const frac = layout.elements[0]!;
    expect(frac.kind).toBe("fraction");
    expect(frac.fractionLine).toBeDefined();
    expect(frac.slots?.length).toBe(2);
    expect(frac.slots?.[0]?.text).toBe("1");
    expect(frac.slots?.[1]?.text).toBe("2");
    expect(frac.widthPx).toBeGreaterThan(0);
    expect(frac.heightPx).toBeGreaterThan(14);
  });

  it("lays out radicals with surd bar", () => {
    const radicalMath = {
      radical: {
        children: "x + y",
        degree: "3",
      },
    };
    const layout = layoutMathTree(radicalMath, 14);
    expect(layout.elements.length).toBe(1);
    const rad = layout.elements[0]!;
    expect(rad.kind).toBe("radical");
    expect(rad.radicalBar).toBeDefined();
    expect(rad.symbol).toBe("3");
    expect(rad.slots?.[0]?.text).toBe("x + y");
  });

  it("lays out n-ary large operators with limits", () => {
    const sumMath = {
      sum: {
        children: "i",
        subScript: "i=1",
        superScript: "n",
      },
    };
    const layout = layoutMathTree(sumMath, 14);
    expect(layout.elements.length).toBe(1);
    const sum = layout.elements[0]!;
    expect(sum.kind).toBe("nary");
    expect(sum.symbol).toBe("∑");
    expect(sum.slots?.length).toBe(3);
  });

  it("lays out delimiters enclosing content", () => {
    const delimMath = {
      delimiter: {
        open: "[",
        close: "]",
        children: "matrix",
      },
    };
    const layout = layoutMathTree(delimMath, 14);
    expect(layout.elements.length).toBe(1);
    const delim = layout.elements[0]!;
    expect(delim.kind).toBe("delimiter");
    expect(delim.symbol).toBe("[]");
  });
});
