import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCXSync, type JSONContent } from "../index";
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

describe("math structured run operands", () => {
  /**
   * A parsed m:oMath carries its runs as `{ text, properties }` (the m:rPr
   * style), not the bare `{ text }` the authoring path emits. The linear label
   * used to lose every such operand (`\frac{}{}`); it must linearize both run
   * shapes, and the structured operands must survive the package round-trip
   * on the verbatim `math` attr.
   */
  const math = {
    display: false,
    children: [
      {
        fraction: {
          numerator: [{ text: "a", properties: { style: "italic" } }],
          denominator: [{ text: "b", properties: { style: "italic" } }],
        },
      },
      { text: " + " },
      { radical: { children: [{ text: "x+1" }], degree: [{ text: "3" }] } },
    ],
  };

  it("linearizes structured and bare run operands alike", () => {
    expect(convertOMMLToLinear(math)).toBe("\\frac{a}{b} + \\sqrt[3]{x+1}");
  });

  it("keeps m:oMath attrs and operands through an export/import cycle", () => {
    const json: JSONContent = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "mathInline", attrs: { math, linear: convertOMMLToLinear(math) } }],
        },
      ],
    };
    const gen1 = generateDOCXSync(json, { prepare: false }) as Uint8Array;
    const xml1 = new TextDecoder().decode(unzipSync(gen1)["word/document.xml"]);
    expect(xml1).toContain("<m:oMath>");
    expect(xml1).toContain("<m:f>");
    expect(xml1).toContain('<m:sty m:val="i"/>');
    expect(xml1).toContain("<m:t>a</m:t>");
    expect(xml1).toContain("<m:t>b</m:t>");

    const reparsed = parseDOCXSync(gen1);
    const node = reparsed.content?.[0]?.content?.find((c) => c.type === "mathInline");
    expect(node?.attrs?.linear).toBe("\\frac{a}{b} + \\sqrt[3]{x+1}");
    const operands = JSON.stringify(node?.attrs?.math);
    expect(operands).toContain('"text":"a"');
    expect(operands).toContain('"style":"italic"');

    const xml2 = new TextDecoder().decode(
      unzipSync(generateDOCXSync(reparsed, { prepare: false }) as Uint8Array)["word/document.xml"],
    );
    expect(xml2).toContain("<m:f>");
    expect(xml2).toContain("<m:t>a</m:t>");
    expect(xml2).toContain("<m:t>x+1</m:t>");
  });
});
