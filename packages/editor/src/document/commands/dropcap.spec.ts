import { projectDocumentOptions } from "@docen/docx/layout";
import { describe, expect, it } from "vitest";

describe("drop cap projection", () => {
  it("projects dropped drop cap correctly", () => {
    const doc = {
      sections: [
        {
          children: [
            {
              paragraph: {
                dropCap: { val: "drop", lines: 4, distance: 150 },
                children: ["Once upon a time in a faraway kingdom..."],
              },
            },
          ],
        },
      ],
    };
    const projected = projectDocumentOptions(doc as any);
    const para = projected.sections[0].blocks[0] as any;
    expect(para.kind).toBe("paragraph");
    expect(para.dropCap).toBeDefined();
    expect(para.dropCap.type).toBe("dropped");
    expect(para.dropCap.lines).toBe(4);
    expect(para.dropCap.distancePx).toBe(10); // 150 twips = 10 px
  });

  it("projects in-margin drop cap correctly", () => {
    const doc = {
      sections: [
        {
          children: [
            {
              paragraph: {
                dropCap: { val: "margin", lines: 3 },
                children: ["Paragraph with margin drop cap"],
              },
            },
          ],
        },
      ],
    };
    const projected = projectDocumentOptions(doc as any);
    const para = projected.sections[0].blocks[0] as any;
    expect(para.dropCap).toBeDefined();
    expect(para.dropCap.type).toBe("margin");
    expect(para.dropCap.lines).toBe(3);
  });
});
