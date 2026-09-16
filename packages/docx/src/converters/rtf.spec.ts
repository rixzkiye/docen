import { describe, expect, it } from "vitest";

import type { JSONContent } from "../core";
import { generateRTF, parseRTF } from "./rtf";

describe("RTF Converter", () => {
  it("generates and parses basic formatted text", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "Normal text " },
            { type: "text", text: "Bold", marks: [{ type: "bold" }] },
            { type: "text", text: " and " },
            { type: "text", text: "Italic", marks: [{ type: "italic" }] },
            { type: "text", text: " and " },
            { type: "text", text: "Underline", marks: [{ type: "underline" }] },
          ],
        },
      ],
    };

    const rtf = generateRTF(doc);
    expect(rtf).toContain("{\\rtf1");
    expect(rtf).toContain("\\b Bold\\b0");
    expect(rtf).toContain("\\i Italic\\i0");
    expect(rtf).toContain("\\ul Underline\\ulnone");

    const parsed = parseRTF(rtf);
    expect(parsed.type).toBe("doc");
    expect(parsed.content).toHaveLength(1);
    const p = parsed.content![0]!;
    expect(p.type).toBe("paragraph");

    const texts = p.content!.map((c) => c.text).join("");
    expect(texts).toBe("Normal text Bold and Italic and Underline");

    const boldRun = p.content!.find((c) => c.text === "Bold");
    expect(boldRun?.marks?.some((m) => m.type === "bold")).toBe(true);

    const italicRun = p.content!.find((c) => c.text === "Italic");
    expect(italicRun?.marks?.some((m) => m.type === "italic")).toBe(true);

    const ulRun = p.content!.find((c) => c.text === "Underline");
    expect(ulRun?.marks?.some((m) => m.type === "underline")).toBe(true);
  });

  it("handles colors and font sizes", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "text",
              text: "Colored",
              marks: [{ type: "textStyle", attrs: { color: "#ff0000", fontSize: 16 } }],
            },
          ],
        },
      ],
    };

    const rtf = generateRTF(doc);
    expect(rtf).toContain("\\colortbl");
    expect(rtf).toContain("\\red255\\green0\\blue0");
    expect(rtf).toContain("\\cf1");
    expect(rtf).toContain("\\fs32");

    const parsed = parseRTF(rtf);
    const run = parsed.content![0]!.content![0]!;
    expect(run.text).toBe("Colored");
    const styleMark = run.marks?.find((m) => m.type === "textStyle");
    expect(styleMark?.attrs?.color).toBe("#ff0000");
    expect(styleMark?.attrs?.fontSize).toBe("16pt");
  });

  it("handles unicode characters", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "Hello 世界 🌍" }],
        },
      ],
    };

    const rtf = generateRTF(doc);
    expect(rtf).toContain("\\u");

    const parsed = parseRTF(rtf);
    const text = parsed.content![0]!.content!.map((c) => c.text).join("");
    expect(text).toContain("Hello 世界");
  });

  it("generates headings and paragraph alignments", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "heading",
          attrs: { level: 1, textAlign: "center" },
          content: [{ type: "text", text: "Title" }],
        },
      ],
    };

    const rtf = generateRTF(doc);
    expect(rtf).toContain("\\qc");
    expect(rtf).toContain("\\outlinelevel0");
    expect(rtf).toContain("Title");
  });

  it("generates lists and tables", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "bulletList",
          content: [
            {
              type: "listItem",
              content: [
                {
                  type: "paragraph",
                  content: [{ type: "text", text: "Bullet item" }],
                },
              ],
            },
          ],
        },
        {
          type: "table",
          content: [
            {
              type: "tableRow",
              content: [
                {
                  type: "tableCell",
                  content: [
                    {
                      type: "paragraph",
                      content: [{ type: "text", text: "Cell 1" }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const rtf = generateRTF(doc);
    expect(rtf).toContain("\\bullet");
    expect(rtf).toContain("Bullet item");
    expect(rtf).toContain("\\trowd");
    expect(rtf).toContain("Cell 1");
    expect(rtf).toContain("\\row");
  });
});
