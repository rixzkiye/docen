import { describe, expect, it } from "vitest";

import type { JSONContent } from "../core";
import { generatePlainText, parsePlainText } from "./text";

describe("Plain Text Converter", () => {
  it("generates plain text from paragraphs, lists, and tables", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "Paragraph 1" }],
        },
        {
          type: "bulletList",
          content: [
            {
              type: "listItem",
              content: [{ type: "paragraph", content: [{ type: "text", text: "Item A" }] }],
            },
            {
              type: "listItem",
              content: [{ type: "paragraph", content: [{ type: "text", text: "Item B" }] }],
            },
          ],
        },
      ],
    };

    const text = generatePlainText(doc);
    expect(text).toContain("Paragraph 1");
    expect(text).toContain("• Item A\n• Item B");
  });

  it("parses plain text into JSONContent paragraphs", () => {
    const raw = "First paragraph\n\nSecond paragraph";
    const doc = parsePlainText(raw);
    expect(doc.type).toBe("doc");
    expect(doc.content).toHaveLength(2);
    expect(doc.content![0]!.content![0]!.text).toBe("First paragraph");
    expect(doc.content![1]!.content![0]!.text).toBe("Second paragraph");
  });
});
