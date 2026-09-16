import { describe, expect, it } from "vitest";

import type { JSONContent } from "../core";
import { generateHTML, parseHTML } from "./html";

describe("HTML Converter", () => {
  it("generates styled HTML from JSONContent", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "heading",
          attrs: { level: 1, textAlign: "center" },
          content: [{ type: "text", text: "Main Title" }],
        },
        {
          type: "paragraph",
          content: [
            { type: "text", text: "Hello " },
            { type: "text", text: "world", marks: [{ type: "bold" }, { type: "italic" }] },
            { type: "text", text: ", " },
            {
              type: "text",
              text: "link",
              marks: [{ type: "link", attrs: { href: "https://example.com" } }],
            },
          ],
        },
      ],
    };

    const html = generateHTML(doc);
    expect(html).toContain('<h1 style="text-align: center">Main Title</h1>');
    expect(html).toContain("<em><strong>world</strong></em>");
    expect(html).toContain('<a href="https://example.com">link</a>');
  });

  it("generates full HTML document with styles", () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "Hello doc" }],
        },
      ],
    };

    const full = generateHTML(doc, { fullDocument: true, title: "Test Document" });
    expect(full).toContain("<!DOCTYPE html>");
    expect(full).toContain("<title>Test Document</title>");
    expect(full).toContain("<p>Hello doc</p>");
  });

  it("parses HTML into JSONContent", async () => {
    const html = "<p>Testing <strong>bold</strong> and <em>italic</em></p>";
    const parsed = await parseHTML(html);
    expect(parsed.type).toBe("doc");
    expect(parsed.content).toHaveLength(1);
    const p = parsed.content![0]!;
    expect(p.type).toBe("paragraph");
    const boldRun = p.content!.find((c) => c.text === "bold");
    expect(boldRun?.marks?.some((m) => m.type === "bold")).toBe(true);
  });
});
