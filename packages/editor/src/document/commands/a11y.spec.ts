// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";

import { scanA11yIssues } from "../../ui/components/workspace/a11y-checker-pane";
import { A11yMirror } from "../canvas/a11y-mirror";

describe("accessibility engine", () => {
  it("scans documents for accessibility issues", () => {
    const doc = {
      type: "doc",
      content: [
        {
          type: "heading",
          attrs: { level: 1 },
          content: [{ type: "text", text: "Main Title" }],
        },
        {
          type: "heading",
          attrs: { level: 3 }, // Jump from 1 to 3!
          content: [{ type: "text", text: "Skipped Heading" }],
        },
        {
          type: "image",
          attrs: { src: "test.png", alt: "" }, // Missing alt!
        },
        {
          type: "table",
          content: [
            {
              type: "tableRow",
              content: [
                {
                  type: "tableCell",
                  content: [{ type: "paragraph", content: [{ type: "text", text: "Data" }] }],
                },
              ],
            },
          ],
        },
        {
          type: "paragraph",
          content: [
            {
              type: "text",
              text: "click here",
              marks: [{ type: "link", attrs: { href: "https://example.com" } }],
            },
          ],
        },
      ],
    };

    const issues = scanA11yIssues(doc as any);
    expect(issues.length).toBe(4);

    const jump = issues.find((i) => i.rule === "heading-order");
    expect(jump).toBeDefined();
    expect(jump?.severity).toBe("warning");

    const alt = issues.find((i) => i.rule === "alt-text");
    expect(alt).toBeDefined();
    expect(alt?.severity).toBe("error");

    const tbl = issues.find((i) => i.rule === "table-header");
    expect(tbl).toBeDefined();
    expect(tbl?.severity).toBe("warning");

    const link = issues.find((i) => i.rule === "link-text");
    expect(link).toBeDefined();
    expect(link?.severity).toBe("tip");
  });

  it("builds a semantic off-screen DOM tree from LayoutDoc", () => {
    const mirror = new A11yMirror();
    expect(mirror.root).toBeDefined();
    expect(mirror.contentContainer).toBeDefined();

    const mockLayoutDoc: any = {
      sections: [
        {
          blocks: [
            {
              kind: "paragraph",
              inline: [
                { kind: "text", text: "Hello accessibility " },
                { kind: "text", text: "link", link: { url: "https://example.com" } },
                { kind: "math", label: "x = y" },
              ],
            },
            {
              kind: "table",
              rows: [
                {
                  cells: [
                    {
                      blocks: [
                        { kind: "paragraph", inline: [{ kind: "text", text: "Header Col" }] },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    mirror.update(mockLayoutDoc);

    const p = mirror.contentContainer.querySelector("p");
    expect(p).not.toBeNull();
    expect(p?.textContent).toContain("Hello accessibility");

    const a = mirror.contentContainer.querySelector("a");
    expect(a).not.toBeNull();
    expect(a?.getAttribute("href")).toBe("https://example.com");

    const math = mirror.contentContainer.querySelector('[role="math"]');
    expect(math).not.toBeNull();

    const table = mirror.contentContainer.querySelector("table");
    expect(table).not.toBeNull();
    expect(table?.getAttribute("role")).toBe("table");
  });
});
