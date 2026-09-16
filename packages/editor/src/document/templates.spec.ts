import type { JSONContent } from "@docen/docx";
import { describe, expect, it } from "vitest";

import { BUILTIN_TEMPLATES, findTemplate, templateLocale } from "./templates";

const blocksOf = (json: JSONContent): JSONContent[] => json.content ?? [];
const headingOf = (node: JSONContent): unknown => node.attrs?.heading;

describe("built-in template catalog", () => {
  it("ships at least three templates with unique ids and metadata", () => {
    expect(BUILTIN_TEMPLATES.length).toBeGreaterThanOrEqual(3);
    const ids = BUILTIN_TEMPLATES.map((tpl) => tpl.id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const tpl of BUILTIN_TEMPLATES) {
      expect(tpl.nameKey.length).toBeGreaterThan(0);
      expect(tpl.descriptionKey.length).toBeGreaterThan(0);
      expect(typeof tpl.build).toBe("function");
    }
  });

  it("finds templates by id and misses cleanly", () => {
    expect(findTemplate("report")?.id).toBe("report");
    expect(findTemplate("does-not-exist")).toBeUndefined();
  });

  it("returns a fresh, independent model per build", () => {
    const tpl = findTemplate("report")!;
    const first = tpl.build();
    const second = tpl.build();
    expect(first).not.toBe(second);
    expect(first).toEqual(second);
    // Mutating one instance must not leak into the template or the next build.
    blocksOf(first).push({ type: "paragraph" });
    expect(blocksOf(tpl.build()).length).toBe(blocksOf(second).length);
  });
});

describe("blank template", () => {
  it("is one empty paragraph (the editor normalizes the document defaults)", () => {
    const json = findTemplate("blank")!.build();
    expect(json.type).toBe("doc");
    expect(blocksOf(json)).toEqual([{ type: "paragraph" }]);
  });
});

describe("report template", () => {
  it("carries a title, heading outline, and a placeholder TOC field", () => {
    const blocks = blocksOf(findTemplate("report")!.build());
    expect(headingOf(blocks[0]!)).toBe("Title");
    const headings = blocks.filter((node) => headingOf(node) === "Heading1");
    expect(headings).toHaveLength(3);
    const toc = blocks.find((node) => node.type === "tocField");
    expect(toc).toBeDefined();
    expect(toc!.attrs?.options).toEqual({ headingStyleRange: "1-3", hyperlink: true });
    // The tocField schema is block+ — the placeholder paragraph keeps it valid
    // until Update Field renders the entries.
    expect(blocksOf(toc!)).toEqual([{ type: "paragraph" }]);
    // No empty text placeholders in the body.
    const bodies = blocks.filter(
      (node) => headingOf(node) !== "Title" && node.type === "paragraph",
    );
    for (const paragraph of bodies) {
      expect((paragraph.content?.[0]?.text ?? "").length).toBeGreaterThan(0);
    }
  });
});

describe("letter template", () => {
  it("carries sender, recipient, salutation, and closing placeholders", () => {
    const blocks = blocksOf(findTemplate("letter")!.build());
    const texts = blocks.map((node) => node.content?.[0]?.text ?? "");
    expect(texts.some((text) => text.includes("Sincerely"))).toBe(true);
    expect(texts.some((text) => text.startsWith("Dear"))).toBe(true);
    // Multi-line address placeholders become one paragraph per line.
    expect(texts.filter((text) => text === "Company")).toHaveLength(1);
  });
});

describe("template localization", () => {
  it("maps language tags to body locales", () => {
    expect(templateLocale("zh-CN")).toBe("zh-CN");
    expect(templateLocale("zh")).toBe("zh-CN");
    expect(templateLocale("en-US")).toBe("en");
    expect(templateLocale(undefined)).toBe("en");
    expect(templateLocale(null)).toBe("en");
  });

  it("writes the body in the requested locale (en is the default)", () => {
    const tpl = findTemplate("report")!;
    const en = tpl.build("en");
    const zh = tpl.build("zh-CN");
    expect(en.content?.[0]?.content?.[0]?.text).toBe("Report Title");
    expect(zh.content?.[0]?.content?.[0]?.text).toBe("报告标题");
    expect(tpl.build().content?.[0]?.content?.[0]?.text).toBe("Report Title");
  });
});
