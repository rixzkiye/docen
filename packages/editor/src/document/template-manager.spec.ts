import { describe, expect, it } from "vitest";

import { mergeTemplateStyles, parseThemeXml } from "./template-manager";
import { BUILTIN_TEMPLATES, findTemplate } from "./templates";

describe("Template Manager: parseThemeXml (W7.3)", () => {
  it("extracts theme colors and fonts from DrawingML theme XML", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Aptos Display"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Aptos"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`;

    const theme = parseThemeXml(xml);
    expect(theme).toBeDefined();
    expect(theme?.name).toBe("Office");
    expect(theme?.colors.accent1).toBe("4f81bd");
    expect(theme?.colors.accent2).toBe("c0504d");
    expect(theme?.colors.accent3).toBe("9bbb59");
    expect(theme?.colors.dk1).toBe("000000");
    expect(theme?.colors.lt1).toBe("ffffff");
    expect(theme?.fonts.majorFont).toBe("Aptos Display");
    expect(theme?.fonts.minorFont).toBe("Aptos");
  });

  it("handles empty or partial XML with safe fallbacks", () => {
    expect(parseThemeXml("")).toBeUndefined();

    const partialXml = `<a:theme><a:clrScheme name="Custom"></a:clrScheme></a:theme>`;
    const theme = parseThemeXml(partialXml);
    expect(theme).toBeDefined();
    expect(theme?.colors.accent1).toBe("4f81bd"); // fallback
    expect(theme?.fonts.majorFont).toBe("Calibri Light");
  });
});

describe("Template Manager: mergeTemplateStyles", () => {
  it("merges paragraph styles, updating matching and adding new", () => {
    const current = {
      paragraphStyles: [
        { id: "Normal", name: "Normal", run: { font: "Arial" } },
        { id: "Custom1", name: "Custom 1", run: { color: "000000" } },
      ],
    };

    const template = {
      paragraphStyles: [
        { id: "Normal", name: "Normal", run: { font: "Calibri", size: 24 } },
        { id: "TemplateStyle", name: "Template Style", run: { bold: true } },
      ],
    };

    const merged = mergeTemplateStyles(current as any, template as any);
    expect(merged.paragraphStyles).toHaveLength(3);

    const normal = merged.paragraphStyles?.find((s) => s.id === "Normal");
    expect(normal?.run?.font).toBe("Calibri");
    expect(normal?.run?.size).toBe(24);

    const custom1 = merged.paragraphStyles?.find((s) => s.id === "Custom1");
    expect(custom1).toBeDefined();

    const tplStyle = merged.paragraphStyles?.find((s) => s.id === "TemplateStyle");
    expect(tplStyle?.run?.bold).toBe(true);
  });

  it("merges character styles and default heading styles", () => {
    const current = {
      default: { heading1: { run: { size: 28 } } },
    };

    const template = {
      default: {
        heading1: { run: { size: 32, bold: true } },
        title: { run: { size: 40 } },
      },
    };

    const merged = mergeTemplateStyles(current as any, template as any);
    expect(merged.default?.heading1?.run?.size).toBe(32);
    expect(merged.default?.title?.run?.size).toBe(40);
  });

  it("handles null or undefined style containers gracefully", () => {
    expect(mergeTemplateStyles(undefined, undefined)).toEqual({});
    const cur = { paragraphStyles: [{ id: "A", name: "A" }] };
    expect(mergeTemplateStyles(cur as any, undefined)).toEqual(cur);
    const tpl = { paragraphStyles: [{ id: "B", name: "B" }] };
    expect(mergeTemplateStyles(undefined, tpl as any)).toEqual(tpl);
  });
});

describe("Template Gallery: BUILTIN_TEMPLATES (W7.3)", () => {
  it("includes all required built-in templates (Blank, Report, Letter, Resume, Invoice)", () => {
    const ids = BUILTIN_TEMPLATES.map((t) => t.id);
    expect(ids).toContain("blank");
    expect(ids).toContain("report");
    expect(ids).toContain("letter");
    expect(ids).toContain("resume");
    expect(ids).toContain("invoice");
    expect(ids).toHaveLength(5);
  });

  it("builds valid JSONContent for Resume template in en and zh-CN", () => {
    const resume = findTemplate("resume")!;
    expect(resume).toBeDefined();

    const enDoc = resume.build("en");
    expect(enDoc.type).toBe("doc");
    expect(enDoc.content?.length).toBeGreaterThan(5);

    const zhDoc = resume.build("zh-CN");
    expect(zhDoc.type).toBe("doc");
    expect(zhDoc.content?.length).toBeGreaterThan(5);
  });

  it("builds valid JSONContent for Invoice template with table in en and zh-CN", () => {
    const invoice = findTemplate("invoice")!;
    expect(invoice).toBeDefined();

    const enDoc = invoice.build("en");
    expect(enDoc.type).toBe("doc");
    const tableNode = enDoc.content?.find((node) => node.type === "table");
    expect(tableNode).toBeDefined();
    expect(tableNode?.content?.length).toBeGreaterThanOrEqual(4);

    const zhDoc = invoice.build("zh-CN");
    expect(zhDoc.type).toBe("doc");
    const zhTable = zhDoc.content?.find((node) => node.type === "table");
    expect(zhTable).toBeDefined();
  });
});
