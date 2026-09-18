import { sectionMarginDefaults } from "@office-open/docx";
import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";

import {
  DOCEN_DEFAULT_EAST_ASIA_LANGUAGE,
  resolveFontName,
  resolvePageSize,
} from "../extensions/utils";
import {
  DOCEN_DEFAULT_PAGE_MARGIN,
  DOCEN_DEFAULT_PAGE_SIZE,
  compileDocument,
  docenDefaultSectionProperties,
  generateDOCXSync,
  normalizeDocument,
  parseDOCXSync,
} from "../index";
import { projectFlowBox } from "../layout/project/page";

/**
 * Item-12 guard: generation must carry docen-owned page geometry and never
 * inherit office-open's zh-CN defaults (A4 + 1800-twip side margins, 851/992
 * header/footer, 等线 theme substitution). The oracle is the generated
 * package itself — bytes in, XML out — not the implementation's own constants.
 */

/** Parse `w:pgSz` / `w:pgMar` attribute maps out of document.xml. */
function sectionGeometry(): {
  pgSz: Record<string, string>;
  pgMar: Record<string, string>;
} {
  const bytes = generateDOCXSync({ type: "doc", content: [{ type: "paragraph" }] });
  const files = unzipSync(new Uint8Array(bytes));
  const xml = strFromU8(files["word/document.xml"]);
  const attrsOf = (tag: string): Record<string, string> => {
    const el = new RegExp(`<w:${tag}([^>]*)/?>`).exec(xml)?.[1] ?? "";
    return Object.fromEntries(
      [...el.matchAll(/w:([A-Za-z]+)="([^"]*)"/g)].map((m) => [m[1]!, m[2]!]),
    );
  };
  return { pgSz: attrsOf("pgSz"), pgMar: attrsOf("pgMar") };
}

describe("docen-owned section defaults (item 12)", () => {
  it("stamps explicit A4 geometry instead of office-open's zh-CN defaults", () => {
    const { pgSz, pgMar } = sectionGeometry();
    expect(pgSz).toMatchObject({
      w: String(DOCEN_DEFAULT_PAGE_SIZE.WIDTH),
      h: String(DOCEN_DEFAULT_PAGE_SIZE.HEIGHT),
    });
    expect(pgMar).toMatchObject({
      top: String(DOCEN_DEFAULT_PAGE_MARGIN.TOP),
      right: String(DOCEN_DEFAULT_PAGE_MARGIN.RIGHT),
      bottom: String(DOCEN_DEFAULT_PAGE_MARGIN.BOTTOM),
      left: String(DOCEN_DEFAULT_PAGE_MARGIN.LEFT),
      header: String(DOCEN_DEFAULT_PAGE_MARGIN.HEADER),
      footer: String(DOCEN_DEFAULT_PAGE_MARGIN.FOOTER),
    });
    // office-open's zh-CN Normal side margins (1.25") must not leak.
    expect(pgMar.left).not.toBe("1800");
    expect(pgMar.right).not.toBe("1800");
    expect(pgMar.header).not.toBe("851");
    expect(pgMar.footer).not.toBe("992");
    expect(sectionMarginDefaults.LEFT).toBe(1800);
  });

  it("preserves a model-supplied sectPr verbatim", () => {
    const json = {
      type: "doc",
      content: [{ type: "paragraph" }],
      attrs: {
        sectionProperties: {
          pageSize: { width: 12240, height: 15840 },
          pageMargin: { top: 720, right: 720, bottom: 720, left: 720 },
        },
      },
    };
    const compiled = compileDocument(json);
    const props = compiled.sections[0]?.properties as
      | { pageSize?: { width?: number }; pageMargin?: { top?: number } }
      | undefined;
    expect(props?.pageSize?.width).toBe(12240);
    expect(props?.pageMargin?.top).toBe(720);
  });

  it("normalizeDocument hands hand-built docs docen geometry, not a null sectPr", () => {
    const normalized = normalizeDocument({ type: "doc", content: [{ type: "paragraph" }] });
    expect(normalized.attrs?.sectionProperties).toMatchObject({
      pageSize: { width: DOCEN_DEFAULT_PAGE_SIZE.WIDTH, height: DOCEN_DEFAULT_PAGE_SIZE.HEIGHT },
      pageMargin: { left: DOCEN_DEFAULT_PAGE_MARGIN.LEFT, right: DOCEN_DEFAULT_PAGE_MARGIN.RIGHT },
    });
  });

  it("round-trips generated geometry through parse → resolve", () => {
    const bytes = generateDOCXSync({ type: "doc", content: [{ type: "paragraph" }] });
    const resolved = parseDOCXSync(new Uint8Array(bytes));
    expect(resolved.attrs?.sectionProperties).toMatchObject({
      pageSize: { width: DOCEN_DEFAULT_PAGE_SIZE.WIDTH, height: DOCEN_DEFAULT_PAGE_SIZE.HEIGHT },
      pageMargin: { left: DOCEN_DEFAULT_PAGE_MARGIN.LEFT },
    });
  });

  it("render fallbacks use the same owned constants", () => {
    expect(resolvePageSize(undefined)).toEqual({
      width: DOCEN_DEFAULT_PAGE_SIZE.WIDTH,
      height: DOCEN_DEFAULT_PAGE_SIZE.HEIGHT,
    });
    const flow = projectFlowBox(undefined);
    const contentWidthTw = DOCEN_DEFAULT_PAGE_SIZE.WIDTH - 2 * DOCEN_DEFAULT_PAGE_MARGIN.LEFT;
    expect(flow.contentWidthPx).toBeCloseTo(contentWidthTw / 15, 5);
    const props = docenDefaultSectionProperties();
    const margin = props.pageMargin as unknown as { left?: number } | undefined;
    expect(margin?.left).toBe(DOCEN_DEFAULT_PAGE_MARGIN.LEFT);
  });

  it("does not silently substitute a zh-CN East Asian face for theme-only runs", () => {
    expect(DOCEN_DEFAULT_EAST_ASIA_LANGUAGE).toBe("en-US");
    // Theme-only run with no document language: the Latin theme face, never 等线.
    expect(resolveFontName({ eastAsiaTheme: "minorEastAsia" })).toBe("Calibri");
    expect(resolveFontName({ asciiTheme: "majorHAnsi" })).toBe("Calibri Light");
    // A caller that knows the document language still gets the CJK face.
    expect(resolveFontName({ eastAsiaTheme: "minorEastAsia" }, "zh-CN")).toBe("等线");
    expect(resolveFontName({ eastAsiaTheme: "majorEastAsia" }, "zh-CN")).toBe("等线 Light");
  });

  it("generated styles carry no CJK labels or zh-CN language", () => {
    const bytes = generateDOCXSync({ type: "doc", content: [{ type: "paragraph" }] });
    const files = unzipSync(new Uint8Array(bytes));
    const styles = strFromU8(files["word/styles.xml"]);
    expect(styles).not.toMatch(/[\u4e00-\u9fff]/);
    expect(styles).not.toContain('w:val="zh-CN"');
    expect(styles).not.toContain('w:eastAsia="zh-CN"');
    expect(styles).toContain('w:asciiTheme="minorHAnsi"');
  });
});
