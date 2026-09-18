// @vitest-environment happy-dom
import { UNSUPPORTED_CONTENT_KINDS } from "@docen/docx";
import { describe, expect, it } from "vitest";

import { CONTENT_WARNING_KIND_KEYS, contentWarningText } from "./content-warning";
import { ribbonEn, ribbonZhCN } from "./i18n";

/**
 * Item-14 editor surface: the warning bar's translations must cover every
 * report kind in both locales, and the rendered text must name what is
 * uneditable instead of showing raw keys.
 */
describe("unsupported-content warning text", () => {
  it("defines a label key for every report kind", () => {
    expect(Object.keys(CONTENT_WARNING_KIND_KEYS).sort()).toEqual(
      [...UNSUPPORTED_CONTENT_KINDS].sort(),
    );
  });

  it("has every key in both locales", () => {
    const keys = [
      "contentWarning.title",
      "contentWarning.body",
      "contentWarning.dismiss",
      ...Object.values(CONTENT_WARNING_KIND_KEYS),
    ];
    for (const key of keys) {
      expect(ribbonEn.translations[key], `en ${key}`).toBeTruthy();
      expect(ribbonZhCN.translations[key], `zh ${key}`).toBeTruthy();
    }
  });

  it("renders counted localized labels, not raw keys", () => {
    const translate = (key: string): string => ribbonEn.translations[key] ?? key;
    const text = contentWarningText(
      {
        items: [
          { kind: "altChunk", count: 1 },
          { kind: "smartArt", count: 2 },
        ],
        total: 3,
      },
      translate,
    );
    expect(text).toContain(translate("contentWarning.title"));
    expect(text).toContain("imported content (altChunk)");
    expect(text).toContain("2 × SmartArt graphics");
    expect(text).not.toMatch(/contentWarning\./);
  });
});
