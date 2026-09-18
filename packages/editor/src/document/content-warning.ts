import type { UnsupportedContentKind, UnsupportedContentReport } from "@docen/docx";

/**
 * Localization keys for the unsupported-content warning bar. Kept as literal
 * strings in one map so the i18n coverage spec can prove every kind has a
 * translation (the call site resolves them dynamically).
 */
export const CONTENT_WARNING_KIND_KEYS: Readonly<Record<UnsupportedContentKind, string>> = {
  altChunk: "contentWarning.kind.altChunk",
  subDoc: "contentWarning.kind.subDoc",
  smartArt: "contentWarning.kind.smartArt",
  oleObject: "contentWarning.kind.oleObject",
  rawXml: "contentWarning.kind.rawXml",
  customXml: "contentWarning.kind.customXml",
  contentPart: "contentWarning.kind.contentPart",
};

/** The localized warning text for a report: title + counted item labels. */
export function contentWarningText(
  report: UnsupportedContentReport,
  translate: (key: string) => string,
): string {
  const labels = report.items
    .map((item) => {
      const label = translate(CONTENT_WARNING_KIND_KEYS[item.kind]);
      return item.count > 1 ? `${item.count} × ${label}` : label;
    })
    .join(", ");
  return `${translate("contentWarning.title")}: ${translate("contentWarning.body").replace(
    "{items}",
    labels,
  )}`;
}
