/**
 * Template (.dotx) attachment and management domain.
 * Supports extracting styles, numbering, docDefaults, and DrawingML themes
 * from .dotx binary templates and applying them to existing documents
 * without overwriting document body content, with Word-compatible linkStyles support.
 */

import { parseDOCXSync, type StylesOptions } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { unzipSync } from "@office-open/core";

import type { ThemeColorScheme, ThemeDefinition, ThemeFontScheme } from "./commands/themes";

export interface DotxTemplatePackage {
  styles?: StylesOptions;
  numbering?: unknown;
  docDefaults?: unknown;
  theme?: ThemeDefinition;
  settings?: Record<string, unknown>;
  rawParts?: { path: string; data: Uint8Array }[];
}

export interface AttachTemplateOptions {
  /**
   * Word's settings.linkStyles: when true, styles from the template are
   * automatically updated in the document, and the document records linkStyles: true.
   */
  autoUpdateStyles?: boolean;
}

/**
 * Parse an Office theme XML (word/theme/theme1.xml) to extract color scheme
 * and font scheme definitions.
 */
export function parseThemeXml(xml: string): ThemeDefinition | undefined {
  if (!xml) return undefined;

  // Extract color scheme name and colors
  const clrSchemeMatch = xml.match(/<a:clrScheme[^>]*name="([^"]*)"/i);
  const colorSchemeName = clrSchemeMatch?.[1] ?? "Template Theme";

  const extractColor = (tag: string, defaultHex: string): string => {
    // Check for srgbClr val="RRGGBB" or sysClr lastClr="RRGGBB"
    const tagRegex = new RegExp(
      `<a:${tag}[^>]*>[\\s\\S]*?val="([0-9a-fA-F]{6})"[\\s\\S]*?<\\/a:${tag}>`,
      "i",
    );
    const m = xml.match(tagRegex);
    if (m?.[1]) return m[1].toLowerCase();

    const sysRegex = new RegExp(
      `<a:${tag}[^>]*>[\\s\\S]*?lastClr="([0-9a-fA-F]{6})"[\\s\\S]*?<\\/a:${tag}>`,
      "i",
    );
    const sm = xml.match(sysRegex);
    if (sm?.[1]) return sm[1].toLowerCase();

    return defaultHex.toLowerCase();
  };

  const colors: ThemeColorScheme = {
    id: "template-theme",
    name: colorSchemeName,
    dk1: extractColor("dk1", "000000"),
    lt1: extractColor("lt1", "ffffff"),
    dk2: extractColor("dk2", "1f497d"),
    lt2: extractColor("lt2", "eeece1"),
    accent1: extractColor("accent1", "4f81bd"),
    accent2: extractColor("accent2", "c0504d"),
    accent3: extractColor("accent3", "9bbb59"),
    accent4: extractColor("accent4", "8064a2"),
    accent5: extractColor("accent5", "4bacc6"),
    accent6: extractColor("accent6", "f79646"),
    hlink: extractColor("hlink", "0000ff"),
    folHlink: extractColor("folHlink", "800080"),
  };

  // Extract font scheme
  const fontSchemeMatch = xml.match(/<a:fontScheme[^>]*name="([^"]*)"/i);
  const fontSchemeName = fontSchemeMatch?.[1] ?? "Template Fonts";

  const extractTypeface = (tag: string, defaultFont: string): string => {
    const reg = new RegExp(`<a:${tag}[^>]*>[\\s\\S]*?<a:latin[^>]*typeface="([^"]+)"`, "i");
    const m = xml.match(reg);
    return m?.[1] ?? defaultFont;
  };

  const fonts: ThemeFontScheme = {
    id: "template-fonts",
    name: fontSchemeName,
    majorFont: extractTypeface("majorFont", "Calibri Light"),
    minorFont: extractTypeface("minorFont", "Calibri"),
  };

  return {
    id: "template-theme",
    name: colorSchemeName,
    colors,
    fonts,
  };
}

/**
 * Parse a .dotx template file to extract styles, numbering, docDefaults, and themes.
 */
export function parseDotxTemplate(data: Uint8Array | ArrayBuffer): DotxTemplatePackage {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const parts = unzipSync(bytes);

  // Parse document model using docen converter
  const docJson = parseDOCXSync(bytes);
  const attrs = (docJson.attrs ?? {}) as {
    styles?: StylesOptions;
    numbering?: unknown;
    docDefaults?: unknown;
    documentExtras?: Record<string, unknown>;
  };

  let theme: ThemeDefinition | undefined;
  const themePart = parts["word/theme/theme1.xml"];
  if (themePart) {
    const xml = new TextDecoder().decode(themePart);
    theme = parseThemeXml(xml);
  }

  const rawParts: { path: string; data: Uint8Array }[] = [];
  for (const [path, data] of Object.entries(parts)) {
    if (path.startsWith("word/theme/") || path.startsWith("word/media/")) {
      rawParts.push({ path, data });
    }
  }

  return {
    styles: attrs.styles,
    numbering: attrs.numbering,
    docDefaults: attrs.docDefaults,
    theme,
    settings: attrs.documentExtras?.settings as Record<string, unknown> | undefined,
    rawParts,
  };
}

/**
 * Merge styles from a template into existing document styles.
 */
export function mergeTemplateStyles(
  current: StylesOptions | undefined,
  template: StylesOptions | undefined,
): StylesOptions {
  if (!template) return current ? { ...current } : {};
  if (!current) return { ...template };

  const merged: StylesOptions = { ...current };

  // 1. Paragraph styles
  if (template.paragraphStyles) {
    const currentList = [...(current.paragraphStyles ?? [])];
    for (const tplStyle of template.paragraphStyles) {
      const idx = currentList.findIndex(
        (s) => s.id === tplStyle.id || (tplStyle.name && s.name === tplStyle.name),
      );
      if (idx >= 0) {
        currentList[idx] = { ...currentList[idx], ...tplStyle };
      } else {
        currentList.push(tplStyle);
      }
    }
    merged.paragraphStyles = currentList;
  }

  // 2. Character styles
  if (template.characterStyles) {
    const currentChars = [...(current.characterStyles ?? [])];
    for (const tplChar of template.characterStyles) {
      const idx = currentChars.findIndex(
        (s) => s.id === tplChar.id || (tplChar.name && s.name === tplChar.name),
      );
      if (idx >= 0) {
        currentChars[idx] = { ...currentChars[idx], ...tplChar };
      } else {
        currentChars.push(tplChar);
      }
    }
    merged.characterStyles = currentChars;
  }

  // 3. Table styles
  if (template.tableStyles) {
    const currentTables = [...(current.tableStyles ?? [])];
    for (const tplTable of template.tableStyles) {
      const idx = currentTables.findIndex((s) => s.id === tplTable.id);
      if (idx >= 0) {
        currentTables[idx] = { ...currentTables[idx], ...tplTable };
      } else {
        currentTables.push(tplTable);
      }
    }
    merged.tableStyles = currentTables;
  }

  // 4. Default built-in styles (heading1-9, title, etc.)
  if (template.default) {
    merged.default = {
      ...current.default,
      ...template.default,
    };
  }

  // 5. docDefaultsXml
  if (template.docDefaultsXml) {
    merged.docDefaultsXml = template.docDefaultsXml;
  }
  const tplDefaults = (template as Record<string, unknown>).docDefaults;
  if (tplDefaults) {
    (merged as Record<string, unknown>).docDefaults = {
      ...((current as Record<string, unknown>)?.docDefaults as Record<string, unknown> | undefined),
      ...(tplDefaults as Record<string, unknown>),
    };
  }

  return merged;
}

/**
 * Attach a template to a document, updating styles, numbering, docDefaults, and themes
 * while preserving the document's body content.
 */
export function attachTemplate(
  editor: Editor,
  data: Uint8Array | ArrayBuffer,
  options: AttachTemplateOptions = {},
): DotxTemplatePackage {
  const pkg = parseDotxTemplate(data);

  const currentAttrs = (editor.state.doc.attrs ?? {}) as {
    styles?: StylesOptions;
    numbering?: unknown;
    docDefaults?: unknown;
    documentExtras?: Record<string, unknown>;
  };

  const mergedStyles = mergeTemplateStyles(currentAttrs.styles, pkg.styles);

  const extras = { ...currentAttrs.documentExtras };
  const settings = { ...(extras.settings as Record<string, unknown> | undefined) };

  if (options.autoUpdateStyles != null) {
    settings.linkStyles = Boolean(options.autoUpdateStyles);
  }

  if (pkg.theme) {
    settings.theme = {
      id: pkg.theme.id,
      name: pkg.theme.name,
      colors: pkg.theme.colors,
      fonts: pkg.theme.fonts,
    };
  }

  const tr = editor.state.tr;
  tr.setDocAttribute("styles", mergedStyles);

  if (pkg.numbering) {
    tr.setDocAttribute("numbering", pkg.numbering);
  }

  if (pkg.docDefaults) {
    tr.setDocAttribute("docDefaults", pkg.docDefaults);
  }

  tr.setDocAttribute("documentExtras", {
    ...extras,
    settings,
  });

  editor.view.dispatch(tr);
  return pkg;
}
