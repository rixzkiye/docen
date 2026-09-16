import type { StylesOptions } from "@docen/docx";

import { resolveLang, type RibbonTab } from "../../ui";
import { FONT_NAMES, FONT_SIZES_CN, FONT_SIZES_PT, UNDERLINE_STYLES } from "../font-lists";
import {
  borderItems,
  btn,
  col,
  combo,
  group,
  opt,
  parsedItems,
  picker,
  row,
  sep,
  spacingItems,
  split,
  styleGalleryItems,
  tabNode,
} from "./shared";

// --- Option sets (menu/combobox items) ---------------------------------------

export const fontItems = (): string => JSON.stringify(FONT_NAMES.map((text) => ({ text })));

// Font-size options: a zh locale lists the Chinese names ("小四 (12)") above
// the point sizes; other locales show point sizes only. The emitted `value` is
// always the pt string, so the two lists stay compatible across locales.
export const sizeItems = (): string => {
  const zh = resolveLang().toLowerCase().startsWith("zh");
  const cn = zh
    ? FONT_SIZES_CN.map(([name, pt]) => ({ text: `${name} (${pt})`, value: String(pt) }))
    : [];
  const pt = FONT_SIZES_PT.map((p) => ({ text: String(p), value: String(p) }));
  return JSON.stringify([...cn, ...pt]);
};

export const pasteItems = (): string =>
  JSON.stringify([
    { text: opt("paste"), value: "paste" },
    { text: opt("paste-special"), value: "paste-special" },
    { text: opt("keep-text-only"), value: "keep-text-only" },
  ]);

// Word's Underline split menu — the ST_Underline patterns plus the clear
// entry ("none" and the pattern list live in font-lists.ts, shared with the
// Font dialog's Underline style dropdown).
export const underlineItems = (): string =>
  JSON.stringify([
    { text: opt("none"), event: "underline-style", value: "none" },
    ...UNDERLINE_STYLES.map(([value, key]) => ({
      text: opt(key),
      event: "underline-style",
      value,
    })),
  ]);

export const caseItems = (): string =>
  JSON.stringify([
    { text: opt("sentence-case"), value: "sentence" },
    { text: opt("lowercase"), value: "lower" },
    { text: opt("uppercase"), value: "upper" },
    { text: opt("capitalize"), value: "capitalize" },
    { text: opt("toggle-case"), value: "toggle" },
  ]);

// Word's Chinese Layout (中文版式) drop-down in the Paragraph group — both
// entries open the shared two-lines-in-one dialog (the dialog's bracket
// checkbox covers 合并字符's no-bracket form).
export const chineseLayoutItems = (): string =>
  JSON.stringify([
    { text: opt("combine-characters"), event: "two-lines-in-one", value: "combine-characters" },
    { text: opt("two-lines-in-one"), event: "two-lines-in-one", value: "two-lines-in-one" },
  ]);

export const bulletItems = (): string =>
  JSON.stringify([
    { text: opt("bullet"), value: "bullet" },
    { text: opt("circle"), value: "circle" },
    { text: opt("square"), value: "square" },
    // Word's Change List Level — demote one level (the menu stand-in for Tab).
    { text: opt("change-list-level"), value: "in", event: "multilevel-list" },
  ]);

export const numberItems = (): string =>
  JSON.stringify([
    { text: opt("decimal"), value: "decimal" },
    { text: opt("lower-alpha"), value: "lower-alpha" },
    { text: opt("lower-roman"), value: "lower-roman" },
    { text: opt("change-list-level"), value: "in", event: "multilevel-list" },
  ]);

// Word's multilevel List Library — each entry names its marker shape (the
// gallery thumbnails' text) and applies that preset; the last entry opens
// Define New Multilevel List.
export const multilevelItems = (): string =>
  JSON.stringify([
    { text: "1., 1.1., 1.1.1.", event: "multilevel-list", value: "preset:cascade" },
    { text: "1), 1.1), 1.1.1)", event: "multilevel-list", value: "preset:cascade-paren" },
    { text: "1., a., i.", event: "multilevel-list", value: "preset:hybrid" },
    { text: "1), a), i)", event: "multilevel-list", value: "preset:hybrid-paren" },
    { text: "一、（一）1.", event: "multilevel-list", value: "preset:cjk" },
    { text: "第X章 第X节", event: "multilevel-list", value: "preset:cjk-chapter" },
    { text: opt("define-new-multilevel-list"), event: "define-new-list" },
  ]);

export const findItems = (): string =>
  JSON.stringify([
    { text: opt("find"), value: "find" },
    { text: opt("go-to"), value: "go-to" },
  ]);

export const selectItems = (): string =>
  JSON.stringify([
    { text: opt("select-all"), value: "all" },
    // Need a canvas selection model for objects / similar-format picks.
    { text: opt("select-objects"), value: "objects", disabled: true },
    { text: opt("select-similar"), value: "similar", disabled: true },
  ]);

export const homeTab = (styles?: StylesOptions | null): RibbonTab =>
  tabNode("home", [
    group(
      "clipboard",
      [
        split("paste", "paste", parsedItems(pasteItems()), { size: "large" }),
        col([
          btn("cut", "cut", { iconOnly: true }),
          btn("copy", "copy", { iconOnly: true }),
          btn("format-painter", "format-painter", { iconOnly: true, toggle: true }),
        ]),
      ],
      "clipboard-dialog",
    ),
    group(
      "font",
      [
        col([
          row([
            combo("font-name", "Microsoft YaHei", parsedItems(fontItems()), {
              source: "local-fonts",
            }),
            combo("font-size", "14", parsedItems(sizeItems()), { comboboxSize: "short" }),
            btn("font-size", "grow-font", { iconOnly: true }),
            btn("font-size", "shrink-font", { iconOnly: true }),
            split("case", "change-case", parsedItems(caseItems()), { iconOnly: true }),
            btn("clear-format", "clear-format", { iconOnly: true }),
          ]),
          row([
            btn("bold", "bold", { iconOnly: true, toggle: true }),
            btn("italic", "italic", { iconOnly: true, toggle: true }),
            split("underline", "underline", parsedItems(underlineItems()), {
              iconOnly: true,
              toggle: true,
            }),
            btn("strike", "strike", { iconOnly: true, toggle: true }),
            btn("superscript", "superscript", { iconOnly: true, toggle: true }),
            btn("subscript", "subscript", { iconOnly: true, toggle: true }),
            sep(),
            btn("phonetic-guide", "phonetic-guide", { iconOnly: true }),
            sep(),
            picker("highlight", "highlight", "FFFF00", { palette: "highlight" }),
            picker("font-color", "font-color", "000000"),
          ]),
        ]),
      ],
      "font-dialog",
    ),
    group(
      "paragraph",
      [
        col([
          row([
            split("list", "bullet-list", parsedItems(bulletItems()), {
              iconOnly: true,
              toggle: true,
            }),
            split("numbering", "ordered-list", parsedItems(numberItems()), {
              iconOnly: true,
              toggle: true,
            }),
            split("multilevel", "multilevel-list", parsedItems(multilevelItems()), {
              iconOnly: true,
            }),
            btn("indent-decrease", "indent-decrease", { iconOnly: true }),
            btn("indent-increase", "indent-increase", { iconOnly: true }),
            btn("direction-ltr", "direction-ltr", { iconOnly: true }),
            btn("direction-rtl", "direction-rtl", { iconOnly: true }),
            split("two-in-one", "two-lines-in-one", parsedItems(chineseLayoutItems()), {
              iconOnly: true,
            }),
            btn("sort", "sort", { iconOnly: true }),
            btn("show-marks", "show-marks", { iconOnly: true, toggle: true }),
          ]),
          row([
            btn("align-left", "align-left", { iconOnly: true, toggle: true }),
            btn("align-center", "align-center", { iconOnly: true, toggle: true }),
            btn("align-right", "align-right", { iconOnly: true, toggle: true }),
            btn("justify", "justify", { iconOnly: true, toggle: true }),
            btn("align-distribute", "justify-distribute", { iconOnly: true, toggle: true }),
            sep(),
            split("line-spacing", "line-spacing", parsedItems(spacingItems()), { iconOnly: true }),
            picker("shading", "shading", "FFFF00"),
            split("border", "border", parsedItems(borderItems()), { iconOnly: true }),
          ]),
        ]),
      ],
      "paragraph-dialog",
    ),
    group(
      "styles",
      [
        col([
          {
            type: "gallery",
            event: "style",
            items: styleGalleryItems(styles),
            visibleCount: 3,
          },
        ]),
      ],
      "styles-pane",
    ),
    // Markdown input mode — a typing-mode flag, not a document command
    // (Word has no counterpart; this is docen's own input surface).
    group("markdown", [btn("markdown-input", "markdown-input", { size: "large", toggle: true })]),
    group(
      "editing",
      [
        split("search", "search", parsedItems(findItems()), { size: "large" }),
        col([
          btn("replace", "replace", { iconOnly: true }),
          split("board", "select", parsedItems(selectItems()), { iconOnly: true }),
        ]),
      ],
      "find-dialog",
    ),
  ]);
