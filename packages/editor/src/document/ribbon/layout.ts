import type { RibbonTab } from "../../ui";
import {
  alignObjectsItems,
  btn,
  cmd,
  col,
  group,
  groupItems,
  menu,
  opt,
  parsedItems,
  positionItems,
  rotateItems,
  row,
  spacingItems,
  split,
  tabNode,
  wrapItems,
} from "./shared";

export const marginsItems = (): string =>
  JSON.stringify([
    { text: opt("normal-margin"), value: "normal" },
    { text: opt("narrow"), value: "narrow" },
    { text: opt("moderate"), value: "moderate" },
    { text: opt("wide"), value: "wide" },
    // Custom Margins opens the Page Setup dialog (host-handled).
    { text: opt("custom-margins"), value: "custom" },
  ]);

export const orientationItems = (): string =>
  JSON.stringify([
    { text: opt("portrait"), value: "portrait" },
    { text: opt("landscape"), value: "landscape" },
  ]);

export const sizePaperItems = (): string =>
  JSON.stringify([
    { text: opt("letter"), value: "letter" },
    { text: opt("legal"), value: "legal" },
    { text: opt("tabloid"), value: "tabloid" },
    { text: opt("a3"), value: "a3" },
    { text: opt("a4"), value: "a4" },
    { text: opt("a5"), value: "a5" },
    { text: opt("b5"), value: "b5" },
    { text: opt("statement"), value: "statement" },
    { text: opt("executive"), value: "executive" },
    // More Paper Sizes opens the same Page Setup dialog (host-handled).
    { text: opt("more-sizes"), value: "more" },
  ]);

export const columnsItems = (): string =>
  JSON.stringify([
    { text: opt("one-col"), value: "1" },
    { text: opt("two-col"), value: "2" },
    { text: opt("three-col"), value: "3" },
    // More Columns opens the Columns dialog (host-handled).
    { text: opt("more-columns"), value: "more" },
  ]);

export const breaksItems = (): string =>
  JSON.stringify([
    { text: cmd("page-break"), value: "page-break", event: "page-break" },
    { text: opt("column-break"), value: "column-break", event: "column-break" },
    // Text Wrapping opens Word's layout-options dialog (not built yet).
    { text: opt("text-wrapping"), value: "text-wrapping", event: "text-wrapping", disabled: true },
    // Word's four section-break types; Even/Odd Page need engine support for
    // starting sections on even/odd pages (with blank interleaves) — greyed
    // until then.
    { text: opt("next-page-section"), value: "section-break-next", event: "section-break-next" },
    {
      text: opt("continuous-section"),
      value: "section-break-continuous",
      event: "section-break-continuous",
    },
    { text: opt("even-page-section"), value: "even", disabled: true },
    { text: opt("odd-page-section"), value: "odd", disabled: true },
  ]);

// Word's Line Numbers menu: the numbering mode writes w:lnNumType on the
// current section; the trailing options entry opens Word's Line Numbering
// dialog (start-at / count-by / distance).
export const lineNumbersItems = (): string =>
  JSON.stringify([
    { text: opt("no-line-numbers"), value: "none", event: "line-numbers" },
    { text: opt("continuous-line-numbers"), value: "continuous", event: "line-numbers" },
    { text: opt("restart-each-page"), value: "newPage", event: "line-numbers" },
    { text: opt("restart-each-section"), value: "newSection", event: "line-numbers" },
    { text: opt("line-numbering-options"), value: "options", event: "line-numbers" },
  ]);

// Word's Hyphenation menu: None, Manual, Automatic, Hyphenation Options.
export const hyphenationItems = (): string =>
  JSON.stringify([
    { text: opt("hyphenation-none"), value: "none" },
    { text: opt("hyphenation-manual"), value: "manual" },
    { text: opt("hyphenation-auto"), value: "auto" },
    { text: opt("hyphenation-options"), value: "options" },
  ]);

export const indentItems = (): string =>
  JSON.stringify([
    { text: opt("increase-indent"), value: "increase", event: "indent-increase" },
    { text: opt("decrease-indent"), value: "decrease", event: "indent-decrease" },
  ]);

export const layoutTab = (): RibbonTab =>
  tabNode("layout", [
    group(
      "page-setup",
      [
        split("margins", "margins", parsedItems(marginsItems()), { size: "large" }),
        split("orientation", "orientation", parsedItems(orientationItems()), { size: "large" }),
        split("page-size", "page-size", parsedItems(sizePaperItems()), { size: "large" }),
        split("columns", "columns", parsedItems(columnsItems()), { size: "large" }),
        split("page-break", "page-break", parsedItems(breaksItems()), {
          size: "large",
          label: cmd("breaks"),
        }),
        split("number-symbol", "line-numbers", parsedItems(lineNumbersItems()), { size: "large" }),
        split("hyphenation", "hyphenation", parsedItems(hyphenationItems()), {
          size: "large",
          label: cmd("hyphenation"),
        }),
      ],
      "page-setup-dialog",
    ),
    group(
      "paragraph",
      [
        split("indent-increase", "indent-increase", parsedItems(indentItems()), { size: "large" }),
        split("line-spacing", "line-spacing", parsedItems(spacingItems()), { size: "large" }),
      ],
      "paragraph-dialog",
    ),
    group("arrange", [
      col([
        row([
          menu("orientation", "position", parsedItems(positionItems())),
          menu("wrap", "wrap", parsedItems(wrapItems())),
        ]),
        row([btn("orientation", "bring-forward"), btn("orientation", "send-backward")]),
      ]),
      menu("align-left", "align-objects", parsedItems(alignObjectsItems()), { size: "large" }),
      menu("group-objects", "drawing-group", parsedItems(groupItems()), { size: "large" }),
      menu("rotate", "rotate", parsedItems(rotateItems()), { size: "large" }),
    ]),
  ]);
