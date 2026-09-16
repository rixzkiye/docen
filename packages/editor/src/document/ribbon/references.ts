import type { RibbonTab } from "../../ui";
import { btn, cmd, group, opt, parsedItems, split, tabNode } from "./shared";

// References > Add Text: the TOC levels (Word's menu minus the missing-level
// caption line).
export const addTextItems = (): string =>
  JSON.stringify([
    { text: opt("add-text-level-1"), value: "level-1" },
    { text: opt("add-text-level-2"), value: "level-2" },
    { text: opt("add-text-level-3"), value: "level-3" },
    { text: opt("add-text-none"), value: "none" },
  ]);

// References > Table of Contents drop-down (Word lists two auto galleries;
// one auto build covers the same command here): the auto build, the custom
// dialog, and the removal.
export const tocItems = (): string =>
  JSON.stringify([
    { text: opt("toc-auto"), value: "toc", event: "toc" },
    { text: opt("toc-custom"), value: "toc-custom", event: "toc-dialog" },
    { text: opt("toc-remove"), value: "remove-toc", event: "remove-toc" },
  ]);

// References > Update Table: Word's page-numbers-only vs whole-table pass.
export const updateTocItems = (): string =>
  JSON.stringify([
    { text: opt("toc-update-page"), value: "update-toc-page", event: "update-toc-page" },
    { text: opt("toc-update-all"), value: "update-toc", event: "update-toc" },
  ]);

export const footnoteItems = (): string =>
  JSON.stringify([
    { text: cmd("insert-footnote"), value: "footnote" },
    { text: opt("endnote"), value: "endnote" },
    { text: opt("next-footnote"), value: "next" },
    { text: opt("previous-footnote"), value: "prev" },
  ]);

export const referencesTab = (): RibbonTab =>
  tabNode("references", [
    group("toc", [
      split("toc", "toc", parsedItems(tocItems()), { size: "large" }),
      split("multilevel", "add-text", parsedItems(addTextItems()), { size: "large" }),
      split("sync", "update-toc", parsedItems(updateTocItems()), { size: "large" }),
    ]),
    group(
      "footnotes",
      [split("footnote", "insert-footnote", parsedItems(footnoteItems()), { size: "large" })],
      "note-settings-dialog",
    ),
    group("citations", [
      btn("comment-add", "insert-citation", { size: "large" }),
      btn("people", "manage-sources", { size: "large" }),
      btn("document-print", "bibliography", { size: "large" }),
    ]),
    group("captions", [
      btn("comment-add", "insert-caption", { size: "large" }),
      btn("document-print", "table-of-figures", { size: "large" }),
      btn("sync", "update-figures", { size: "large" }),
      btn("link", "cross-reference", { size: "large" }),
    ]),
    group("index", [
      btn("comment-add", "mark-entry", { size: "large" }),
      btn("document-print", "insert-index", { size: "large" }),
      btn("sync", "update-index", { size: "large" }),
    ]),
    group("toa", [
      btn("comment-add", "mark-citation", { size: "large" }),
      btn("document-print", "insert-toa", { size: "large" }),
      btn("sync", "update-toa", { size: "large" }),
    ]),
  ]);
