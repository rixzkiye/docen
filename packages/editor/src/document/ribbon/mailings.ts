import type { RibbonTab } from "../../ui";
import { btn, cmd, col, grid, group, opt, parsedItems, split, tabNode } from "./shared";

export const startMergeItems = (): string =>
  JSON.stringify([
    // The kinds the editor can carry: letters merge one copy per recipient,
    // directory streams the records into one document. Envelope/label/e-mail
    // merges need the page-geometry and delivery engines (greyed).
    { text: opt("letters"), value: "letters", event: "start-merge" },
    { text: opt("directory"), value: "directory", event: "start-merge" },
    { text: opt("email"), value: "email", disabled: true },
    { text: opt("envelopes"), value: "envelopes", disabled: true },
    { text: cmd("labels"), value: "labels", disabled: true },
  ]);

export const finishMergeItems = (): string =>
  JSON.stringify([
    // "Edit individual documents" assembles one section-per-recipient docx.
    // Print/e-mail delivery has no pipeline here (greyed).
    { text: opt("edit-docs"), value: "edit", event: "finish-merge" },
    { text: opt("print-docs"), value: "print", disabled: true },
    { text: opt("send-email"), value: "email", disabled: true },
  ]);

export const mailingsTab = (): RibbonTab =>
  tabNode("mailings", [
    group("create", [
      btn("mail", "envelopes", { size: "large" }),
      btn("mail", "labels", { size: "large" }),
    ]),
    group("start-merge", [
      split("document-print", "start-merge", parsedItems(startMergeItems()), { size: "large" }),
      btn("people", "select-recipients", { size: "large" }),
      btn("edit", "edit-recipients", { size: "large" }),
    ]),
    group("write-fields", [
      btn("document-print", "address-block", { size: "large" }),
      btn("comment-add", "greeting-line", { size: "large" }),
      btn("link", "merge-field", { size: "large" }),
      btn("highlight", "highlight-merge", { size: "large", toggle: true }),
    ]),
    group("preview", [
      btn("search", "preview-results", { size: "large" }),
      col([grid([btn("align-left", "first-record"), btn("align-right", "last-record")])]),
    ]),
    group("finish", [
      split("document-print", "finish-merge", parsedItems(finishMergeItems()), { size: "large" }),
    ]),
  ]);
