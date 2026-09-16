import type { RibbonTab } from "../../ui";
import { btn, col, grid, group, menu, opt, parsedItems, split, tabNode } from "./shared";

// Word's Accept / Reject splits (Review → Tracking): the face accepts or
// rejects the selected revision and moves to the next; the drop-down repeats
// that, adds the accept/reject-all sweep, and — Word's order — the
// "…All Changes Shown" sweep scoped to the markup view's author filter.
export const acceptItems = (): string =>
  JSON.stringify([
    { text: opt("accept-and-next"), event: "accept-change" },
    { text: opt("accept-all-changes-shown"), event: "accept-all-changes-shown" },
    { text: opt("accept-all-changes"), event: "accept-all-changes" },
  ]);

export const rejectItems = (): string =>
  JSON.stringify([
    { text: opt("reject-and-next"), event: "reject-change" },
    { text: opt("reject-all-changes-shown"), event: "reject-all-changes-shown" },
    { text: opt("reject-all-changes"), event: "reject-all-changes" },
  ]);

// Word's Display for Review drop-down (Review → Tracking): the four markup
// views. The host re-stamps label + checked to match the live state (the
// #syncEditModeMenu pattern — the static stamp below is the default "simple").
export const displayItems = (): string =>
  JSON.stringify([
    { text: opt("simple-marks"), event: "display-for-review", value: "simple", checked: true },
    { text: opt("all-marks"), event: "display-for-review", value: "all" },
    { text: opt("no-marks"), event: "display-for-review", value: "none" },
    { text: opt("original-marks"), event: "display-for-review", value: "original" },
  ]);

// Word's Specific People drop-down: the document's reviewers plus the "all"
// clearing entry. Author names are document data (w:ins/@w:author), not i18n.
export const reviewerItems = (authors?: readonly string[]): string =>
  JSON.stringify([
    { text: opt("all-reviewers"), event: "review-specific-people", value: "all", checked: true },
    ...(authors ?? []).map((a) => ({ text: a, event: "review-specific-people", value: a })),
  ]);

export const reviewTab = (authors?: readonly string[]): RibbonTab =>
  tabNode("review", [
    group("proofing", [
      btn("spell-check", "spell-check", { size: "large" }),
      col([grid([btn("word-count", "word-count"), btn("search", "thesaurus")])]),
    ]),
    group("accessibility", [btn("checkmark-circle", "check-accessibility", { size: "large" })]),
    group("language", [
      btn("link", "translate", { size: "large" }),
      btn("text-font", "language", { size: "large" }),
    ]),
    group("comments", [
      btn("comment-add", "new-comment", { size: "large" }),
      col([grid([btn("edit", "edit-comment"), btn("close", "delete-comment")])]),
      col([grid([btn("align-left", "previous-comment"), btn("align-right", "next-comment")])]),
      btn("comment", "show-comments", { size: "large" }),
    ]),
    group("tracking", [
      btn("group-objects", "track-changes", { size: "large" }),
      split("accept", "accept-change", parsedItems(acceptItems()), { size: "large" }),
      split("close", "reject-change", parsedItems(rejectItems()), { size: "large" }),
      split("eye", "display-for-review", parsedItems(displayItems()), { size: "large" }),
      menu("people", "review-specific-people", parsedItems(reviewerItems(authors)), {
        size: "large",
      }),
      col([grid([btn("align-left", "previous-change"), btn("align-right", "next-change")])]),
      btn("reviewing-pane", "reviewing-pane", { size: "large" }),
    ]),
    group("compare", [
      btn("group-objects", "compare", { size: "large" }),
      btn("group-objects", "combine", { size: "large" }),
    ]),
    group("protect", [
      btn("protect", "restrict-editing", { size: "large" }),
      btn("protect", "protect-document", { size: "large" }),
    ]),
  ]);
