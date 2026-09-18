import type { ParagraphChild } from "@office-open/docx";

import { Node } from "../core";
import { PRESERVED_RUN_ELEMENTS } from "./coverage";
import type { ParseInlineRule } from "./types";
import { attrNative } from "./utils";

/**
 * RunMarker — inline atom preserving an empty OOXML run element (`EG_RunInnerContent`)
 * that carries document semantics but no text: the field-result placeholders
 * (`w:pgNum`, `w:dayShort`/`w:dayLong`, `w:month*`, `w:year*`), the note/comment
 * auto-marks (`w:footnoteRef`, `w:endnoteRef`, `w:annotationRef`), the note
 * separators (`w:separator`, `w:continuationSeparator`), and Word's pagination
 * hint (`w:lastRenderedPageBreak`).
 *
 * office-open parses these as `{ tag: true }` entries inside a run's `children`
 * (`RunOptions.children`). Before this node the resolve walk dropped every
 * unrecognized child, so a live PAGE field's `w:pgNum` result, a DATE field's
 * `w:dayShort` segments, or a footnote's `w:footnoteRef` mark vanished on
 * DOCX → JSON and the re-export lost the element. The atom stores the element
 * tag and compiles back to a one-child run, so `<w:r><w:pgNum/></w:r>` (or
 * whichever element) survives the round-trip and keeps its OOXML meaning.
 *
 * The `element` attr is restricted to {@link PRESERVED_RUN_ELEMENTS}; anything
 * outside the table still falls through to the resolve walk's drop path (no
 * writer support), which is the tested remainder.
 */

/** The preserved run-element tag union (keys of the coverage table). */
export type RunElementTag = keyof typeof PRESERVED_RUN_ELEMENTS;

/** The first preserved tag an office-open run child carries, or null:
 *  `{ pgNum: true }` → "pgNum", `{ dayShort: true }` → "dayShort", … */
export function runElementTagOf(value: unknown): RunElementTag | null {
  if (!value || typeof value !== "object") return null;
  for (const key of Object.keys(value)) {
    if (key in PRESERVED_RUN_ELEMENTS) return key as RunElementTag;
  }
  return null;
}

export const parseDocxInline: ParseInlineRule = {
  match: (child): child is ParagraphChild => runElementTagOf(child) !== null,
  convert: (child) => {
    const element = runElementTagOf(child);
    return element ? { type: "runMarker", attrs: { element } } : null;
  },
};

export const RunMarker = Node.create({
  name: "runMarker",
  group: "inline",
  inline: true,
  atom: true,

  parseDocxInline,

  addAttributes() {
    return {
      element: attrNative(),
    };
  },

  parseHTML() {
    return [
      {
        tag: "span.docx-run-marker",
        getAttrs: (element) => ({ element: (element as HTMLElement).dataset.element ?? null }),
      },
    ];
  },

  renderHTML({ HTMLAttributes }) {
    return ["span", { class: "docx-run-marker", "data-element": HTMLAttributes.element ?? "" }];
  },
});
