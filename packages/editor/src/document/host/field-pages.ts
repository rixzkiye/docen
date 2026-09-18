import { fieldRef } from "@docen/docx";
import type { Node as PMNode } from "@tiptap/pm/model";

/**
 * Map every field atom in the live PM document to the page it renders on —
 * the page context `generateDOCX`'s generated-field pass needs for
 * PAGE/NUMPAGES/PAGEREF/SECTION caches (item 17).
 *
 * The map key is the field atom's document-order index among patchable field
 * atoms (simpleField/complexField, form fields excluded) — exactly the order
 * the docx layout pass numbers them. `physicalPageOf` returns the 0-based
 * physical page (the canvas caret map's `pageOf`); this helper applies the
 * section numbering offsets so a restarted/roman section yields its displayed
 * number.
 */
export interface FieldPageView {
  /** Physical page index → 1-based section, parallel to the page list. */
  sectionOfPage: readonly number[];
  /** Per-section page-number offsets (`computePageNumberOffsets`). */
  pageOffsets: readonly number[];
  /** 0-based physical page for a document position, or null/undefined when
   *  the position is unmapped (headless, a header/footer story, …). */
  physicalPageOf: (pos: number) => number | null | undefined;
}

/** Field index → displayed page number. */
export function collectFieldPages(doc: PMNode, view: FieldPageView): Map<number, number> {
  const pages = new Map<number, number>();
  let index = 0;
  doc.descendants((node, pos) => {
    if (node.type.name !== "inlinePassthrough") return true;
    const data = node.attrs?.data;
    if (typeof data !== "string") return true;
    let ref: ReturnType<typeof fieldRef>;
    try {
      ref = fieldRef(JSON.parse(data) as Record<string, unknown>);
    } catch {
      return true;
    }
    if (!ref?.instruction || ref.kind === "formField") return true;
    const fieldIndex = index++;
    const physical = view.physicalPageOf(pos);
    if (physical == null || physical < 0) return true;
    pages.set(fieldIndex, displayPageOf(view, physical));
    return true;
  });
  return pages;
}

/** Bookmark name → displayed page of the bookmark start — PAGEREF resolves
 *  against the target's page, not the field's own. */
export function collectBookmarkPages(doc: PMNode, view: FieldPageView): Map<string, number> {
  const pages = new Map<string, number>();
  doc.descendants((node, pos) => {
    if (node.type.name !== "inlinePassthrough" && node.type.name !== "passthrough") return true;
    const data = node.attrs?.data;
    if (typeof data !== "string") return true;
    let start: { name?: unknown } | undefined;
    try {
      start = (JSON.parse(data) as { bookmarkStart?: { name?: unknown } }).bookmarkStart;
    } catch {
      return true;
    }
    if (!start || typeof start.name !== "string" || start.name === "") return true;
    const physical = view.physicalPageOf(pos);
    if (physical == null || physical < 0) return true;
    if (!pages.has(start.name)) pages.set(start.name, displayPageOf(view, physical));
    return true;
  });
  return pages;
}

/** A 0-based physical page as the section's displayed number. */
function displayPageOf(view: FieldPageView, physical: number): number {
  const section = view.sectionOfPage[physical] ?? 0;
  return physical + 1 + (view.pageOffsets[section] ?? 0);
}
