import { TextSelection } from "@tiptap/pm/state";

import { Extension } from "../core";

/** A copy of a sectPr with the break type applied. `nextPage` is w:type's
 *  OOXML default, so it deletes the attribute (Word omits it) instead of
 *  writing it. The copy matters: compile's section memo is reference-keyed, so
 *  a mutated-in-place value would keep the stale compiled section. */
function withSectionBreakType(
  properties: unknown,
  type: "nextPage" | "continuous" | "evenPage" | "oddPage",
): Record<string, unknown> {
  const next = { ...(properties as Record<string, unknown> | null | undefined) };
  if (type === "nextPage") delete next.type;
  else next.type = type;
  return next;
}

/**
 * SectionBreak — command extension that marks a paragraph as a section boundary.
 *
 * OOXML sections (sectPr) attach to a section's LAST paragraph's pPr, NOT a
 * standalone node. So this extension provides only the `setSectionBreak`
 * command (stamps sectionProperties on the current paragraph); the paragraph
 * extension carries the sectionProperties/sectionHeaders/sectionFooters attrs,
 * and DocxManager splits/merges sections in compile/resolve by reading them off
 * the paragraph.
 *
 * The final section's sectPr rides on doc.attrs.sectionProperties (it lives at
 * <w:body>'s end in OOXML). Single-section documents have no section-carrying
 * paragraph at all.
 *
 * w:type placement: the type belongs to the section WHOSE sectPr carries it and
 * declares how that section starts relative to the previous one (ECMA-376
 * §17.6.22). A break inserted after section A therefore types the FOLLOWING
 * section's sectPr — the next section-boundary paragraph (which closes that
 * section), or the body-level doc attrs when the new section is the document's
 * last. Stamping A's own sectPr would make A itself start on that parity, the
 * classic off-by-one that loses Word's blank interleave.
 *
 * Break semantics: `setSectionBreak` stamps the current paragraph with the
 * current section's own properties (copied from the following boundary's
 * sectPr, or the body-level final sectPr — so page geometry survives the
 * split) AND inserts a fresh empty paragraph after it (the next section's
 * first paragraph), then types the following sectPr and moves the selection
 * into the new paragraph. The page-plugin's `forcesPageBreakAfter` treats a
 * sectionProperties-bearing paragraph as a page break, so repaginate
 * pushes the new paragraph onto the next page — and the caret follows. This
 * mirrors Word's "Section Break (Next Page)".
 *
 * Split hygiene: pressing Enter inside a section-carrying paragraph must NOT
 * split it. A split would place the new paragraph past the section boundary —
 * forcesPageBreakAfter then pushes it onto the next page (next section), and
 * splitBlock would copy sectionProperties onto it (a second break mark). Word
 * instead inserts a fresh paragraph BEFORE the section's last paragraph, so the
 * new paragraph stays in this section and the stamped paragraph remains last.
 * The Enter shortcut does exactly that; non-section paragraphs fall through to
 * the default Enter unchanged.
 */
export const SectionBreak = Extension.create({
  name: "sectionBreak",
  // Run before the paragraph extension's own Enter handler so the section
  // split-fix wins; non-section paragraphs return false and fall through.
  priority: 1000,

  addCommands() {
    return {
      // Section break: stamp the current paragraph as its section's last
      // paragraph (copying the current section's properties — the section
      // keeps its page geometry), insert a fresh empty paragraph as the next
      // section's first paragraph, write the requested w:type on the
      // FOLLOWING section's sectPr, and move the caret into the new paragraph.
      // A new paragraph is inserted rather than split from the current one so
      // it does NOT inherit the boundary marker (which would make it a section
      // boundary too and break forever). The w:type belongs to the following
      // section (ECMA-376 §17.6.22 — it declares how that section starts):
      // the next section-boundary paragraph after the inserted one, or the
      // body-level doc attrs when the new section is the document's last.
      // "nextPage" (the OOXML default) reflows the next section onto a fresh
      // page; "continuous" keeps it flowing on the same page.
      setSectionBreak:
        (options?: { type?: "nextPage" | "continuous" | "evenPage" | "oddPage" }) =>
        ({ tr, state, dispatch }) => {
          if (!dispatch) return true;
          const { $from } = tr.selection;
          const para = $from.parent;
          // A heading is a paragraph in OOXML — allow a section break on a
          // heading too (e.g. a chapter title that ends its section).
          if (para.type.name !== "paragraph" && para.type.name !== "heading") return false;
          const paraPos = $from.before($from.depth);
          // 1. Insert a fresh empty paragraph (new section's first paragraph)
          //    and move the caret into it; repaginate pushes it to the next
          //    page (nextPage) or flows it on (continuous).
          const paraEnd = paraPos + para.nodeSize;
          tr.insert(paraEnd, state.schema.nodes.paragraph.create());
          // 2. Locate the following section's boundary — the first paragraph
          //    after the inserted one carrying a sectPr. None means the new
          //    section is final and its sectPr is body-level.
          const type = options?.type ?? "nextPage";
          const following: Array<{ pos: number; attrs: Record<string, unknown> }> = [];
          tr.doc.content.forEach((node, offset) => {
            if (following.length > 0 || offset <= paraEnd) return;
            // Only a top-level paragraph is a section boundary — the same
            // walk compileDocument's section split uses.
            if (node.type.name !== "paragraph") return;
            const props = (node.attrs as { sectionProperties?: unknown }).sectionProperties;
            if (props == null) return;
            following.push({ pos: offset, attrs: node.attrs as Record<string, unknown> });
          });
          const target = following[0];
          // 3. The current paragraph closes its section, carrying a copy of
          //    the section's CURRENT properties (the following boundary's
          //    sectPr, or the body-level final sectPr) so the split preserves
          //    page geometry. A paragraph that already carries a sectPr is
          //    already a boundary — leave it untouched.
          if (para.attrs.sectionProperties == null) {
            const carried = target
              ? target.attrs.sectionProperties
              : (state.doc.attrs as { sectionProperties?: unknown }).sectionProperties;
            tr.setNodeMarkup(paraPos, undefined, {
              ...para.attrs,
              sectionProperties: carried ?? {},
            });
          }
          if (target) {
            tr.setNodeMarkup(target.pos, undefined, {
              ...target.attrs,
              sectionProperties: withSectionBreakType(target.attrs.sectionProperties, type),
            });
          } else {
            const docProps = (tr.doc.attrs as { sectionProperties?: unknown }).sectionProperties;
            // An absent body sectPr has no type to clear for the default
            // nextPage — leave it absent rather than materializing an empty
            // one (compile then keeps docen's explicit page defaults).
            const hasType = typeof docProps === "object" && docProps != null && "type" in docProps;
            if (type !== "nextPage" || hasType) {
              tr.setDocAttribute("sectionProperties", withSectionBreakType(docProps, type));
            }
          }
          tr.setSelection(TextSelection.near(tr.doc.resolve(paraEnd + 1)));
          tr.scrollIntoView();
          return true;
        },
    };
  },

  addKeyboardShortcuts() {
    return {
      Enter: ({ editor }) => {
        const { $from } = editor.state.selection;
        const para = $from.parent;
        // Not a section-carrying paragraph → let the default splitBlock run.
        if (para.attrs.sectionProperties == null) return false;
        // Section-carrying paragraph (its section's last paragraph): insert a
        // fresh empty paragraph BEFORE it, not split it. Splitting would place
        // the new paragraph past the section boundary, where forcesPageBreakAfter
        // shoves it onto the next page (next section); splitBlock would also
        // copy sectionProperties onto it (a second break mark). Inserting before
        // keeps the new paragraph in this section and the stamped paragraph
        // remains the section's last — matching Word.
        const paraPos = $from.before($from.depth);
        return editor
          .chain()
          .insertContentAt(paraPos, { type: para.type.name })
          .setTextSelection(paraPos + 1)
          .run();
      },
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    sectionBreak: {
      /** Insert a section break after the current paragraph (Next Page by
       *  default). `type` is written on the FOLLOWING section's sectPr
       *  @w:type (ECMA-376 §17.6.22) — the new section's start mode. */
      setSectionBreak: (options?: {
        type?: "nextPage" | "continuous" | "evenPage" | "oddPage";
      }) => ReturnType;
    };
  }
}
