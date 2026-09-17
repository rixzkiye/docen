import type { Node as PMNode } from "@tiptap/pm/model";
import { EditorState } from "@tiptap/pm/state";

/**
 * The mathInline atom carrying a math payload at the selection — the equation
 * context tab's trigger (and the equation-insert caret test). A caret hugging
 * the atom (before or after) or a NodeSelection wrapping it counts; `pos` is
 * the atom's document position so the host can re-select it after inserts
 * shift positions. A legacy `inlinePassthrough` carrying `data.math`
 * (documents saved before the single-representation switch) is still
 * recognized.
 */
export function mathAtomAt(state: EditorState): { node: PMNode; pos: number } | null {
  const { $from } = state.selection;
  const before = $from.parent.childAfter($from.parentOffset);
  const after = $from.parent.childBefore($from.parentOffset);
  for (const child of [before, after]) {
    const node = child.node;
    if (!node) continue;
    if (node.type.name === "mathInline") {
      if (node.attrs?.math) return { node, pos: $from.start() + child.offset };
      continue;
    }
    if (node.type.name !== "inlinePassthrough") continue;
    try {
      const data = JSON.parse(String(node.attrs.data ?? "{}")) as { math?: unknown };
      if (data.math) return { node, pos: $from.start() + child.offset };
    } catch {
      /* opaque payload — not a math atom */
    }
  }
  return null;
}
