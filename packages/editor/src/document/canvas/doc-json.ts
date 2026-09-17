// Incremental Tiptap-JSON serialization for the canvas render path.
//
// The render pipeline's inputs are keyed by object identity (compile caches
// per JSON child, the projection caches per compiled child), so the JSON the
// editing bridge hands the renderer must be referentially stable for every
// subtree a transaction did not touch. ProseMirror's persistent tree already
// guarantees that: an unchanged node is the SAME node object after a
// transaction. This module walks the PM doc into Tiptap's JSON shape while
// memoizing per PM node, so one keystroke rewrites one paragraph's subtree
// and every other node reuses its previous JSON object.
//
// Output shape is ProseMirror's own `Node.toJSON` (and `TextNode.toJSON`)
// verbatim — `editor.getJSON()` returns `state.doc.toJSON()`, so consumers
// cannot tell the two apart.

import type { JSONContent } from "@docen/docx";
import type { Node as PMNode } from "@tiptap/pm/model";

/** Per-story identity memo. One instance per editing story; entries die with
 *  the PM nodes they key (WeakMap), so history retention bounds the cache. */
export interface DocJsonCache {
  nodes: WeakMap<PMNode, JSONContent>;
}

export function createDocJsonCache(): DocJsonCache {
  return { nodes: new WeakMap() };
}

/** Serialize a PM node to Tiptap JSON, reusing the previous serialization of
 *  every node object seen before. Equivalent to `node.toJSON()` for all
 *  inputs constructed by Tiptap's schema. */
export function pmNodeToJSON(node: PMNode, cache: DocJsonCache): JSONContent {
  const hit = cache.nodes.get(node);
  if (hit) return hit;
  const obj: JSONContent = { type: node.type.name };
  // PM's `Node.toJSON`: the attrs object rides through by reference when the
  // node declares any attr (even an empty-shaped one) — the projection never
  // mutates it, and identity is what keeps the downstream caches warm.
  for (const _ in node.attrs) {
    obj.attrs = node.attrs as Record<string, unknown>;
    break;
  }
  if (node.content.size) obj.content = node.content.content.map((c) => pmNodeToJSON(c, cache));
  if (node.marks.length)
    obj.marks = node.marks.map((m) => m.toJSON() as NonNullable<JSONContent["marks"]>[number]);
  // `TextNode.toJSON` appends the text; other nodes never carry one.
  if (node.isText && node.text != null) obj.text = node.text;
  cache.nodes.set(node, obj);
  return obj;
}
