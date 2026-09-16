import type { Editor } from "@docen/docx/core";

import type { OutlineItem } from "../components/outline";

/**
 * Filter an outline tree so only headings up to `maxLevel` are visible
 * (Word's "Show Level" dropdown: Level 1 through Level 9, or All Levels).
 */
export function filterOutlineByLevel(
  items: readonly OutlineItem[],
  maxLevel: number,
): OutlineItem[] {
  if (maxLevel <= 0 || maxLevel >= 9) return [...items];

  const result: OutlineItem[] = [];
  for (const item of items) {
    const level = item.level ?? 1;
    if (level <= maxLevel) {
      const filteredChildren = item.children
        ? filterOutlineByLevel(item.children, maxLevel)
        : undefined;
      result.push({
        ...item,
        children: filteredChildren && filteredChildren.length > 0 ? filteredChildren : undefined,
      });
    }
  }
  return result;
}

/**
 * Promote the heading under caret (Word's Promote / Alt+Shift+Left):
 * Decreases the heading level (e.g. Heading 3 → Heading 2, Heading 2 → Heading 1).
 */
export function promoteHeadingAtCaret(editor: Editor): boolean {
  const { $from } = editor.state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const node = $from.node(d);
    if (node.type.name === "paragraph" || node.type.name === "heading") {
      const style = String((node.attrs as { style?: string })?.style ?? "");
      const match = /^Heading([1-9])$/.exec(style);
      if (match) {
        const lvl = parseInt(match[1]!, 10);
        const nextStyle = lvl > 1 ? `Heading${lvl - 1}` : "Normal";
        const pos = $from.before(d);
        const tr = editor.state.tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          style: nextStyle,
        });
        editor.view.dispatch(tr);
        return true;
      }
    }
  }
  return false;
}

/**
 * Demote the heading under caret (Word's Demote / Alt+Shift+Right):
 * Increases the heading level (e.g. Normal → Heading 1, Heading 1 → Heading 2, Heading 8 → Heading 9).
 */
export function demoteHeadingAtCaret(editor: Editor): boolean {
  const { $from } = editor.state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const node = $from.node(d);
    if (node.type.name === "paragraph" || node.type.name === "heading") {
      const style = String((node.attrs as { style?: string })?.style ?? "");
      const match = /^Heading([1-9])$/.exec(style);
      const lvl = match ? parseInt(match[1]!, 10) : 0;
      if (lvl < 9) {
        const nextStyle = `Heading${lvl + 1}`;
        const pos = $from.before(d);
        const tr = editor.state.tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          style: nextStyle,
        });
        editor.view.dispatch(tr);
        return true;
      }
    }
  }
  return false;
}

/**
 * Move the current block up (Word's Alt+Shift+Up):
 * Swaps the current block with its previous sibling.
 */
export function moveBlockUp(editor: Editor): boolean {
  const { $from } = editor.state.selection;
  const index = $from.index(0);
  if (index <= 0) return false;

  const doc = editor.state.doc;
  const prevNode = doc.child(index - 1);
  const currentNode = doc.child(index);

  let prevPos = 0;
  for (let i = 0; i < index - 1; i++) {
    prevPos += doc.child(i).nodeSize;
  }
  const curPos = prevPos + prevNode.nodeSize;

  const tr = editor.state.tr;
  tr.delete(curPos, curPos + currentNode.nodeSize);
  tr.insert(prevPos, currentNode);
  editor.view.dispatch(tr);
  return true;
}

/**
 * Move the current block down (Word's Alt+Shift+Down):
 * Swaps the current block with its next sibling.
 */
export function moveBlockDown(editor: Editor): boolean {
  const { $from } = editor.state.selection;
  const index = $from.index(0);
  const doc = editor.state.doc;
  if (index >= doc.childCount - 1) return false;

  const currentNode = doc.child(index);
  const nextNode = doc.child(index + 1);

  let curPos = 0;
  for (let i = 0; i < index; i++) {
    curPos += doc.child(i).nodeSize;
  }
  const nextPos = curPos + currentNode.nodeSize;

  const tr = editor.state.tr;
  tr.delete(nextPos, nextPos + nextNode.nodeSize);
  tr.insert(curPos, nextNode);
  editor.view.dispatch(tr);
  return true;
}

/**
 * Select text with similar formatting (Word's "Select All Text With Similar Formatting"):
 * Matches marks (bold, italic, font, color, etc.) of the current selection across the document.
 */
export function selectSimilarFormatting(editor: Editor): boolean {
  const { from } = editor.state.selection;
  const doc = editor.state.doc;
  const $from = doc.resolve(from);
  const currentMarks = $from.marks();
  const currentMarkNames = new Set(currentMarks.map((m) => m.type.name));

  let firstMatch = -1;
  let lastMatch = -1;

  doc.descendants((node, pos) => {
    if (!node.isText) return;
    const nodeMarkNames = new Set(node.marks.map((m) => m.type.name));
    let matches = true;
    for (const name of currentMarkNames) {
      if (!nodeMarkNames.has(name)) {
        matches = false;
        break;
      }
    }
    if (matches && currentMarkNames.size === nodeMarkNames.size) {
      if (firstMatch === -1) firstMatch = pos;
      lastMatch = pos + node.nodeSize;
    }
  });

  if (firstMatch !== -1 && lastMatch > firstMatch) {
    const TextSelection = (editor.state.selection as any).constructor;
    if (TextSelection && typeof TextSelection.create === "function") {
      const sel = TextSelection.create(doc, firstMatch, lastMatch);
      editor.view.dispatch(editor.state.tr.setSelection(sel));
      return true;
    }
  }
  return false;
}
