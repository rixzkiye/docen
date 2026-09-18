import { detectHeadingLevel, type StylesOptions } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { type Node as PMNode } from "@tiptap/pm/model";
import { TextSelection, type Transaction } from "@tiptap/pm/state";

import type { OutlineItem } from "../components/outline";

/**
 * Detect the heading level of a ProseMirror node (1-9, or null if not a heading).
 */
export function detectNodeHeadingLevel(node: PMNode, styles?: StylesOptions): number | null {
  if (node.type.name !== "paragraph" && node.type.name !== "heading") return null;
  const lvl = detectHeadingLevel(
    {
      heading: (node.attrs.heading as string) || undefined,
      style: (node.attrs.style as string) || undefined,
      outlineLevel: node.attrs.outlineLevel as number | undefined,
    },
    styles,
  );
  return lvl ?? null;
}

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
 * Find the range of a heading section (the heading itself and all subordinate body
 * paragraphs and lower-level subheadings, until the next heading of equal or higher level).
 */
export function findHeadingSectionRange(
  doc: PMNode,
  headingIndex: number,
  styles?: StylesOptions,
): { startIndex: number; endIndex: number; startPos: number; endPos: number } {
  if (headingIndex < 0 || headingIndex >= doc.childCount) {
    return { startIndex: headingIndex, endIndex: headingIndex, startPos: 0, endPos: 0 };
  }
  const headingNode = doc.child(headingIndex);
  const lvl = detectNodeHeadingLevel(headingNode, styles);
  let startPos = 0;
  for (let i = 0; i < headingIndex; i++) {
    startPos += doc.child(i).nodeSize;
  }
  if (lvl == null) {
    return {
      startIndex: headingIndex,
      endIndex: headingIndex + 1,
      startPos,
      endPos: startPos + headingNode.nodeSize,
    };
  }
  let endIndex = headingIndex + 1;
  let curPos = startPos + headingNode.nodeSize;
  while (endIndex < doc.childCount) {
    const child = doc.child(endIndex);
    const childLvl = detectNodeHeadingLevel(child, styles);
    if (childLvl != null && childLvl <= lvl) {
      break;
    }
    curPos += child.nodeSize;
    endIndex++;
  }
  return {
    startIndex: headingIndex,
    endIndex,
    startPos,
    endPos: curPos,
  };
}

/**
 * Move an entire heading section (heading + subordinate items) before or after
 * another section in the document.
 */
export function moveHeadingSection(
  editor: Editor,
  sourceHeadingIndex: number,
  targetHeadingIndex: number,
  position: "before" | "after",
  customTr?: Transaction,
): boolean {
  const doc = editor.state.doc;
  const styles = (doc.attrs as { styles?: StylesOptions }).styles;
  const srcRange = findHeadingSectionRange(doc, sourceHeadingIndex, styles);
  if (srcRange.startIndex === srcRange.endIndex) return false;
  const tgtRange = findHeadingSectionRange(doc, targetHeadingIndex, styles);
  if (tgtRange.startIndex === tgtRange.endIndex) return false;
  if (sourceHeadingIndex >= tgtRange.startIndex && sourceHeadingIndex < tgtRange.endIndex) {
    return false;
  }

  const tr = customTr ?? editor.state.tr;
  const blocks: PMNode[] = [];
  for (let i = 0; i < doc.childCount; i++) {
    blocks.push(doc.child(i));
  }

  const count = srcRange.endIndex - srcRange.startIndex;
  const moved = blocks.splice(srcRange.startIndex, count);

  let insertIdx = position === "before" ? tgtRange.startIndex : tgtRange.endIndex;
  if (srcRange.startIndex < insertIdx) {
    insertIdx -= count;
  }

  blocks.splice(insertIdx, 0, ...moved);
  const newDoc = tr.doc.type.create(tr.doc.attrs, blocks);
  tr.replaceWith(0, tr.doc.content.size, newDoc.content);

  let newPos = 1;
  for (let i = 0; i < insertIdx; i++) {
    newPos += blocks[i]!.nodeSize;
  }
  tr.setSelection(TextSelection.near(tr.doc.resolve(newPos)));

  if (!customTr) editor.view.dispatch(tr);
  return true;
}

/**
 * Promote the heading under caret (Word's Promote / Alt+Shift+Left):
 * Decreases the heading level (e.g. Heading 3 → Heading 2, Heading 2 → Heading 1).
 */
export function promoteHeadingAtCaret(editor: Editor, customTr?: Transaction): boolean {
  const tr = customTr ?? editor.state.tr;
  const { $from } = tr.selection;
  for (let d = $from.depth; d > 0; d--) {
    const node = $from.node(d);
    if (node.type.name === "paragraph" || node.type.name === "heading") {
      const current = String(
        (node.attrs as { heading?: string; style?: string })?.heading ??
          (node.attrs as { heading?: string; style?: string })?.style ??
          "",
      );
      const match = /^Heading([1-9])$/.exec(current);
      if (match) {
        const lvl = parseInt(match[1]!, 10);
        const nextHeading = lvl > 1 ? `Heading${lvl - 1}` : null;
        const nextStyle = lvl > 1 ? `Heading${lvl - 1}` : "Normal";
        const pos = $from.before(d);
        tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          heading: nextHeading,
          style: nextStyle,
        });
        if (!customTr) editor.view.dispatch(tr);
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
export function demoteHeadingAtCaret(editor: Editor, customTr?: Transaction): boolean {
  const tr = customTr ?? editor.state.tr;
  const { $from } = tr.selection;
  for (let d = $from.depth; d > 0; d--) {
    const node = $from.node(d);
    if (node.type.name === "paragraph" || node.type.name === "heading") {
      const current = String(
        (node.attrs as { heading?: string; style?: string })?.heading ??
          (node.attrs as { heading?: string; style?: string })?.style ??
          "",
      );
      const match = /^Heading([1-9])$/.exec(current);
      const lvl = match ? parseInt(match[1]!, 10) : 0;
      if (lvl < 9) {
        const nextHeading = `Heading${lvl + 1}`;
        const pos = $from.before(d);
        tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          heading: nextHeading,
          style: nextHeading,
        });
        if (!customTr) editor.view.dispatch(tr);
        return true;
      }
    }
  }
  return false;
}

/**
 * Promote heading either at blockIndex or under caret.
 */
export function promoteHeading(
  editor: Editor,
  blockIndex?: number,
  customTr?: Transaction,
): boolean {
  if (blockIndex == null) return promoteHeadingAtCaret(editor, customTr);
  const tr = customTr ?? editor.state.tr;
  const doc = tr.doc;
  if (blockIndex < 0 || blockIndex >= doc.childCount) return false;
  const node = doc.child(blockIndex);
  const styles = (doc.attrs as { styles?: StylesOptions }).styles;
  const lvl = detectNodeHeadingLevel(node, styles);
  if (lvl == null) return false;
  let pos = 0;
  for (let i = 0; i < blockIndex; i++) pos += doc.child(i).nodeSize;
  const nextHeading = lvl > 1 ? `Heading${lvl - 1}` : null;
  const nextStyle = lvl > 1 ? `Heading${lvl - 1}` : "Normal";
  tr.setNodeMarkup(pos, undefined, {
    ...node.attrs,
    heading: nextHeading,
    style: nextStyle,
  });
  if (!customTr) editor.view.dispatch(tr);
  return true;
}

/**
 * Demote heading either at blockIndex or under caret.
 */
export function demoteHeading(
  editor: Editor,
  blockIndex?: number,
  customTr?: Transaction,
): boolean {
  if (blockIndex == null) return demoteHeadingAtCaret(editor, customTr);
  const tr = customTr ?? editor.state.tr;
  const doc = tr.doc;
  if (blockIndex < 0 || blockIndex >= doc.childCount) return false;
  const node = doc.child(blockIndex);
  const styles = (doc.attrs as { styles?: StylesOptions }).styles;
  const lvl = detectNodeHeadingLevel(node, styles) ?? 0;
  if (lvl >= 9) return false;
  let pos = 0;
  for (let i = 0; i < blockIndex; i++) pos += doc.child(i).nodeSize;
  const nextHeading = `Heading${lvl + 1}`;
  tr.setNodeMarkup(pos, undefined, {
    ...node.attrs,
    heading: nextHeading,
    style: nextHeading,
  });
  if (!customTr) editor.view.dispatch(tr);
  return true;
}

/**
 * Set heading level (1-9, or 0 / null for Body Text / Normal).
 */
export function setHeadingLevel(
  editor: Editor,
  level: number | null,
  blockIndex?: number,
  customTr?: Transaction,
): boolean {
  const tr = customTr ?? editor.state.tr;
  const doc = tr.doc;
  const idx = blockIndex ?? tr.selection.$from.index(0);
  if (idx < 0 || idx >= doc.childCount) return false;
  const node = doc.child(idx);
  let pos = 0;
  for (let i = 0; i < idx; i++) pos += doc.child(i).nodeSize;
  const headingAttr = level && level >= 1 && level <= 9 ? `Heading${level}` : null;
  const styleAttr = headingAttr ?? "Normal";
  tr.setNodeMarkup(pos, undefined, {
    ...node.attrs,
    heading: headingAttr,
    style: styleAttr,
  });
  if (!customTr) editor.view.dispatch(tr);
  return true;
}

/**
 * Move the current block or heading section up (Word's Alt+Shift+Up).
 */
export function moveBlockUp(editor: Editor, customTr?: Transaction): boolean {
  const tr = customTr ?? editor.state.tr;
  const { $from } = tr.selection;
  const index = $from.index(0);
  if (index <= 0) return false;

  const doc = tr.doc;
  const styles = (doc.attrs as { styles?: StylesOptions }).styles;
  const currentNode = doc.child(index);
  const isHeading = detectNodeHeadingLevel(currentNode, styles) != null;

  if (isHeading) {
    const srcRange = findHeadingSectionRange(doc, index, styles);
    if (srcRange.startIndex <= 0) return false;
    const prevHeadingIndex = srcRange.startIndex - 1;
    return moveHeadingSection(editor, srcRange.startIndex, prevHeadingIndex, "before", customTr);
  }

  const prevNode = doc.child(index - 1);
  let prevPos = 0;
  for (let i = 0; i < index - 1; i++) {
    prevPos += doc.child(i).nodeSize;
  }
  const curPos = prevPos + prevNode.nodeSize;

  tr.delete(curPos, curPos + currentNode.nodeSize);
  tr.insert(prevPos, currentNode);
  tr.setSelection(TextSelection.near(tr.doc.resolve(prevPos + 1)));
  if (!customTr) editor.view.dispatch(tr);
  return true;
}

/**
 * Move the current block or heading section down (Word's Alt+Shift+Down).
 */
export function moveBlockDown(editor: Editor, customTr?: Transaction): boolean {
  const tr = customTr ?? editor.state.tr;
  const { $from } = tr.selection;
  const index = $from.index(0);
  const doc = tr.doc;
  if (index >= doc.childCount - 1) return false;

  const styles = (doc.attrs as { styles?: StylesOptions }).styles;
  const currentNode = doc.child(index);
  const isHeading = detectNodeHeadingLevel(currentNode, styles) != null;

  if (isHeading) {
    const srcRange = findHeadingSectionRange(doc, index, styles);
    if (srcRange.endIndex >= doc.childCount) return false;
    const nextHeadingIndex = srcRange.endIndex;
    return moveHeadingSection(editor, srcRange.startIndex, nextHeadingIndex, "after", customTr);
  }

  const nextNode = doc.child(index + 1);
  let curPos = 0;
  for (let i = 0; i < index; i++) {
    curPos += doc.child(i).nodeSize;
  }
  const nextPos = curPos + currentNode.nodeSize;

  tr.delete(nextPos, nextPos + nextNode.nodeSize);
  tr.insert(curPos, nextNode);
  tr.setSelection(TextSelection.near(tr.doc.resolve(curPos + nextNode.nodeSize + 1)));
  if (!customTr) editor.view.dispatch(tr);
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
    const sel = TextSelection.create(doc, firstMatch, lastMatch);
    editor.view.dispatch(editor.state.tr.setSelection(sel));
    return true;
  }
  return false;
}
