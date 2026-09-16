import {
  formatMarkNames,
  parseFormatRecords,
  runPropsFromMarks,
  runPropsToMarks,
  SECTION_ATTR_KEYS,
  type RunFormatRecord,
  type RunPropMark,
} from "@docen/docx";
import { Extension } from "@docen/docx/core";
import type { MarkType, Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, TextSelection } from "@tiptap/pm/state";
import type { EditorState, Transaction } from "@tiptap/pm/state";
import { AddMarkStep, RemoveMarkStep, ReplaceAroundStep, ReplaceStep } from "@tiptap/pm/transform";

import { getSettings, type IdentitySettings } from "../settings";

/**
 * Track Changes — Word revision tracking for the viewless canvas route.
 *
 * The `insertion`/`deletion` marks (engine extensions) already round-trip
 * w:ins/w:del and render struck/underlined; this extension adds the missing
 * workflow: a tracking toggle, marking of live edits, and accept/reject +
 * navigation commands. All names are the ribbon Review tab's event attributes
 * (`editor.commands.<event>`), wired in WIRED_DISPATCH.
 *
 * Text-edit marking scope is TEXT EDITS INSIDE ONE PARAGRAPH — the same inline
 * scope the OOXML round-trip covers (office-open parses inline w:ins/w:del
 * only). Structural edits (Enter, paste with block content, node deletion)
 * apply untracked, and undo/redo replays as-is — tracking an undo would
 * re-mark the reverted edit and corrupt the revision list.
 *
 * Deletions keep the text (Word: struck, removed on accept): an appendTransaction
 * plugin re-inserts the removed runs from the pre-edit doc — preserving their
 * own rPr marks — under a `deletion` mark placed AFTER the inserted text
 * (LibreOffice's replacement order). Deleting already-struck text is a no-op
 * (Word refuses to delete a deletion); the restore keeps it struck.
 *
 * Format changes (w:rPrChange/w:pPrChange) are recorded by the same plugin:
 * mark add/remove steps on rPr marks and attr-only paragraph replacements
 * (setNodeMarkup) become `formatChange` marks / paragraph `revision` attrs
 * carrying the OLD properties. Accept keeps the new formatting and drops the
 * record; reject restores the old rPr/pPr. The revision author comes from the
 * settings store's identity (name, else initials, else Word's historical
 * default "docen").
 */

/** Fallback revision author when no user identity is set — keeps documents
 *  authored before the settings store (and the specs) on the historical
 *  name. */
export const DEFAULT_REVISION_AUTHOR = "docen";

/** The w:author a new revision gets: the user name, else the initials, else
 *  {@link DEFAULT_REVISION_AUTHOR} (Word writes the full name). */
export function revisionAuthorOf(identity: IdentitySettings): string {
  return identity.name.trim() || identity.initials.trim() || DEFAULT_REVISION_AUTHOR;
}

/** The revision author from the shared settings store — read per edit so a
 *  name change applies to the next revision. */
function revisionAuthor(): string {
  return revisionAuthorOf(getSettings().identity);
}

const revisionDate = (): string =>
  // Word writes second precision ("2026-08-28T09:30:00Z").
  new Date().toISOString().replace(/\.\d+Z$/, "Z");

interface RevisionAttrs {
  id: number;
  author: string;
  date: string;
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const str = (value: unknown): string => (typeof value === "string" ? value : "");

/** A node's w:pPrChange record (paragraph `revision` attr), or undefined. */
function revisionAttrOf(node: PMNode): Record<string, unknown> | undefined {
  const rev = (node.attrs as Record<string, unknown>).revision;
  return isRecord(rev) ? rev : undefined;
}

const isTrackChangeMark = (name: string): boolean => name === "insertion" || name === "deletion";

/** Highest existing revision id + 1 — w:id is a document-unique integer across
 *  text revisions (w:ins/w:del), run format records (w:rPrChange), and
 *  paragraph w:pPrChange records. */
function nextRevisionId(doc: PMNode): number {
  let max = 0;
  doc.descendants((node) => {
    const rev = node.isText ? undefined : revisionAttrOf(node);
    if (rev && typeof rev.id === "number" && rev.id > max) max = rev.id;
    if (!node.isText) return true;
    for (const mark of node.marks) {
      if (!isTrackChangeMark(mark.type.name)) continue;
      const id = (mark.attrs as { id?: unknown }).id;
      if (typeof id === "number" && id > max) max = id;
    }
    for (const record of formatRecordsOf(node)) {
      if (record.id > max) max = record.id;
    }
    return true;
  });
  return max + 1;
}

/** The format-change records on a text node, oldest first (empty when the
 *  mark is absent). */
function formatRecordsOf(node: PMNode | null | undefined): RunFormatRecord[] {
  if (!node?.isText) return [];
  for (const mark of node.marks) {
    if (mark.type.name !== "formatChange") continue;
    return parseFormatRecords((mark.attrs as { records?: unknown }).records);
  }
  return [];
}

/** A text node's rPr marks (name + attrs) — the mark-exact snapshot a format
 *  record's reject restores. */
function rprMarksOf(node: PMNode | null | undefined, names: ReadonlySet<string>): RunPropMark[] {
  if (!node?.isText) return [];
  const marks: RunPropMark[] = [];
  for (const mark of node.marks) {
    if (!names.has(mark.type.name)) continue;
    marks.push({ type: mark.type.name, attrs: { ...(mark.attrs as Record<string, unknown>) } });
  }
  return marks;
}

const sameMarks = (a: readonly RunPropMark[], b: readonly RunPropMark[]): boolean =>
  JSON.stringify(a) === JSON.stringify(b);

/** A tracked edit's metadata. Consecutive same-author edits merge into one
 *  record (Word): text touching a record by this author reuses its attrs
 *  instead of allocating a fresh id — typing a sentence makes ONE w:ins. */
function revisionAttrs(doc: PMNode, from: number, types: MarkType[]): RevisionAttrs {
  const author = revisionAuthor();
  const touch = (node: PMNode | null): RevisionAttrs | null => {
    if (!node?.isText) return null;
    for (const mark of node.marks) {
      if (!types.includes(mark.type)) continue;
      if ((mark.attrs as { author?: string }).author !== author) continue;
      return mark.attrs as unknown as RevisionAttrs;
    }
    return null;
  };
  const $pos = doc.resolve(from);
  return (
    touch($pos.nodeBefore) ??
    touch($pos.nodeAfter) ?? { id: nextRevisionId(doc), author, date: revisionDate() }
  );
}

/** Plugin state = tracking on/off; the toggle command rides a meta so the
 *  flag changes with the transaction that turned it on. */
const trackChangesKey = new PluginKey<boolean>("docenTrackChanges");

/** Review commands (accept/reject) deliberately delete revision text — their
 *  transactions carry this meta so the marking plugin never "restores" what
 *  the review just removed (or records the restore as a fresh format change). */
const skipTrackingKey = new PluginKey<boolean>("docenTrackChangesSkip");

/** The paragraph attrs that are NOT w:pPr — editor-only round-trip markers
 *  plus the revision carrier itself: dropped from a recorded pPrChange and
 *  preserved verbatim by a reject restore. */
const NON_PPR_ATTRS: ReadonlySet<string> = new Set([
  "revision",
  "codeLanguage",
  ...SECTION_ATTR_KEYS,
]);

/** The old pPr a w:pPrChange carries: every non-null persisted prop that is
 *  not editor-only (the office-open ParagraphPropertiesOptions subset). */
function paragraphRevisionProps(attrs: Record<string, unknown>): Record<string, unknown> {
  const props: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (NON_PPR_ATTRS.has(key) || value == null) continue;
    props[key] = value;
  }
  return props;
}

/** True when two paragraphs' persisted pPr differ (editor-only attrs and the
 *  revision carrier itself ignored). */
function paragraphPropsChanged(a: Record<string, unknown>, b: Record<string, unknown>): boolean {
  const before = paragraphRevisionProps(a);
  const after = paragraphRevisionProps(b);
  const keys = new Set([...Object.keys(before), ...Object.keys(after)]);
  for (const key of keys) {
    if (JSON.stringify(before[key]) !== JSON.stringify(after[key])) return true;
  }
  return false;
}

/** One contiguous text span whose rPr is uniform before and after the step and
 *  which carries the same existing format records — the unit one w:rPrChange
 *  covers (Word records format changes per formatting run). */
interface RunFormatSegment {
  from: number;
  to: number;
  before: RunPropMark[];
  after: RunPropMark[];
  records: RunFormatRecord[];
}

/** Split a mark step's range into per-run segments: text nodes sharing the
 *  same before-marks, after-marks and existing records merge, so a mixed run
 *  (bold + plain) yields one record per run instead of one snapshot applied to
 *  the whole selection. */
function runFormatSegments(
  oldState: EditorState,
  tr: Transaction,
  step: AddMarkStep | RemoveMarkStep,
  names: ReadonlySet<string>,
): RunFormatSegment[] {
  const segments: RunFormatSegment[] = [];
  oldState.doc.nodesBetween(step.from, step.to, (node, pos) => {
    if (!node.isText) return true;
    const from = Math.max(pos, step.from);
    const to = Math.min(pos + node.nodeSize, step.to);
    if (from >= to) return true;
    const before = rprMarksOf(node, names);
    // The appended transaction starts from the post-transaction doc, so the
    // step's own mark change is already reflected there.
    const afterNode = tr.doc.nodeAt(from);
    const after = rprMarksOf(afterNode, names);
    const records = formatRecordsOf(afterNode);
    const last = segments[segments.length - 1];
    if (
      last &&
      last.to === from &&
      sameMarks(last.before, before) &&
      sameMarks(last.after, after) &&
      JSON.stringify(last.records) === JSON.stringify(records)
    ) {
      last.to = to;
      return true;
    }
    segments.push({ from, to, before, after, records });
    return true;
  });
  return segments;
}

/** Record AddMark/RemoveMark steps on rPr marks as format revisions
 *  (w:rPrChange) — one record per changed run, per author. Mark steps are
 *  size-neutral, so the pre-transaction offsets stay valid in the appended
 *  transaction's doc. */
function markRunFormatChanges(
  tr: Transaction,
  transactions: readonly Transaction[],
  oldState: EditorState,
  formatType: MarkType,
): boolean {
  const names = new Set(formatMarkNames());
  const author = revisionAuthor();
  let touched = false;
  for (const source of transactions) {
    for (const step of source.steps) {
      if (!(step instanceof AddMarkStep || step instanceof RemoveMarkStep)) continue;
      if (!names.has(step.mark.type.name)) continue;
      for (const segment of runFormatSegments(oldState, tr, step, names)) {
        const records = segment.records.map((record) => ({ ...record }));
        const mine = records.findIndex((record) => record.author === author);
        if (mine >= 0) {
          // Same-author merge: Word folds repeated edits into one rPrChange
          // holding the ORIGINAL before-state; only `after` advances.
          records[mine] = { ...records[mine]!, after: segment.after };
        } else {
          records.push({
            id: nextRevisionId(tr.doc),
            author,
            date: revisionDate(),
            before: segment.before,
            after: segment.after,
            props: runPropsFromMarks(segment.before),
          });
        }
        tr.addMark(
          segment.from,
          segment.to,
          formatType.create({ records: JSON.stringify(records) }),
        );
        touched = true;
      }
    }
  }
  return touched;
}

/** Record attr-only paragraph replacements (setNodeMarkup) as paragraph format
 *  revisions (w:pPrChange): the new attrs stay, the complete old pPr rides the
 *  `revision` attr. */
function markParagraphFormatChanges(
  tr: Transaction,
  transactions: readonly Transaction[],
): boolean {
  const author = revisionAuthor();
  let touched = false;
  for (const source of transactions) {
    for (let i = 0; i < source.steps.length; i++) {
      const step = source.steps[i];
      if (!(step instanceof ReplaceAroundStep)) continue;
      // setNodeMarkup's shape: replace the node, preserve its content (the
      // gap), insert a fresh node whose attrs are the only change. Structural
      // replacements (splits/wraps) carry open depths and fall through.
      if (step.slice.openStart !== 0 || step.slice.openEnd !== 0) continue;
      if (step.insert !== 1 || step.gapFrom !== step.from + 1 || step.gapTo !== step.to - 1) {
        continue;
      }
      const oldNode = source.docs[i].nodeAt(step.from);
      const target = tr.doc.nodeAt(step.from);
      if (!oldNode || !target || !oldNode.isTextblock || target.type !== oldNode.type) continue;
      // Same author keeps the ORIGINAL before-state (folded into one record);
      // a different author records their own change — OOXML carries a single
      // w:pPrChange, so the mirror is replaced and the record must never keep
      // the previous author's attribution.
      const pending = revisionAttrOf(target);
      if (pending && str(pending.author) === author) continue;
      const oldAttrs = oldNode.attrs as Record<string, unknown>;
      if (!paragraphPropsChanged(oldAttrs, target.attrs as Record<string, unknown>)) continue;
      tr.setNodeMarkup(step.from, undefined, {
        ...(target.attrs as Record<string, unknown>),
        revision: {
          id: nextRevisionId(tr.doc),
          author,
          date: revisionDate(),
          ...paragraphRevisionProps(oldAttrs),
        },
      });
      touched = true;
    }
  }
  return touched;
}

/** Mark one transaction round's text ReplaceStep — inserted text under an
 *  insertion mark, removed runs restored struck under a deletion mark. */
function markTextEdits(
  tr: Transaction,
  transactions: readonly Transaction[],
  oldState: EditorState,
  newState: EditorState,
  insertionType: MarkType,
  deletionType: MarkType,
): boolean {
  // One text ReplaceStep per transaction is the tracked shape; anything
  // else (multi-step commands, block edits) applies untracked.
  const steps = transactions.flatMap((source) => source.steps);
  if (steps.length !== 1 || !(steps[0] instanceof ReplaceStep)) return false;
  const step = steps[0];
  const { from, to, slice } = step;

  // The inserted side must be plain text: a single text node, no open
  // depths (structure would need block-level revisions to track).
  const inserted = slice.content;
  let insertedText = "";
  if (slice.openStart !== 0 || slice.openEnd !== 0) return false;
  if (inserted.childCount === 1 && inserted.firstChild!.isText) {
    insertedText = inserted.firstChild!.text ?? "";
  } else if (inserted.childCount !== 0) {
    return false;
  }

  // The removed side must be plain text inside one paragraph of the OLD doc.
  const $from = oldState.doc.resolve(from);
  const $to = oldState.doc.resolve(to);
  if (!$from.sameParent($to) || !$from.parent.isTextblock) return false;
  const removedNodes: PMNode[] = [];
  if (to > from) oldState.doc.slice(from, to).content.forEach((node) => removedNodes.push(node));
  for (const node of removedNodes) {
    if (!node.isText || node.text?.includes("\n")) return false;
  }
  if (!insertedText && removedNodes.length === 0) return false;

  const removedAlreadyStruck = removedNodes.every((node) =>
    node.marks.some((m) => m.type === deletionType),
  );

  // Inserted text → insertion mark, reusing a touching record.
  if (insertedText) {
    tr.addMark(
      from,
      from + insertedText.length,
      insertionType.create(revisionAttrs(newState.doc, from, [insertionType])),
    );
  }
  // Removed text stays: re-insert the pre-edit runs (original rPr marks
  // preserved) struck, after the inserted text. Text that was ALREADY struck
  // keeps its record — the delete is refused, the restore just keeps it.
  if (removedNodes.length > 0) {
    const at = from + insertedText.length;
    const record = revisionAttrs(newState.doc, at, [deletionType]);
    const marked = removedAlreadyStruck
      ? removedNodes
      : removedNodes.map((node) =>
          node.marks.some((m) => m.type === deletionType)
            ? node
            : node.mark([...node.marks, deletionType.create(record)]),
        );
    tr.insert(at, marked);
    // Word parks the caret BEFORE the text it just struck: Backspace keeps
    // marking further left, forward Delete stays put, and a delete that hit
    // already-struck text crosses it instead of stalling on it — without
    // this the caret maps behind the restored runs and every later
    // Backspace retargets the same struck character forever.
    tr.setSelection(TextSelection.create(tr.doc, at));
  }
  return true;
}

const trackChangesPlugin = new Plugin<boolean>({
  key: trackChangesKey,
  state: {
    init: () => false,
    apply: (tr, value) => tr.getMeta(trackChangesKey) ?? value,
  },
  appendTransaction(transactions, oldState, newState) {
    if (!trackChangesKey.getState(newState)) return null;
    // Rounds re-fed by PM (its "appendedTransaction" tag marks our own trs),
    // the toggle itself, review commands, and undo/redo replay untouched.
    if (
      transactions.some(
        (tr) =>
          tr.getMeta("appendedTransaction") !== undefined ||
          tr.getMeta(trackChangesKey) !== undefined ||
          tr.getMeta(skipTrackingKey) !== undefined ||
          tr.getMeta("history$") !== undefined,
      )
    ) {
      return null;
    }

    const tr = newState.tr;
    let changed = false;
    const insertionType = newState.schema.marks.insertion;
    const deletionType = newState.schema.marks.deletion;
    if (insertionType && deletionType) {
      changed = markTextEdits(tr, transactions, oldState, newState, insertionType, deletionType);
    }
    const formatType = newState.schema.marks.formatChange;
    if (formatType) {
      // Format recording is independent of the text shape: a mark toggle is
      // one AddMark/RemoveMark step, a paragraph attr change an attr-only
      // ReplaceAroundStep. Every round must be size-neutral, though — a size
      // change shifts the step offsets those paths read (typing/paste rounds
      // belong to the text path or to the deferred structural tracking).
      const sizeNeutral = transactions.every(
        (source) => source.before.content.size === source.doc.content.size,
      );
      if (sizeNeutral) {
        if (markRunFormatChanges(tr, transactions, oldState, formatType)) changed = true;
        if (markParagraphFormatChanges(tr, transactions)) changed = true;
      }
    }
    return changed ? tr : null;
  },
});

/** One revision as the reviewing pane lists it: the range plus the record's
 *  author/date and the tracked text (the pane shows what was typed or what
 *  Word would remove; a paragraph format change shows its prop diff). */
export type RevisionType = "insertion" | "deletion" | "format";

/** Contiguous revision runs of one record, merged for accept/reject picking. */
interface RevisionRange {
  from: number;
  to: number;
  type: RevisionType;
  id: unknown;
  /** The record's author (w:ins/@w:author) — the display filter's key. */
  author: string;
  date?: string;
  /** Paragraph format change (w:pPrChange): the paragraph node position —
   *  accept/reject rewrites its attrs in place. */
  paraPos?: number;
  /** Run format change (w:rPrChange): the record itself — reject restores this
   *  record's own before-state (mark-exact for editor-originated records). */
  run?: RunFormatRecord;
}

export interface RevisionInfo extends RevisionRange {
  date: string;
  text: string;
}

/** The document's revisions in document order (paragraph-level format changes
 *  first, then their inline runs, then text revisions), with each record's
 *  metadata read off the first node carrying it. */
export function collectRevisions(doc: PMNode): RevisionInfo[] {
  return revisionRanges(doc).map((range) => ({
    ...range,
    date: range.date ?? "",
    text:
      range.paraPos != null
        ? paragraphChangeSummary(doc, range.paraPos)
        : doc.textBetween(range.from, range.to, "\n"),
  }));
}

/** "alignment: center → left, bold: on → off" — a compact picture of what a
 *  paragraph format change did (the pane's card text). */
function paragraphChangeSummary(doc: PMNode, paraPos: number): string {
  const node = doc.nodeAt(paraPos);
  const revision = node ? revisionAttrOf(node) : undefined;
  if (!node || !revision) return "";
  const before = paragraphRevisionProps(revision);
  const after = paragraphRevisionProps(node.attrs as Record<string, unknown>);
  const keys = new Set([...Object.keys(before), ...Object.keys(after)]);
  const parts: string[] = [];
  for (const key of keys) {
    if (JSON.stringify(before[key]) === JSON.stringify(after[key])) continue;
    parts.push(`${key}: ${propLabel(before[key])} → ${propLabel(after[key])}`);
  }
  return parts.join(", ");
}

function propLabel(value: unknown): string {
  if (value == null) return "none";
  if (typeof value === "boolean") return value ? "on" : "off";
  if (typeof value === "string" || typeof value === "number" || typeof value === "bigint") {
    return String(value);
  }
  // Objects (indent/spacing/border bundles) summarize without stringifying.
  return "{…}";
}

function revisionRanges(doc: PMNode): RevisionRange[] {
  const out: RevisionRange[] = [];
  // One record id appears on one contiguous range only; collecting per id
  // keeps document order even when a node carries several records (a node's
  // records would otherwise interleave with the next node's).
  const formatRanges = new Map<number, RevisionRange>();
  doc.descendants((node, pos) => {
    // Paragraph-level w:pPrChange lives on the node, not on a mark.
    const revision = node.isText ? undefined : revisionAttrOf(node);
    if (revision) {
      out.push({
        from: pos + 1,
        to: pos + node.nodeSize - 1,
        type: "format",
        id: revision.id ?? null,
        author: str(revision.author),
        date: str(revision.date),
        paraPos: pos,
      });
    }
    if (!node.isText) return true;
    const from = pos;
    const to = pos + node.nodeSize;
    const track = node.marks.find((mark) => isTrackChangeMark(mark.type.name));
    if (track) {
      const id = (track.attrs as { id?: unknown }).id;
      const author = str((track.attrs as { author?: unknown }).author);
      const last = out[out.length - 1];
      if (last && last.type === track.type.name && last.id === id && last.to === from) {
        last.to = to;
      } else {
        out.push({
          from,
          to,
          type: track.type.name as "insertion" | "deletion",
          id,
          author,
        });
      }
    }
    for (const record of formatRecordsOf(node)) {
      const existing = formatRanges.get(record.id);
      if (existing) {
        if (existing.to === from) existing.to = to;
        continue;
      }
      formatRanges.set(record.id, {
        from,
        to,
        type: "format",
        id: record.id,
        author: record.author,
        date: record.date,
        run: record,
      });
    }
    return true;
  });
  out.push(...formatRanges.values());
  // Stable sort by start — the per-id collection appended run records after
  // the walk, and equal starts keep insertion-before-format (document order).
  return out.sort((a, b) => a.from - b.from);
}

/** Word's "…All Changes Shown" scope: the revisions whose author sits in the
 *  display filter (undefined/empty filter = every revision). */
function shownRanges(
  ranges: RevisionRange[],
  authors: readonly string[] | undefined,
): RevisionRange[] {
  if (!authors || authors.length === 0) return ranges;
  return ranges.filter((r) => r.author !== "" && authors.includes(r.author));
}

/** The revision a command acts on: the one overlapping the selection (an empty
 *  selection inside one counts), else the first after it — Word moves forward
 *  without wrapping. */
function pickRevision(ranges: RevisionRange[], from: number, to: number): RevisionRange | null {
  return (
    ranges.find((r) => r.from <= to && r.to >= from) ?? ranges.find((r) => r.from >= to) ?? null
  );
}

/** Replace a run's whole rPr with this exact mark set — the reject restore:
 *  every rPr mark drops, then the snapshot's marks come back, so a reject
 *  never leaves a reconstructed carrier (e.g. textStyle) behind. */
function setRunMarks(
  tr: Transaction,
  state: EditorState,
  from: number,
  to: number,
  marks: readonly RunPropMark[],
): void {
  for (const name of formatMarkNames()) {
    const markType = state.schema.marks[name];
    if (markType) tr.removeMark(from, to, markType);
  }
  for (const mark of marks) {
    const markType = state.schema.marks[mark.type];
    if (markType) tr.addMark(from, to, markType.create(mark.attrs));
  }
}

/** Restore one run record's before-state on a node segment. The newest record
 *  restores its exact before-mark snapshot; an older record reverts only the
 *  props it changed and that no later record overwrote (a later author's
 *  change is never clobbered); a DOCX-loaded record (no mark snapshot)
 *  reconstructs its before props. */
function restoreRunRecord(
  tr: Transaction,
  state: EditorState,
  from: number,
  to: number,
  record: RunFormatRecord,
  newest: boolean,
  names: ReadonlySet<string>,
): void {
  if (newest && record.before) {
    setRunMarks(tr, state, from, to, record.before);
    return;
  }
  if (!record.after) {
    setRunMarks(tr, state, from, to, runPropsToMarks(record.props));
    return;
  }
  const current = runPropsFromMarks(rprMarksOf(tr.doc.nodeAt(from), names));
  const before = record.before ? runPropsFromMarks(record.before) : record.props;
  const after = runPropsFromMarks(record.after);
  const desired: Record<string, unknown> = { ...current };
  for (const key of new Set([...Object.keys(before), ...Object.keys(after)])) {
    // A later record still owns this prop (its after-value stands) — keep it.
    if (JSON.stringify(current[key]) !== JSON.stringify(after[key])) continue;
    if (before[key] === undefined) delete desired[key];
    else desired[key] = before[key];
  }
  setRunMarks(tr, state, from, to, runPropsToMarks(desired));
}

/** Accept (drop the record, keep the formatting) or reject (restore the
 *  record's before-state) one run format record on its segment(s). */
function applyRunFormatRecord(
  tr: Transaction,
  state: EditorState,
  range: RevisionRange,
  accept: boolean,
): void {
  const record = range.run;
  const formatType = state.schema.marks.formatChange;
  if (!record || !formatType) return;
  const segments: { from: number; to: number }[] = [];
  state.doc.nodesBetween(range.from, range.to, (node, pos) => {
    if (!node.isText) return true;
    const from = Math.max(pos, range.from);
    const to = Math.min(pos + node.nodeSize, range.to);
    if (from < to) segments.push({ from, to });
    return true;
  });
  const names = new Set(formatMarkNames());
  for (const segment of segments) {
    // Read from the live transaction: a sweep may already have removed another
    // record from this node.
    const records = formatRecordsOf(tr.doc.nodeAt(segment.from));
    const index = records.findIndex((r) => r.id === record.id);
    if (index < 0) continue;
    const remaining = records.filter((_, i) => i !== index);
    if (remaining.length > 0) {
      tr.addMark(
        segment.from,
        segment.to,
        formatType.create({ records: JSON.stringify(remaining) }),
      );
    } else {
      tr.removeMark(segment.from, segment.to, formatType);
    }
    if (accept) continue;
    restoreRunRecord(
      tr,
      state,
      segment.from,
      segment.to,
      record,
      index === records.length - 1,
      names,
    );
  }
}

/** Accept (keep the new attrs) or reject (restore the recorded old pPr) one
 *  paragraph format change. */
function applyParagraphFormatChange(
  tr: Transaction,
  state: EditorState,
  range: RevisionRange,
  accept: boolean,
): void {
  const paraPos = range.paraPos;
  if (paraPos == null) return;
  const node = tr.doc.nodeAt(paraPos);
  if (!node) return;
  const attrs = node.attrs as Record<string, unknown>;
  if (accept) {
    tr.setNodeMarkup(paraPos, undefined, { ...attrs, revision: null });
    return;
  }
  // Word: the old pPr wins as a whole — reset to the paragraph's defaults,
  // apply the recorded old props, and keep the editor-only round-trip attrs.
  const next: Record<string, unknown> = {
    ...(node.type.create(null).attrs as Record<string, unknown>),
    ...paragraphRevisionProps(revisionAttrOf(node) ?? {}),
    revision: null,
  };
  for (const key of NON_PPR_ATTRS) {
    if (key === "revision") continue;
    const value = attrs[key];
    if (value !== undefined) next[key] = value;
  }
  tr.setNodeMarkup(paraPos, undefined, next);
}

/** Apply one revision range in either direction — the shared body of the
 *  single-item commands and the four "…All Changes" sweeps. */
function applyRange(
  tr: Transaction,
  state: EditorState,
  range: RevisionRange,
  accept: boolean,
): void {
  if (range.type === "format") {
    if (range.paraPos != null) applyParagraphFormatChange(tr, state, range, accept);
    else applyRunFormatRecord(tr, state, range, accept);
  } else if (range.type === "insertion") {
    if (accept) tr.removeMark(range.from, range.to, state.schema.marks.insertion!);
    else tr.delete(range.from, range.to);
  } else if (accept) {
    tr.delete(range.from, range.to);
  } else {
    tr.removeMark(range.from, range.to, state.schema.marks.deletion!);
  }
}

export const TrackChanges = Extension.create({
  name: "docenTrackChanges",

  addProseMirrorPlugins() {
    return [trackChangesPlugin];
  },

  addCommands() {
    return {
      "track-changes":
        (enabled?: boolean) =>
        ({ state, tr, dispatch }) => {
          const current = trackChangesKey.getState(state) ?? false;
          const next = typeof enabled === "boolean" ? enabled : !current;
          if (next === current) return true;
          if (dispatch) tr.setMeta(trackChangesKey, next);
          return true;
        },
      "accept-change":
        (id?: string) =>
        ({ state, tr, dispatch }) => {
          const ranges = revisionRanges(state.doc);
          // With an id (the reviewing pane's cards) act on that revision;
          // without one, on the revision overlapping the selection.
          const target = id
            ? ranges.find((r) => String(r.id) === id)
            : pickRevision(ranges, state.selection.from, state.selection.to);
          if (!target) return false;
          if (!dispatch) return true;
          applyRange(tr, state, target, true);
          tr.setMeta(skipTrackingKey, true);
          // Word's button is "Accept and Move to Next": after accepting, the
          // caret jumps to the next remaining revision (>= the accepted start
          // — the removal shifted later content back); none left, stay put.
          const next = revisionRanges(tr.doc).find((r) => r.from >= target.from);
          tr.setSelection(TextSelection.near(tr.doc.resolve(next ? next.from : target.from)));
          dispatch(tr.scrollIntoView());
          return true;
        },
      "accept-all-changes":
        () =>
        ({ state, tr, dispatch }) => {
          const ranges = revisionRanges(state.doc);
          if (ranges.length === 0) return false;
          if (!dispatch) return true;
          // Descending order: removing text (an accepted deletion) shifts
          // nothing ahead of it, so the earlier ranges' offsets stay valid.
          // One transaction — one undo step (Word's Accept All).
          for (const r of [...ranges].reverse()) applyRange(tr, state, r, true);
          tr.setMeta(skipTrackingKey, true);
          dispatch(tr);
          return true;
        },
      "reject-change":
        (id?: string) =>
        ({ state, tr, dispatch }) => {
          const ranges = revisionRanges(state.doc);
          const target = id
            ? ranges.find((r) => String(r.id) === id)
            : pickRevision(ranges, state.selection.from, state.selection.to);
          if (!target) return false;
          if (!dispatch) return true;
          applyRange(tr, state, target, false);
          tr.setMeta(skipTrackingKey, true);
          // "Reject and Move to Next" — same walk as accept.
          const next = revisionRanges(tr.doc).find((r) => r.from >= target.from);
          tr.setSelection(TextSelection.near(tr.doc.resolve(next ? next.from : target.from)));
          dispatch(tr.scrollIntoView());
          return true;
        },
      "reject-all-changes":
        () =>
        ({ state, tr, dispatch }) => {
          const ranges = revisionRanges(state.doc);
          if (ranges.length === 0) return false;
          if (!dispatch) return true;
          // Mirror of accept-all: a rejected insertion loses its text, a
          // rejected deletion just loses its mark, and a rejected format
          // change restores its old props. Descending order keeps the earlier
          // offsets valid through the deletions; one transaction, one undo
          // step (Word's Reject All).
          for (const r of [...ranges].reverse()) applyRange(tr, state, r, false);
          tr.setMeta(skipTrackingKey, true);
          dispatch(tr);
          return true;
        },
      // Word's "Accept/Reject All Changes Shown" — the same sweep scoped to
      // the display filter's authors (the host passes the current list).
      "accept-all-changes-shown":
        (authors?: string[]) =>
        ({ state, tr, dispatch }) => {
          const ranges = shownRanges(revisionRanges(state.doc), authors);
          if (ranges.length === 0) return false;
          if (!dispatch) return true;
          for (const r of [...ranges].reverse()) applyRange(tr, state, r, true);
          tr.setMeta(skipTrackingKey, true);
          dispatch(tr);
          return true;
        },
      "reject-all-changes-shown":
        (authors?: string[]) =>
        ({ state, tr, dispatch }) => {
          const ranges = shownRanges(revisionRanges(state.doc), authors);
          if (ranges.length === 0) return false;
          if (!dispatch) return true;
          for (const r of [...ranges].reverse()) applyRange(tr, state, r, false);
          tr.setMeta(skipTrackingKey, true);
          dispatch(tr);
          return true;
        },
      "previous-change":
        () =>
        ({ state, tr, dispatch }) => {
          const target = revisionRanges(state.doc)
            .filter((r) => r.to < state.selection.from)
            .pop();
          if (!target) return false;
          if (!dispatch) return true;
          tr.setSelection(TextSelection.create(tr.doc, target.from));
          dispatch(tr.scrollIntoView());
          return true;
        },
      "next-change":
        () =>
        ({ state, tr, dispatch }) => {
          const target = revisionRanges(state.doc).find((r) => r.from > state.selection.to);
          if (!target) return false;
          if (!dispatch) return true;
          tr.setSelection(TextSelection.create(tr.doc, target.from));
          dispatch(tr.scrollIntoView());
          return true;
        },
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    docenTrackChanges: {
      "track-changes": (enabled?: boolean) => ReturnType;
      "accept-change": (id?: string) => ReturnType;
      "accept-all-changes": () => ReturnType;
      "reject-change": (id?: string) => ReturnType;
      "reject-all-changes": () => ReturnType;
      "accept-all-changes-shown": (authors?: string[]) => ReturnType;
      "reject-all-changes-shown": (authors?: string[]) => ReturnType;
      "previous-change": () => ReturnType;
      "next-change": () => ReturnType;
    };
  }
}
