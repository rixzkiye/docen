// The spell checker: word tokenization over the PM doc, dictionary lookup
// against built-in and installed language packs (extendable at runtime),
// custom dictionary persistence, and edit-distance suggestions.

import type { Node as PMNode } from "@tiptap/pm/model";

import { customDictionary } from "./proofing/custom-dictionary";
import { packManager } from "./proofing/pack-manager";
import type { ProofingLanguagePack, ProofingOptions, SpellingIssue } from "./proofing/types";
import { englishWords } from "./spelling-dictionary";

export type {
  GrammarIssue,
  GrammarRule,
  ProofingIssue,
  ProofingLanguagePack,
  ProofingOptions,
  ProofingPackInfo,
  SpellingIssue,
  ThesaurusEntry,
  ThesaurusMeaning,
} from "./proofing/types";
export { customDictionary } from "./proofing/custom-dictionary";
export { checkGrammar } from "./proofing/grammar";
export { packManager } from "./proofing/pack-manager";
export { getSynonyms, lookupThesaurus } from "./proofing/thesaurus";

/** A western word — letters with optional inner apostrophes/hyphens. */
const WORD_RE = /^[A-Za-z\u00C0-\u024F\u1E00-\u1EFF][A-Za-z\u00C0-\u024F\u1E00-\u1EFF'’-]*$/;

/** Session-global extra words (Add to Dictionary) and skipped words (Ignore
 *  All) — shared by every check run so edits don't lose the user's calls. */
const addedWords = new Set<string>();
const ignoredWords = new Set<string>();
/** Ignore Once: exact positions, so only the flagged occurrence is exempt
 *  while the rest of the word keeps its squiggles (Word's per-hit ignore). */
const ignoredOnce = new Set<string>();

/** A word is correct when it hits the session's ignored set, added words,
 *  the persistent custom dictionary, or the language pack's lexicon. */
function known(word: string, pack?: ProofingLanguagePack | null): boolean {
  const lower = word.toLowerCase();
  if (ignoredWords.has(lower) || addedWords.has(lower) || customDictionary.has(lower)) {
    return true;
  }
  const dict = pack?.spellWords ?? englishWords;
  return dict.has(lower) || dict.has(word);
}

/** Extract the language tag on a PM node from its textStyle mark if set. */
function extractLang(node: PMNode): string | undefined {
  const style = node.marks.find((m) => m.type.name === "textStyle");
  if (!style?.attrs.language) return undefined;
  const langAttr = style.attrs.language as { value?: string; val?: string } | string;
  if (typeof langAttr === "string") return langAttr;
  return langAttr.value || langAttr.val;
}

/** Spell check one document: every text node is segmented into words and
 *  looked up against its run language pack. Numbers, URLs, CJK runs and
 *  anything non-word-like are not candidates. Runs marked "do not check
 *  spelling" (w:noProof via the language dialog) are skipped. */
export function checkSpelling(doc: PMNode, options?: Partial<ProofingOptions>): SpellingIssue[] {
  if (options?.checkSpellingAsYouType === false || options?.hideSpellingErrors === true) {
    return [];
  }

  const issues: SpellingIssue[] = [];
  const segmenters = new Map<string, Intl.Segmenter>();

  const getSegmenter = (lang: string): Intl.Segmenter => {
    let seg = segmenters.get(lang);
    if (!seg) {
      try {
        seg = new Intl.Segmenter(lang, { granularity: "word" });
      } catch {
        seg = new Intl.Segmenter("en", { granularity: "word" });
      }
      segmenters.set(lang, seg);
    }
    return seg;
  };

  doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return;
    const style = node.marks.find((m) => m.type.name === "textStyle");
    if (style?.attrs.noProof === true) return;

    const langTag = extractLang(node) || "en";
    const pack = packManager.getPack(langTag);
    const segmenter = getSegmenter(langTag);

    for (const { segment, index, isWordLike } of segmenter.segment(node.text)) {
      if (!isWordLike || !WORD_RE.test(segment) || known(segment, pack)) continue;
      const from = pos + index;
      if (ignoredOnce.has(`${from}:${segment.toLowerCase()}`)) continue;
      issues.push({
        from,
        to: from + segment.length,
        word: segment,
        ...(langTag && langTag !== "en" ? { lang: langTag } : {}),
      });
    }
  });

  return issues;
}

/** Add a word to the session and custom persistent dictionary. */
export function addSpellWord(word: string): void {
  const w = word.toLowerCase();
  addedWords.add(w);
  customDictionary.add(w);
}

/** Skip every occurrence of a word for this session. */
export function ignoreSpellWord(word: string): void {
  ignoredWords.add(word.toLowerCase());
}

/** Skip one occurrence (Ignore Once) — keyed on its position, so other hits
 *  of the same word stay flagged. The exemption lapses when the text moves
 *  (an edit reflows the positions), matching how transient Word's is. */
export function ignoreSpellOnce(issue: { from: number; word: string }): void {
  ignoredOnce.add(`${issue.from}:${issue.word.toLowerCase()}`);
}

/** Damerau-Levenshtein distance (with transpositions) — small words get a
 *  tight radius, longer ones two edits, matching Word's suggestion feel. */
function editDistance(a: string, b: string): number {
  const m = a.length;
  const n = b.length;
  let prev2: number[] = [];
  let prev1: number[] = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) {
    const cur = [i, ...Array<number>(n).fill(0)];
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      cur[j] = Math.min(prev1[j] + 1, cur[j - 1] + 1, prev1[j - 1] + cost);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        cur[j] = Math.min(cur[j], prev2[j - 2] + 1);
      }
    }
    prev2 = prev1;
    prev1 = cur;
  }
  return prev1[n];
}

/** Replacement candidates for a misspelling: dictionary words within edit
 *  range, closest first (ties keep the list's frequency order). */
export function spellSuggestions(word: string, limit = 5, lang?: string): string[] {
  const w = word.toLowerCase();
  const radius = w.length <= 4 ? 1 : 2;
  const scored: Array<[string, number]> = [];
  const seen = new Set<string>();

  const checkCand = (cand: string, bias = 0) => {
    if (seen.has(cand)) return;
    seen.add(cand);
    if (Math.abs(cand.length - w.length) > radius) return;
    const d = editDistance(w, cand);
    if (d <= radius) scored.push([cand, d + bias]);
  };

  for (const added of addedWords) {
    checkCand(added, -0.5); // user words win ties
  }
  for (const custom of customDictionary.all()) {
    checkCand(custom, -0.5);
  }

  const pack = packManager.getPack(lang);
  const words = pack?.spellWords ?? englishWords;
  for (const cand of words) {
    checkCand(cand, 0);
  }

  scored.sort((a, b) => a[1] - b[1]);
  return scored.slice(0, limit).map(([s]) => s);
}
