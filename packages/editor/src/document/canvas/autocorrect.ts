/**
 * As-you-type autocorrect, applied only where typed characters enter the
 * document (the bridge's beforeinput insertText leg). IME commits, spell-check
 * corrections and pastes ride other insert paths and stay untouched — CJK
 * input therefore never sees any of this by construction.
 *
 * The scope is Word's AutoCorrect / AutoFormat As You Type behaviors: the
 * user's replacement table (defaults = Word's common built-ins), smart
 * quotes, the dash forms (-- em dash, - between words en dash), ellipsis,
 * ordinal superscripts, sentence capitalization and URL/email linkification.
 * Every rule is gated by its AutoCorrect setting and every function is pure —
 * the bridge applies the returned fix to one transaction.
 */

import type { AutocorrectReplacement, AutocorrectSettings } from "../settings";

/** The slice of a ProseMirror transaction {@link applyAutocorrect} drives.
 *  Deliberately loose: the workspace resolves two prosemirror-state copies
 *  (the engine's and the editor integration's) and both must satisfy it. */
export interface AutocorrectTransaction {
  /** The current document — its schema is read through {@link AutocorrectSchema}. */
  doc: unknown;
  insertText(text: string, from: number, to: number): AutocorrectTransaction;
  addMark(from: number, to: number, mark: unknown): AutocorrectTransaction;
}

/** The schema slice the fix's marks resolve against. `Schema.marks` is typed
 *  as an OrderedMap but is a plain null-prototype map at runtime
 *  (prosemirror-model's buildMap), so it is addressed by index. */
interface AutocorrectMarks {
  [name: string]: { create(attrs?: unknown): unknown } | undefined;
}

/** Straight quote → [opening, closing] curly pair, keyed by the typed char. */
const SMART_QUOTES: Record<string, [string, string]> = {
  '"': ["“", "”"],
  "'": ["‘", "’"],
};

/** Context that makes a typed quote an opening one: nothing behind it,
 *  whitespace, or an opening bracket/quote/dash. Anything else — a word
 *  character, comma, closing quote — makes it closing (Word: don't → don't
 *  with a closing apostrophe, since a letter precedes). */
const OPENS_QUOTE = /^$|[\s([{<—–“‘'-]$/u;

/** The typed quote's curly replacement, or null for anything else. */
export function smartQuoteOf(typed: string, charBefore: string): string | null {
  const pair = SMART_QUOTES[typed];
  if (!pair) return null;
  return OPENS_QUOTE.test(charBefore) ? pair[0]! : pair[1]!;
}

/** Word's built-in correction table (common subset) — the table a fresh
 *  profile starts with; deleting entries in AutoCorrect Options replaces it. */
export const DEFAULT_AUTOCORRECT_REPLACEMENTS: readonly AutocorrectReplacement[] = [
  { from: "teh", to: "the" },
  { from: "adn", to: "and" },
  { from: "taht", to: "that" },
  { from: "thier", to: "their" },
  { from: "recieve", to: "receive" },
  { from: "seperate", to: "separate" },
  { from: "occured", to: "occurred" },
  { from: "wich", to: "which" },
  { from: "writen", to: "written" },
  { from: "beleive", to: "believe" },
  { from: "freind", to: "friend" },
  { from: "definately", to: "definitely" },
];

/** Word's Math AutoCorrect standard dictionary. */
export const MATH_AUTOCORRECT_MAP: Record<string, string> = {
  "\\alpha": "α",
  "\\beta": "β",
  "\\gamma": "γ",
  "\\delta": "δ",
  "\\epsilon": "ε",
  "\\zeta": "ζ",
  "\\eta": "η",
  "\\theta": "θ",
  "\\iota": "ι",
  "\\kappa": "κ",
  "\\lambda": "λ",
  "\\mu": "μ",
  "\\nu": "ν",
  "\\xi": "ξ",
  "\\pi": "π",
  "\\rho": "ρ",
  "\\sigma": "σ",
  "\\tau": "τ",
  "\\upsilon": "υ",
  "\\phi": "φ",
  "\\chi": "χ",
  "\\psi": "ψ",
  "\\omega": "ω",
  "\\Gamma": "Γ",
  "\\Delta": "Δ",
  "\\Theta": "Θ",
  "\\Lambda": "Λ",
  "\\Xi": "Ξ",
  "\\Pi": "Π",
  "\\Sigma": "Σ",
  "\\Phi": "Φ",
  "\\Psi": "Ψ",
  "\\Omega": "Ω",
  "\\sqrt": "√",
  "\\cbrt": "∛",
  "\\sum": "∑",
  "\\prod": "∏",
  "\\int": "∫",
  "\\oint": "∮",
  "\\infty": "∞",
  "\\pm": "±",
  "\\mp": "∓",
  "\\times": "×",
  "\\div": "÷",
  "\\cdot": "·",
  "\\ast": "∗",
  "\\circ": "∘",
  "\\bullet": "∙",
  "\\le": "≤",
  "\\ge": "≥",
  "\\leq": "≤",
  "\\geq": "≥",
  "\\ne": "≠",
  "\\neq": "≠",
  "\\approx": "≈",
  "\\equiv": "≡",
  "\\sim": "∼",
  "\\cong": "≅",
  "\\propto": "∝",
  "\\partial": "∂",
  "\\nabla": "∇",
  "\\in": "∈",
  "\\notin": "∉",
  "\\subset": "⊂",
  "\\supset": "⊃",
  "\\subseteq": "⊆",
  "\\supseteq": "⊇",
  "\\cap": "∩",
  "\\cup": "∪",
  "\\forall": "∀",
  "\\exists": "∃",
  "\\to": "→",
  "\\rightarrow": "→",
  "\\leftarrow": "←",
  "\\leftrightarrow": "↔",
  "\\Rightarrow": "⇒",
  "\\Leftarrow": "⇐",
  "\\Leftrightarrow": "⇔",
  "\\deg": "°",
  "\\degree": "°",
  "\\angle": "∠",
};

/**
 * The rule configuration the typed leg reads per keystroke. `replacements`
 * is the effective table (the user's when customized, else the defaults) and
 * `exceptions` are words the table must leave alone (case-insensitive).
 */
export interface AutocorrectConfig {
  smartQuotes: boolean;
  emDash: boolean;
  ellipsis: boolean;
  hyperlinkAutoformat: boolean;
  capitalizeFirstLetter: boolean;
  ordinalSuperscript: boolean;
  replacements: readonly AutocorrectReplacement[];
  exceptions: readonly string[];
}

/** Every rule on, built-in table — what a caller gets without settings. */
export const DEFAULT_AUTOCORRECT: AutocorrectConfig = {
  smartQuotes: true,
  emDash: true,
  ellipsis: true,
  hyperlinkAutoformat: true,
  capitalizeFirstLetter: true,
  ordinalSuperscript: false,
  replacements: DEFAULT_AUTOCORRECT_REPLACEMENTS,
  exceptions: [],
};

/** The persisted settings projected onto the rule config (absent user table =
 *  built-in defaults; an explicitly empty table stays empty). */
export function autocorrectConfigOf(settings: AutocorrectSettings): AutocorrectConfig {
  return {
    smartQuotes: settings.smartQuotes,
    emDash: settings.emDash,
    ellipsis: settings.ellipsis,
    hyperlinkAutoformat: settings.hyperlinkAutoformat,
    capitalizeFirstLetter: settings.capitalizeFirstLetter,
    ordinalSuperscript: settings.ordinalSuperscript,
    replacements: settings.table ? settings.table.replacements : DEFAULT_AUTOCORRECT_REPLACEMENTS,
    exceptions: settings.table ? settings.table.exceptions : [],
  };
}

/** Boundary characters that trigger a word correction behind them. */
const WORD_BOUNDARY = /[\s,.!?;:'")\]}]/u;

/** A word may be corrected only at the paragraph start or after whitespace /
 *  opening punctuation — a word glued into a larger token (`http://teh`,
 *  `foo.teh`, `a-teh`) is never rewritten. */
const WORD_PREFIX_OK = /[\s([{<"'“‘—–]$/u;

/** Preserve the corrected word's case shape: lower → lower, ALL CAPS → ALL
 *  CAPS, Capitalized → Capitalized (Word's initial-caps rule). */
function matchCase(replacement: string, source: string): string {
  if (source === source.toUpperCase() && source !== source.toLowerCase()) {
    return replacement.toUpperCase();
  }
  if (/^\p{Lu}/u.test(source)) {
    return replacement.charAt(0).toUpperCase() + replacement.slice(1);
  }
  return replacement;
}

/**
 * The replacement-table correction for `word`, or null. Matching is
 * case-insensitive and the stored replacement follows the typed word's case
 * shape; a word on the exceptions list is left alone.
 */
export function correctionOf(
  word: string,
  config: Pick<AutocorrectConfig, "replacements" | "exceptions"> = DEFAULT_AUTOCORRECT,
): string | null {
  const lower = word.toLowerCase();
  if (config.exceptions.some((entry) => entry.toLowerCase() === lower)) return null;
  for (const entry of config.replacements) {
    if (entry.from.toLowerCase() !== lower) continue;
    return entry.to ? matchCase(entry.to, word) : null;
  }
  return null;
}

/** One mark the fix applies after its text lands. `start` is an offset
 *  relative to the replacement's start and may be negative — linkification
 *  marks the URL already behind the caret. */
export interface AutocorrectMark {
  name: string;
  attrs?: Record<string, unknown>;
  start: number;
  length: number;
}

/** The autocorrect result of one typed character: the replacement text (the
 *  typed char included — corrections ride their boundary character), `back`,
 *  how many characters before the caret it overwrites, and any marks to stamp
 *  over the result. Null when the character needs no rewrite. */
export interface AutocorrectFix {
  text: string;
  back: number;
  marks?: readonly AutocorrectMark[];
}

/** The ordinal suffix for a number: 1st, 2nd, 3rd, 4th, 11th… */
function ordinalSuffix(value: number): "st" | "nd" | "rd" | "th" {
  const mod100 = value % 100;
  if (mod100 >= 11 && mod100 <= 13) return "th";
  switch (value % 10) {
    case 1:
      return "st";
    case 2:
      return "nd";
    case 3:
      return "rd";
    default:
      return "th";
  }
}

/** The typed character completing a digit + st/nd/rd/th ordinal: the already
 *  typed suffix char is replaced with the full suffix, superscript-marked. */
function ordinalFix(typed: string, textBefore: string): AutocorrectFix | null {
  if (!/[a-zA-Z]/.test(typed)) return null;
  const match = /(\d+)([snrt])$/.exec(textBefore);
  if (!match) return null;
  // A digit run glued to letters is not an ordinal (abc1st).
  if (match.index > 0 && /\p{L}$/u.test(textBefore.slice(0, match.index))) return null;
  const suffix = (match[2]! + typed).toLowerCase();
  if (!/^(st|nd|rd|th)$/.test(suffix)) return null;
  if (ordinalSuffix(Number(match[1])) !== suffix) return null;
  const text = match[2]! + typed;
  return { text, back: 1, marks: [{ name: "superscript", start: 0, length: text.length }] };
}

/** Sentence terminators that license the next letter's capitalization. */
const SENTENCE_END = /[.!?…]["'”’)\]]*\s+$/u;

/** Multi-dot abbreviations ("e.g.", "i.e.", "U.S.") never end a sentence. */
const MULTI_DOT_ABBREVIATION = /(?:[A-Za-z]\.){2,}$/;

/** Common abbreviations after which Word does not capitalize (single word
 *  before the period; lowercased). */
const ABBREVIATIONS: ReadonlySet<string> = new Set([
  "approx",
  "cf",
  "co",
  "corp",
  "dept",
  "dr",
  "etc",
  "est",
  "fig",
  "gen",
  "gov",
  "hon",
  "ibid",
  "inc",
  "jr",
  "ltd",
  "mr",
  "mrs",
  "ms",
  "mt",
  "no",
  "nos",
  "pres",
  "prof",
  "rep",
  "rev",
  "sen",
  "sr",
  "st",
  "vs",
  "vol",
]);

/** Capitalize a typed lowercase letter that opens a sentence (or paragraph);
 *  a built-in abbreviation / user exception before the terminator suppresses
 *  it. Only single-character commits are touched (IME strings pass through). */
function capitalizeFix(
  typed: string,
  textBefore: string,
  exceptions: readonly string[],
): AutocorrectFix | null {
  if (!/^\p{Ll}$/u.test(typed)) return null;
  if (!/^\s*$/u.test(textBefore)) {
    if (!SENTENCE_END.test(textBefore)) return null;
    const head = textBefore.replace(/\s+$/u, "");
    if (MULTI_DOT_ABBREVIATION.test(head)) return null;
    const last = /([\p{L}]+)\.$/u.exec(head)?.[1]?.toLowerCase();
    if (last && (ABBREVIATIONS.has(last) || exceptions.some((e) => e.toLowerCase() === last))) {
      return null;
    }
  }
  return { text: typed.toUpperCase(), back: 0 };
}

// ── Hyperlink autoformat ─────────────────────────────────────────────────────

const SCHEME_URL = /^(?:https?|ftp):\/\/[^\s<>"'`]+/i;
const WWW_URL = /^www\.[^\s<>"'`]+/i;
const EMAIL = /^[\w.+-]+@[\w-]+(?:\.[\w-]+)+/;
const UNC_PATH = /^\\\\[^\s<>"'`\\]+\\[^\s<>"'`]*/;

/** Strip sentence punctuation off a matched URL's tail (keeping a closing
 *  paren that balances an opening one inside the URL). */
function trimUrlTail(url: string): string {
  let end = url.length;
  while (end > 0) {
    const char = url[end - 1]!;
    if (/[.,;:!?…'"”’\]}>]/u.test(char)) {
      end--;
      continue;
    }
    if (char === ")") {
      const slice = url.slice(0, end);
      const opens = (slice.match(/\(/g) ?? []).length;
      const closes = (slice.match(/\)/g) ?? []).length;
      if (closes > opens) {
        end--;
        continue;
      }
    }
    break;
  }
  return url.slice(0, end);
}

/**
 * The URL/email ending `textBefore`, or null. Recognizes scheme URLs
 * (http/https/ftp), `www.` hosts, email addresses and UNC paths — the forms
 * Word's "Internet and network paths with hyperlinks" autoformats. The
 * returned `text` is the exact run to mark; `href` is its link target.
 */
export function hyperlinkAtEnd(textBefore: string): { href: string; text: string } | null {
  const token = /[^\s]+$/.exec(textBefore)?.[0];
  if (!token) return null;
  const opened = /^[("'<{[]+/.exec(token)?.[0] ?? "";
  const body = token.slice(opened.length);
  for (const pattern of [SCHEME_URL, WWW_URL, EMAIL, UNC_PATH]) {
    const match = pattern.exec(body);
    if (!match || match.index !== 0) continue;
    const url = trimUrlTail(match[0]);
    if (!url) continue;
    if (pattern === EMAIL) return { href: `mailto:${url}`, text: url };
    if (pattern === WWW_URL) return { href: `https://${url}`, text: url };
    return { href: url, text: url };
  }
  return null;
}

/** The link mark (+ Word's Hyperlink character style) over the URL `length`
 *  characters ending at the insertion point. */
function hyperlinkMarks(length: number, href: string): AutocorrectMark[] {
  return [
    { name: "link", attrs: { href }, start: -length, length },
    { name: "textStyle", attrs: { style: "Hyperlink" }, start: -length, length },
  ];
}

/**
 * The hyperlink-autoformat fix for a paragraph whose text ends in a URL: the
 * caller inserts nothing and only stamps the marks over the URL behind the
 * caret (the bridge's Enter leg runs it just before splitting the paragraph).
 */
export function hyperlinkFix(
  textBefore: string,
  config: AutocorrectConfig = DEFAULT_AUTOCORRECT,
): AutocorrectFix | null {
  if (!config.hyperlinkAutoformat) return null;
  const link = hyperlinkAtEnd(textBefore);
  if (!link) return null;
  return { text: "", back: 0, marks: hyperlinkMarks(link.text.length, link.href) };
}

/** True when the text's trailing token is inside a URL/email — the symbol
 *  rules (dashes, ellipsis) must not rewrite URL characters. */
function inUrlToken(textBefore: string): boolean {
  const token = /[^\s]+$/.exec(textBefore)?.[0];
  if (!token) return false;
  return /^(?:[a-z][a-z0-9+.-]*:\/\/|www\.|\\\\)/i.test(token) || token.includes("@");
}

/**
 * The autocorrect result of one typed character, gated by `config`.
 *
 * `textBefore` is the paragraph text before the caret — quotes read one
 * character back, word corrections read to the last boundary, capitalization
 * reads the sentence terminator, and URL detection reads the trailing token.
 * The returned fix is applied by {@link applyAutocorrect}; null when the
 * character needs no rewrite.
 */
export function autocorrectOf(
  typed: string,
  textBefore: string,
  config: AutocorrectConfig = DEFAULT_AUTOCORRECT,
): AutocorrectFix | null {
  // Multi-character commits (IME strings, autocomplete, paste-style input)
  // never transform: the rules read one typed character's context.
  const single = typed.length === 1 ? typed : "";

  if (config.smartQuotes && single) {
    const quote = smartQuoteOf(single, textBefore.slice(-1));
    if (quote) return { text: quote, back: 0 };
  }

  // Hyperlink autoformat: the typed space completes a URL.
  if (config.hyperlinkAutoformat && typed === " ") {
    const link = hyperlinkAtEnd(textBefore);
    if (link) {
      return { text: " ", back: 0, marks: hyperlinkMarks(link.text.length, link.href) };
    }
  }

  if (!inUrlToken(textBefore)) {
    // The second hyphen of a pair becomes Word's em dash (the first is already
    // behind the caret). Between words, a spaced hyphen + typed space becomes
    // an en dash. The third dot completes an ellipsis.
    if (config.emDash && single === "-" && textBefore.endsWith("-")) {
      return { text: "—", back: 1 };
    }
    if (config.emDash && typed === " " && /\S\s+-$/.test(textBefore)) {
      return { text: "– ", back: 1 };
    }
    if (config.ellipsis && single === "." && textBefore.endsWith("..")) {
      return { text: "…", back: 2 };
    }
  }

  if (config.ordinalSuperscript && single) {
    const ordinal = ordinalFix(single, textBefore);
    if (ordinal) return ordinal;
  }

  if (config.capitalizeFirstLetter && single) {
    const capitalized = capitalizeFix(single, textBefore, config.exceptions);
    if (capitalized) return capitalized;
  }

  // Math AutoCorrect: \cmd followed by boundary, operator, or space
  if (
    WORD_BOUNDARY.test(typed) ||
    /[\s+\-*/^=_()[\]{}<>,.!?;:'"\\|`~]/u.test(typed) ||
    typed === "\n"
  ) {
    const mathMatch = /\\[a-zA-Z]+$/.exec(textBefore);
    if (mathMatch) {
      const sym = MATH_AUTOCORRECT_MAP[mathMatch[0]];
      if (sym) {
        return { text: sym + (typed === " " ? "" : typed), back: mathMatch[0].length };
      }
    }
  }

  // Replacement table: the typed char is a boundary, the word behind it may
  // need fixing. The boundary trails the corrected word — "teh " becomes
  // "the " with the typed space kept as the word separator.
  if (!WORD_BOUNDARY.test(typed) && typed !== "\n") return null;
  const word = /[\p{L}\p{N}]+$/u.exec(textBefore)?.[0];
  if (!word) return null;
  const prefix = textBefore.slice(0, textBefore.length - word.length);
  if (prefix !== "" && !WORD_PREFIX_OK.test(prefix)) return null;
  const fixed = correctionOf(word, config);
  if (!fixed) return null;
  return { text: fixed + typed, back: word.length };
}

/**
 * Apply one fix to a transaction: replace `[start, to]` with the fix's text,
 * then stamp each mark. Shared by the bridge's typed leg (url/ordinal marks)
 * and its Enter leg (linkification with no text change). Returns the caller's
 * transaction type (the chained result is the same class at runtime).
 */
export function applyAutocorrect<T extends AutocorrectTransaction>(
  tr: T,
  start: number,
  to: number,
  fix: AutocorrectFix,
): T {
  let next: AutocorrectTransaction = fix.text ? tr.insertText(fix.text, start, to) : tr;
  if (fix.marks?.length) {
    const marks = (tr.doc as { type: { schema: { marks: AutocorrectMarks } } }).type.schema.marks;
    for (const mark of fix.marks) {
      const type = marks[mark.name];
      if (!type) continue;
      const from = start + mark.start;
      next = next.addMark(from, from + mark.length, type.create(mark.attrs));
    }
  }
  // The structural view exists only for the shared methods; the runtime object
  // is the caller's transaction (or its chained result of the same class).
  return next as T;
}
