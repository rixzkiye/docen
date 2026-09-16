/**
 * Offline rule-based Grammar and Style checker.
 * Detects common syntax errors, capitalization, repeated words, spacing,
 * article usage, and subject-verb agreements.
 */

import type { Node as PMNode } from "@tiptap/pm/model";

import type { GrammarIssue, ProofingOptions } from "./types";

interface RuleMatch {
  from: number;
  to: number;
  message: string;
  ruleId: string;
  category: "grammar" | "style";
  replacements: string[];
}

/** Check text for repeated consecutive words (e.g., "the the", "in in"). */
function checkRepeatedWords(text: string, basePos: number): RuleMatch[] {
  const matches: RuleMatch[] = [];
  const re = /\b([A-Za-z]+)\s+\1\b/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const from = basePos + m.index;
    const word = m[1]!;
    matches.push({
      from,
      to: from + m[0].length,
      message: `Repeated word: "${word}"`,
      ruleId: "repeated-words",
      category: "grammar",
      replacements: [word],
    });
  }
  return matches;
}

/** Check space before punctuation (e.g., "word , " -> "word, "). */
function checkSpaceBeforePunctuation(text: string, basePos: number): RuleMatch[] {
  const matches: RuleMatch[] = [];
  const re = /\s+([,.:;?!])/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const from = basePos + m.index;
    const punct = m[1]!;
    matches.push({
      from,
      to: from + m[0].length,
      message: `Unnecessary space before punctuation: "${punct}"`,
      ruleId: "space-before-punctuation",
      category: "style",
      replacements: [punct],
    });
  }
  return matches;
}

/** Check multiple consecutive spaces inside text. */
function checkMultipleSpaces(text: string, basePos: number): RuleMatch[] {
  const matches: RuleMatch[] = [];
  const re = /[^\s]\s{2,}[^\s]/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const from = basePos + m.index + 1;
    const len = m[0].length - 2;
    matches.push({
      from,
      to: from + len,
      message: "Multiple consecutive spaces",
      ruleId: "multiple-spaces",
      category: "style",
      replacements: [" "],
    });
  }
  return matches;
}

/** Check English sentence capitalization. */
function checkSentenceCapitalization(text: string, basePos: number): RuleMatch[] {
  const matches: RuleMatch[] = [];
  // Match beginning of text if lowercase letter
  const startMatch = /^[ \t]*([a-z])/.exec(text);
  if (startMatch && startMatch[1]) {
    const word = startMatch[1];
    const idx = startMatch.index + startMatch[0].length - word.length;
    matches.push({
      from: basePos + idx,
      to: basePos + idx + word.length,
      message: "Sentence should begin with a capital letter",
      ruleId: "sentence-capitalization",
      category: "grammar",
      replacements: [word.charAt(0).toUpperCase() + word.slice(1)],
    });
  }

  // Match lowercase letter after sentence terminator (. ! ?)
  const re = /([.!?]\s+)([a-z][a-z0-9'-]*)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const prefixLen = m[1]!.length;
    const from = basePos + m.index + prefixLen;
    const word = m[2]!;
    matches.push({
      from,
      to: from + word.length,
      message: "Sentence should begin with a capital letter",
      ruleId: "sentence-capitalization",
      category: "grammar",
      replacements: [word.charAt(0).toUpperCase() + word.slice(1)],
    });
  }
  return matches;
}

/** Check English indefinite article ('a' vs 'an'). */
function checkIndefiniteArticles(text: string, basePos: number): RuleMatch[] {
  const matches: RuleMatch[] = [];
  // 'a' before vowel sound words
  const aVowel = /\b(a)\s+([aeio][a-z]*|un(?![ni])[a-z]*|hour[a-z]*|honor[a-z]*|honest[a-z]*)\b/gi;
  let m: RegExpExecArray | null;
  while ((m = aVowel.exec(text)) !== null) {
    const from = basePos + m.index;
    const to = from + m[1]!.length;
    matches.push({
      from,
      to,
      message: 'Use "an" before a vowel sound',
      ruleId: "indefinite-articles",
      category: "grammar",
      replacements: [m[1] === "A" ? "An" : "an"],
    });
  }

  // 'an' before consonant words
  const anCons =
    /\b(an)\s+([b-df-hj-np-tv-z][a-z]*|uni[a-z]*|use[a-z]*|user[a-z]*|eu[a-z]*|one)\b/gi;
  while ((m = anCons.exec(text)) !== null) {
    const from = basePos + m.index;
    const to = from + m[1]!.length;
    matches.push({
      from,
      to,
      message: 'Use "a" before a consonant sound',
      ruleId: "indefinite-articles",
      category: "grammar",
      replacements: [m[1] === "An" ? "A" : "a"],
    });
  }
  return matches;
}

/** Check common subject-verb agreement issues. */
function checkSubjectVerbAgreement(text: string, basePos: number): RuleMatch[] {
  const matches: RuleMatch[] = [];
  const patterns: Array<{ re: RegExp; repl: (match: string) => string; msg: string }> = [
    {
      re: /\b(they|we)\s+(is)\b/gi,
      repl: (m) => m.replace(/is$/i, (s) => (s === "IS" ? "ARE" : s === "Is" ? "Are" : "are")),
      msg: 'Subject-verb agreement: plural subject requires "are"',
    },
    {
      re: /\b(they|we)\s+(was)\b/gi,
      repl: (m) =>
        m.replace(/was$/i, (s) => (s === "WAS" ? "WERE" : s === "Was" ? "Were" : "were")),
      msg: 'Subject-verb agreement: plural subject requires "were"',
    },
    {
      re: /\b(he|she|it)\s+(are)\b/gi,
      repl: (m) => m.replace(/are$/i, (s) => (s === "ARE" ? "IS" : s === "Are" ? "Is" : "is")),
      msg: 'Subject-verb agreement: singular subject requires "is"',
    },
    {
      re: /\b(he|she|it)\s+(were)\b/gi,
      repl: (m) =>
        m.replace(/were$/i, (s) => (s === "WERE" ? "WAS" : s === "Were" ? "Was" : "was")),
      msg: 'Subject-verb agreement: singular subject requires "was"',
    },
    {
      re: /\b(he|she|it)\s+(have)\b/gi,
      repl: (m) =>
        m.replace(/have$/i, (s) => (s === "HAVE" ? "HAS" : s === "Have" ? "Has" : "has")),
      msg: 'Subject-verb agreement: singular subject requires "has"',
    },
  ];

  for (const { re, repl, msg } of patterns) {
    let m: RegExpExecArray | null;
    while ((m = re.exec(text)) !== null) {
      const from = basePos + m.index;
      const to = from + m[0].length;
      matches.push({
        from,
        to,
        message: msg,
        ruleId: "subject-verb-agreement",
        category: "grammar",
        replacements: [repl(m[0])],
      });
    }
  }

  return matches;
}

/** Check common word confusion patterns (their/there/they're, its/it's). */
function checkCommonConfusion(text: string, basePos: number): RuleMatch[] {
  const matches: RuleMatch[] = [];
  const patterns: Array<{ re: RegExp; repl: string; msg: string }> = [
    {
      re: /\b(their)\s+(are|is|were|was)\b/gi,
      repl: "there",
      msg: 'Did you mean "there"?',
    },
    {
      re: /\b(there)\s+(going|doing|making|coming)\b/gi,
      repl: "they're",
      msg: 'Did you mean "they\'re"?',
    },
    {
      re: /\b(they're)\s+(car|house|room|homework|work|dog|cat|books?|parents?|friends?|names?|lives|way|own|family|money|time)\b/gi,
      repl: "their",
      msg: 'Did you mean "their"?',
    },
    {
      re: /\b(your)\s+(welcome|right|wrong|going)\b/gi,
      repl: "you're",
      msg: 'Did you mean "you\'re"?',
    },
    {
      re: /\b(you're)\s+(car|house|room|homework|work|dog|cat|books?|parents?|friends?|names?)\b/gi,
      repl: "your",
      msg: 'Did you mean "your"?',
    },
    {
      re: /\b(its)\s+(a|an|the)\b/gi,
      repl: "it's",
      msg: 'Did you mean "it\'s"?',
    },
  ];

  for (const { re, repl, msg } of patterns) {
    let m: RegExpExecArray | null;
    while ((m = re.exec(text)) !== null) {
      const targetWord = m[1]!;
      const from = basePos + m.index;
      const to = from + targetWord.length;
      const replacement =
        targetWord[0] === targetWord[0]?.toUpperCase()
          ? repl.charAt(0).toUpperCase() + repl.slice(1)
          : repl;
      matches.push({
        from,
        to,
        message: msg,
        ruleId: "frequently-confused",
        category: "grammar",
        replacements: [replacement],
      });
    }
  }

  return matches;
}

/** Check entire ProseMirror document for grammar and style issues. */
export function checkGrammar(doc: PMNode, options?: Partial<ProofingOptions>): GrammarIssue[] {
  if (options?.checkGrammar === false || options?.hideGrammarErrors === true) {
    return [];
  }

  const issues: GrammarIssue[] = [];

  doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return;
    const style = node.marks.find((m) => m.type.name === "textStyle");
    if (style?.attrs.noProof === true) return;

    const text = node.text;
    const matched = [
      ...checkRepeatedWords(text, pos),
      ...checkSpaceBeforePunctuation(text, pos),
      ...checkMultipleSpaces(text, pos),
      ...checkSentenceCapitalization(text, pos),
      ...checkIndefiniteArticles(text, pos),
      ...checkSubjectVerbAgreement(text, pos),
      ...checkCommonConfusion(text, pos),
    ];

    issues.push(...matched);
  });

  // Sort in document order
  issues.sort((a, b) => a.from - b.from);
  return issues;
}
