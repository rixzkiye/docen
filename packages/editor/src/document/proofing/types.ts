/**
 * Proofing types: spelling, grammar, thesaurus, language packs, and options.
 */

export interface ProofingOptions {
  checkSpellingAsYouType: boolean;
  hideSpellingErrors: boolean;
  checkGrammar: boolean;
  hideGrammarErrors: boolean;
}

export const DEFAULT_PROOFING_OPTIONS: ProofingOptions = {
  checkSpellingAsYouType: true,
  hideSpellingErrors: false,
  checkGrammar: true,
  hideGrammarErrors: false,
};

export interface SpellingIssue {
  from: number;
  to: number;
  word: string;
  lang?: string;
}

export interface GrammarIssue {
  from: number;
  to: number;
  message: string;
  ruleId: string;
  category: "grammar" | "style";
  replacements: string[];
}

export type ProofingIssue =
  | ({ kind: "spelling" } & SpellingIssue)
  | ({ kind: "grammar" } & GrammarIssue);

export interface ThesaurusMeaning {
  partOfSpeech: "noun" | "verb" | "adj" | "adv" | "other";
  definition?: string;
  synonyms: string[];
  antonyms?: string[];
}

export interface ThesaurusEntry {
  word: string;
  meanings: ThesaurusMeaning[];
}

export interface GrammarRule {
  id: string;
  name: string;
  category: "grammar" | "style";
  check(text: string, posOffset: number): GrammarIssue[];
}

export interface ProofingLanguagePack {
  id: string; // "en", "id", "zh", etc.
  name: string;
  spellWords: ReadonlySet<string>;
  thesaurus?: Map<string, ThesaurusEntry>;
  grammarRules?: GrammarRule[];
}

export interface ProofingPackInfo {
  id: string;
  name: string;
  installed: boolean;
  size?: string;
}
