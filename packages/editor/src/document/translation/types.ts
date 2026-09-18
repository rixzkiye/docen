/**
 * Translation types: language packs, dictionary entries, phrase pairs,
 * grammar transformation rules, and translation options.
 */

export interface LanguageInfo {
  code: string;
  name: string;
  nativeName: string;
}

export interface GrammarRule {
  id: string;
  name: string;
  direction: "en-id" | "id-en" | "any";
  apply(tokens: string[]): string[];
}

export interface TranslationLanguagePack {
  id: string; // e.g. "en-id", "id-en", or target language tag "fr", "de", "es", "zh"
  name: string;
  from: string;
  to: string;
  dictionary: Record<string, string>;
  phrases?: Array<[string, string]>; // [source, target]
  grammarRules?: GrammarRule[];
}

export interface TranslationPackInfo {
  id: string;
  name: string;
  installed: boolean;
  size: string;
  from: string;
  to: string;
}

export interface TranslationOptions {
  from?: string; // "auto" or language code (e.g. "en", "id")
  to?: string; // language code (e.g. "id", "en")
  preserveCase?: boolean;
}

export interface TranslationResult {
  sourceText: string;
  translatedText: string;
  detectedSourceLang: string;
  targetLang: string;
}
