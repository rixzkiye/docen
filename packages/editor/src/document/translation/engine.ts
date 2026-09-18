/**
 * Offline Translation Engine:
 * Rule-based, dictionary-driven, phrase-aware offline translation engine.
 * Preserves sentence structure, casing, formatting, and punctuation.
 * 100% offline, zero network requests.
 */

import {
  EN_ADJECTIVES,
  EN_NOUNS,
  EN_TO_ID_PHRASES,
  EN_TO_ID_WORDS,
  ID_ADJECTIVES,
  ID_NOUNS,
  ID_TO_EN_PHRASES,
  ID_TO_EN_WORDS,
} from "./id-en-dictionary";
import { TranslationPackManager, translationPackManager } from "./pack-manager";
import type { TranslationOptions, TranslationResult } from "./types";

/** Helper to escape regex special characters */
function escapeRegExp(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Transfers the casing of a source string onto the translated target string.
 * Supports ALL UPPERCASE, Title / Capitalized, and lowercase.
 */
export function preserveCasing(source: string, target: string): string {
  if (!source || !target) return target;

  const isAllUpper = source === source.toUpperCase() && source !== source.toLowerCase();
  if (isAllUpper) {
    return target.toUpperCase();
  }

  const isCapitalized =
    source.charAt(0) === source.charAt(0).toUpperCase() &&
    source.charAt(0) !== source.charAt(0).toLowerCase();
  if (isCapitalized) {
    return target.charAt(0).toUpperCase() + target.slice(1);
  }

  const isAllLower = source === source.toLowerCase();
  if (isAllLower) {
    return target.toLowerCase();
  }

  return target;
}

/**
 * Automatically detects whether text is English or Indonesian based on
 * token frequency hits against known bilingual vocabularies.
 */
export function detectLanguage(
  text: string,
  _manager: TranslationPackManager = translationPackManager,
): string {
  if (!text || !text.trim()) return "en";

  const words = text.toLowerCase().match(/[\p{L}\p{N}]+(?:[-'][\p{L}\p{N}]+)*/gu) ?? [];
  if (words.length === 0) return "en";

  let enScore = 0;
  let idScore = 0;

  for (const word of words) {
    if (EN_TO_ID_WORDS[word]) enScore++;
    if (ID_TO_EN_WORDS[word]) idScore++;
  }

  if (idScore > enScore) return "id";
  return "en";
}

export class TranslationEngine {
  constructor(private readonly manager: TranslationPackManager = translationPackManager) {}

  /**
   * Main translation method. Translates source text from sourceLang to targetLang.
   */
  translate(text: string, options: TranslationOptions = {}): TranslationResult {
    const rawSource = text ?? "";
    if (!rawSource.trim()) {
      return {
        sourceText: rawSource,
        translatedText: rawSource,
        detectedSourceLang: options.from && options.from !== "auto" ? options.from : "en",
        targetLang: options.to ?? "id",
      };
    }

    let sourceLang = options.from ?? "auto";
    let targetLang = options.to ?? "id";

    // Auto-detect source language if requested
    if (sourceLang === "auto") {
      const detected = detectLanguage(rawSource, this.manager);
      sourceLang = detected;
      // If detected matches targetLang, automatically choose the sensible counter-language
      if (sourceLang === targetLang) {
        targetLang = sourceLang === "en" ? "id" : "en";
      }
    }

    // If source and target languages are identical, return text as-is
    if (sourceLang === targetLang) {
      return {
        sourceText: rawSource,
        translatedText: rawSource,
        detectedSourceLang: sourceLang,
        targetLang,
      };
    }

    const translated = this.#translateInternal(rawSource, sourceLang, targetLang);

    return {
      sourceText: rawSource,
      translatedText: translated,
      detectedSourceLang: sourceLang,
      targetLang,
    };
  }

  #translateInternal(text: string, from: string, to: string): string {
    const fromNorm = this.manager.normalizeLang(from);
    const toNorm = this.manager.normalizeLang(to);

    // Retrieve pack dictionary and phrases
    const pack = this.manager.getPack(fromNorm, toNorm);
    const dict =
      pack?.dictionary ??
      (fromNorm === "en" && toNorm === "id"
        ? EN_TO_ID_WORDS
        : fromNorm === "id" && toNorm === "en"
          ? ID_TO_EN_WORDS
          : {});

    const phrases =
      pack?.phrases ??
      (fromNorm === "en" && toNorm === "id"
        ? EN_TO_ID_PHRASES
        : fromNorm === "id" && toNorm === "en"
          ? ID_TO_EN_PHRASES
          : []);

    let workingText = text;
    const placeholders: Map<string, string> = new Map();
    let placeholderCounter = 0;

    // Step 1: Multi-word phrase matching (longest phrase first)
    for (const [srcPhrase, tgtPhrase] of phrases) {
      const regex = new RegExp(`\\b${escapeRegExp(srcPhrase)}\\b`, "gi");
      workingText = workingText.replace(regex, (match) => {
        const ph = `__DOCEN_TR_PH_${placeholderCounter++}__`;
        const translatedPhrase = preserveCasing(match, tgtPhrase);
        placeholders.set(ph, translatedPhrase);
        return ph;
      });
    }

    // Step 2: Grammar rules (Adjective + Noun swap between EN and ID)
    if (fromNorm === "en" && toNorm === "id") {
      // English: [ADJECTIVE] [NOUN] -> Indonesian: [NOUN] [ADJECTIVE]
      // E.g. "blue car" -> "mobil biru", "important document" -> "dokumen penting"
      const adjNounRegex = /\b([a-zA-Z]+)\s+([a-zA-Z]+)\b/g;
      workingText = workingText.replace(adjNounRegex, (fullMatch, w1, w2) => {
        const w1Lower = w1.toLowerCase();
        const w2Lower = w2.toLowerCase();
        if (EN_ADJECTIVES.has(w1Lower) && EN_NOUNS.has(w2Lower)) {
          const transNoun = dict[w2Lower] ?? w2;
          const transAdj = dict[w1Lower] ?? w1;
          const translatedPair = `${transNoun} ${transAdj}`;
          const ph = `__DOCEN_TR_PH_${placeholderCounter++}__`;
          placeholders.set(ph, preserveCasing(fullMatch, translatedPair));
          return ph;
        }
        return fullMatch;
      });
    } else if (fromNorm === "id" && toNorm === "en") {
      // Indonesian: [NOUN] [ADJECTIVE] -> English: [ADJECTIVE] [NOUN]
      // E.g. "mobil biru" -> "blue car", "dokumen penting" -> "important document"
      const nounAdjRegex = /\b([a-zA-Z]+)\s+([a-zA-Z]+)\b/g;
      workingText = workingText.replace(nounAdjRegex, (fullMatch, w1, w2) => {
        const w1Lower = w1.toLowerCase();
        const w2Lower = w2.toLowerCase();
        if (ID_NOUNS.has(w1Lower) && ID_ADJECTIVES.has(w2Lower)) {
          const transAdj = dict[w2Lower] ?? w2;
          const transNoun = dict[w1Lower] ?? w1;
          const translatedPair = `${transAdj} ${transNoun}`;
          const ph = `__DOCEN_TR_PH_${placeholderCounter++}__`;
          placeholders.set(ph, preserveCasing(fullMatch, translatedPair));
          return ph;
        }
        return fullMatch;
      });

      // Reduplication plurals in Indonesian e.g. "buku-buku" -> "books", "dokumen-dokumen" -> "documents"
      const redupRegex = /\b([a-zA-Z]+)-\1\b/gi;
      workingText = workingText.replace(redupRegex, (fullMatch, baseWord) => {
        const baseLower = baseWord.toLowerCase();
        if (dict[baseLower]) {
          const trans = dict[baseLower];
          const pluralTrans =
            trans.endsWith("s") || trans.endsWith("sh") || trans.endsWith("ch")
              ? `${trans}es`
              : `${trans}s`;
          const ph = `__DOCEN_TR_PH_${placeholderCounter++}__`;
          placeholders.set(ph, preserveCasing(fullMatch, pluralTrans));
          return ph;
        }
        return fullMatch;
      });
    }

    // Step 3: Token-level single word translation with punctuation & whitespace preservation
    // Matches either words or non-words
    const tokenRegex = /([\p{L}\p{N}]+(?:[-'][\p{L}\p{N}]+)*)|([^\p{L}\p{N}]+)/gu;
    const translatedTokens: string[] = [];

    let match: RegExpExecArray | null;
    while ((match = tokenRegex.exec(workingText)) !== null) {
      const word = match[1];
      const nonWord = match[2];

      if (nonWord !== undefined) {
        translatedTokens.push(nonWord);
      } else if (word !== undefined) {
        // If word is a placeholder or part of placeholder, keep as is
        if (word.startsWith("__DOCEN_TR_PH_")) {
          translatedTokens.push(word);
          continue;
        }

        const lower = word.toLowerCase();
        if (dict[lower]) {
          const translatedWord = dict[lower];
          translatedTokens.push(preserveCasing(word, translatedWord));
        } else {
          // Unknown word: retain original word preserving casing
          translatedTokens.push(word);
        }
      }
    }

    let result = translatedTokens.join("");

    // Step 4: Restore placeholders
    for (const [ph, translatedReplacement] of placeholders.entries()) {
      result = result.split(ph).join(translatedReplacement);
    }

    return result;
  }
}

export const translationEngine = new TranslationEngine();

/** Convenience helper function for one-off translations */
export function translateText(
  text: string,
  from = "auto",
  to = "id",
  engine: TranslationEngine = translationEngine,
): string {
  return engine.translate(text, { from, to }).translatedText;
}
