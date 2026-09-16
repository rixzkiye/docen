/**
 * docen hyphenation engine.
 *
 * Implements Liang pattern-based hyphenation for English (en) and PUEBI
 * syllabification rules for Indonesian (id) with zero external dependencies.
 * Inserts discretionary soft hyphens (\u00AD) at eligible break positions.
 */

export interface HyphenationOptions {
  autoHyphenation?: boolean;
  doNotHyphenateCaps?: boolean;
  hyphenationZoneTw?: number;
  consecutiveHyphenLimit?: number;
  minWordLength?: number;
  minPrefix?: number;
  minSuffix?: number;
}

const SOFT_HYPHEN = "\u00AD";

// ── Indonesian (id / id-ID) Syllabification (PUEBI) ──

const ID_VOWELS = new Set(["a", "i", "u", "e", "o"]);
const ID_DIPHTHONGS = ["ai", "au", "ei", "oi"];
const ID_DIGRAPHS = ["ng", "ny", "sy", "kh"];

function isIdVowel(ch: string): boolean {
  return ID_VOWELS.has(ch.toLowerCase());
}

/**
 * Break an Indonesian word into syllables according to PUEBI rules.
 */
export function syllabifyIndonesian(word: string): string[] {
  if (word.length < 4) return [word];
  const lower = word.toLowerCase();
  const breaks: number[] = [];
  let i = 0;

  while (i < word.length) {
    // Check for diphthong at i (diphthongs occur at word boundaries e.g. pulau, pantai)
    const two = lower.slice(i, i + 2);
    if (
      ID_DIPHTHONGS.includes(two) &&
      (i + 2 === word.length ||
        (i + 3 < word.length && !isIdVowel(lower[i + 2]!) && isIdVowel(lower[i + 3]!)))
    ) {
      i += 2;
      continue;
    }
    // Check V-V hiatus break (e.g. da-un, ma-in, su-a-ra)
    if (i > 0 && isIdVowel(lower[i - 1]!) && isIdVowel(lower[i]!)) {
      breaks.push(i);
      i++;
      continue;
    }

    // Check V-C... patterns
    if (isIdVowel(lower[i]!)) {
      let cCount = 0;
      let j = i + 1;
      while (j < word.length && !isIdVowel(lower[j]!)) {
        // Treat digraphs as a single consonant
        const d = lower.slice(j, j + 2);
        if (ID_DIGRAPHS.includes(d)) {
          cCount++;
          j += 2;
        } else {
          cCount++;
          j++;
        }
      }
      if (j < word.length) {
        // We found consonants followed by a vowel at j
        if (cCount === 1) {
          // Rule: V-CV -> break before consonant
          breaks.push(i + 1);
        } else if (cCount >= 2) {
          // Rule: VC-CV -> break after first consonant
          const dFirst = ID_DIGRAPHS.includes(lower.slice(i + 1, i + 3));
          const splitOffset = dFirst ? 2 : 1;
          breaks.push(i + 1 + splitOffset);
        }
      }
      i = j;
      continue;
    }
    i++;
  }

  if (breaks.length === 0) return [word];

  const syllables: string[] = [];
  let prev = 0;
  for (const b of breaks) {
    if (b > prev && b < word.length) {
      syllables.push(word.slice(prev, b));
      prev = b;
    }
  }
  if (prev < word.length) {
    syllables.push(word.slice(prev));
  }
  return syllables;
}

// ── English (en) Liang Pattern Engine ──

/**
 * Compact, high-coverage English hyphenation patterns (Liang's algorithm).
 */
const EN_PATTERNS: readonly string[] = [
  "4ab.",
  "4ac.",
  "4ad.",
  "4af.",
  "4ag.",
  "4al.",
  "4am.",
  "4an.",
  "4ap.",
  "4ar.",
  "4as.",
  "4at.",
  ".an4",
  ".anti4",
  ".auto4",
  ".bi3",
  ".co3",
  ".de3",
  ".dis4",
  ".en4",
  ".ex4",
  ".fore4",
  ".hyper4",
  ".in3",
  ".inter4",
  ".intro4",
  ".macro4",
  ".micro4",
  ".mis4",
  ".multi4",
  ".non4",
  ".out4",
  ".over4",
  ".para4",
  ".peri4",
  ".post4",
  ".pre3",
  ".pro3",
  ".pseudo4",
  ".re3",
  ".semi4",
  ".sub4",
  ".super4",
  ".trans4",
  ".un4",
  ".under4",
  "4able.",
  "4ably.",
  "4age.",
  "4al.",
  "4ance.",
  "4ant.",
  "4ar.",
  "4ary.",
  "4ate.",
  "4ation.",
  "4ative.",
  "4dom.",
  "4ed.",
  "4en.",
  "4ence.",
  "4ent.",
  "4er.",
  "4ers.",
  "4ery.",
  "4es.",
  "4est.",
  "4ful.",
  "4fully.",
  "4hood.",
  "4ic.",
  "4ical.",
  "4ice.",
  "4ify.",
  "4ing.",
  "4ings.",
  "4ion.",
  "4ions.",
  "4ish.",
  "4ism.",
  "4ist.",
  "4ite.",
  "4ity.",
  "4ive.",
  "4ize.",
  "4less.",
  "4ly.",
  "4ment.",
  "4ments.",
  "4ness.",
  "4or.",
  "4ous.",
  "4ship.",
  "4sion.",
  "4tion.",
  "4tions.",
  "4ure.",
  "4ward.",
  "4wise.",
  "a1b",
  "a1c",
  "a1d",
  "a1f",
  "a1g",
  "a1l",
  "a1m",
  "a1n",
  "a1p",
  "a1r",
  "a1s",
  "a1t",
  "a1v",
  "e1b",
  "e1c",
  "e1d",
  "e1f",
  "e1g",
  "e1l",
  "e1m",
  "e1n",
  "e1p",
  "e1r",
  "e1s",
  "e1t",
  "e1v",
  "i1b",
  "i1c",
  "i1d",
  "i1f",
  "i1g",
  "i1l",
  "i1m",
  "i1n",
  "i1p",
  "i1r",
  "i1s",
  "i1t",
  "i1v",
  "o1b",
  "o1c",
  "o1d",
  "o1f",
  "o1g",
  "o1l",
  "o1m",
  "o1n",
  "o1p",
  "o1r",
  "o1s",
  "o1t",
  "o1v",
  "u1b",
  "u1c",
  "u1d",
  "u1f",
  "u1g",
  "u1l",
  "u1m",
  "u1n",
  "u1p",
  "u1r",
  "u1s",
  "u1t",
  "u1v",
  "1ba",
  "1be",
  "1bi",
  "1bo",
  "1bu",
  "1by",
  "1ca",
  "1ce",
  "1ci",
  "1co",
  "1cu",
  "1cy",
  "1da",
  "1de",
  "1di",
  "1do",
  "1du",
  "1dy",
  "1fa",
  "1fe",
  "1fi",
  "1fo",
  "1fu",
  "1fy",
  "1ga",
  "1ge",
  "1gi",
  "1go",
  "1gu",
  "1gy",
  "1ha",
  "1he",
  "1hi",
  "1ho",
  "1hu",
  "1hy",
  "1ja",
  "1je",
  "1ji",
  "1jo",
  "1ju",
  "1jy",
  "1ka",
  "1ke",
  "1ki",
  "1ko",
  "1ku",
  "1ky",
  "1la",
  "1le",
  "1li",
  "1lo",
  "1lu",
  "1ly",
  "1ma",
  "1me",
  "1mi",
  "1mo",
  "1mu",
  "1my",
  "1na",
  "1ne",
  "1ni",
  "1no",
  "1nu",
  "1ny",
  "1pa",
  "1pe",
  "1pi",
  "1po",
  "1pu",
  "1py",
  "1ra",
  "1re",
  "1ri",
  "1ro",
  "1ru",
  "1ry",
  "1sa",
  "1se",
  "1si",
  "1so",
  "1su",
  "1sy",
  "1ta",
  "1te",
  "1ti",
  "1to",
  "1tu",
  "1ty",
  "1va",
  "1ve",
  "1vi",
  "1vo",
  "1vu",
  "1vy",
  "1wa",
  "1we",
  "1wi",
  "1wo",
  "1wu",
  "1wy",
  "1za",
  "1ze",
  "1zi",
  "1zo",
  "1zu",
  "1zy",
  "2b1b",
  "2c1c",
  "2d1d",
  "2f1f",
  "2g1g",
  "2m1m",
  "2n1n",
  "2p1p",
  "2r1r",
  "2s1s",
  "2t1t",
  "2ct",
  "2pt",
  "2rt",
  "2st",
  "2ch",
  "2sh",
  "2th",
  "2ph",
  "2gh",
  "3gra",
  "3gre",
  "3gri",
  "3gro",
  "3gru",
  "3pha",
  "3phe",
  "3phi",
  "3pho",
  "3phu",
  "3tra",
  "3tre",
  "3tri",
  "3tro",
  "3tru",
  "3cla",
  "3cle",
  "3cli",
  "3clo",
  "3clu",
  "3pra",
  "3pre",
  "3pri",
  "3pro",
  "3pru",
  "hyphen4",
  "4phen",
  "hy3ph",
];

interface PatternTrie {
  value?: number[];
  children: Map<string, PatternTrie>;
}

function buildTrie(patterns: readonly string[]): PatternTrie {
  const root: PatternTrie = { children: new Map() };
  for (const pat of patterns) {
    let node = root;
    const digits: number[] = [];
    let chars = "";
    let d = 0;
    for (let i = 0; i < pat.length; i++) {
      const ch = pat[i]!;
      if (ch >= "0" && ch <= "9") {
        d = Number(ch);
      } else {
        digits.push(d);
        d = 0;
        chars += ch;
      }
    }
    digits.push(d);

    for (let i = 0; i < chars.length; i++) {
      const c = chars[i]!;
      let next = node.children.get(c);
      if (!next) {
        next = { children: new Map() };
        node.children.set(c, next);
      }
      node = next;
    }
    node.value = digits;
  }
  return root;
}

let enTrie: PatternTrie | null = null;
function getEnTrie(): PatternTrie {
  if (!enTrie) enTrie = buildTrie(EN_PATTERNS);
  return enTrie;
}

/**
 * Hyphenate an English word using Liang's algorithm.
 */
export function hyphenateEnglish(word: string, minPrefix = 2, minSuffix = 2): string {
  if (word.length < minPrefix + minSuffix) return word;
  const trie = getEnTrie();
  const lower = "." + word.toLowerCase() + ".";
  const levels = new Uint8Array(lower.length + 1);

  for (let i = 0; i < lower.length; i++) {
    let node: PatternTrie | undefined = trie;
    for (let j = i; j < lower.length && node; j++) {
      node = node.children.get(lower[j]!);
      if (node?.value) {
        const val = node.value;
        for (let k = 0; k < val.length; k++) {
          const v = val[k]!;
          if (v > levels[i + k]!) {
            levels[i + k] = v;
          }
        }
      }
    }
  }

  let result = "";
  for (let i = 0; i < word.length; i++) {
    if (i >= minPrefix && i <= word.length - minSuffix) {
      // index in lower is i + 1
      if ((levels[i + 1] ?? 0) % 2 === 1) {
        result += SOFT_HYPHEN;
      }
    }
    result += word[i];
  }
  return result;
}

/**
 * Hyphenate a single word for the requested language ('en' or 'id').
 */
export function hyphenateWord(word: string, lang = "en", options?: HyphenationOptions): string {
  const minLen = options?.minWordLength ?? 5;
  if (word.length < minLen) return word;
  if (word.includes(SOFT_HYPHEN) || word.includes("-")) return word;

  const minPrefix = options?.minPrefix ?? 2;
  const minSuffix = options?.minSuffix ?? 2;

  // Check doNotHyphenateCaps
  if (options?.doNotHyphenateCaps && word === word.toUpperCase() && word !== word.toLowerCase()) {
    return word;
  }

  const isId = lang.toLowerCase().startsWith("id");
  if (isId) {
    const syllables = syllabifyIndonesian(word);
    if (syllables.length <= 1) return word;

    let res = syllables[0]!;
    for (let i = 1; i < syllables.length; i++) {
      const syl = syllables[i]!;
      // ensure prefix and suffix constraints
      if (res.replace(/\u00AD/g, "").length >= minPrefix && syl.length >= minSuffix) {
        res += SOFT_HYPHEN + syl;
      } else {
        res += syl;
      }
    }
    return res;
  }

  return hyphenateEnglish(word, minPrefix, minSuffix);
}

/**
 * Automatically insert soft hyphens into words in a text block.
 */
export function hyphenateText(text: string, lang = "en", options?: HyphenationOptions): string {
  if (!text || text.length < 5) return text;
  // Match word boundary sequences of Unicode letters
  return text.replace(/\p{L}{5,}/gu, (match) => {
    return hyphenateWord(match, lang, options);
  });
}
