/**
 * Offline Thesaurus Engine: provides synonym and antonym lookups
 * for English and Indonesian vocabulary.
 */

import type { ThesaurusEntry, ThesaurusMeaning } from "./types";

const EN_THESAURUS_DATA: Record<string, ThesaurusEntry> = {
  good: {
    word: "good",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Having desirable or positive qualities",
        synonyms: ["great", "fine", "excellent", "superb", "fabulous", "pleasant", "decent"],
        antonyms: ["bad", "poor", "inferior"],
      },
      {
        partOfSpeech: "noun",
        definition: "That which is morally right or beneficial",
        synonyms: ["benefit", "advantage", "virtue", "morality"],
        antonyms: ["evil", "disadvantage"],
      },
    ],
  },
  bad: {
    word: "bad",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Of poor quality or undesirable",
        synonyms: ["poor", "terrible", "awful", "dreadful", "inferior", "substandard"],
        antonyms: ["good", "excellent"],
      },
    ],
  },
  big: {
    word: "big",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Of considerable size or extent",
        synonyms: ["large", "huge", "enormous", "massive", "gigantic", "immense", "substantial"],
        antonyms: ["small", "tiny", "little"],
      },
    ],
  },
  small: {
    word: "small",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Of limited size; of comparatively little size",
        synonyms: ["little", "tiny", "compact", "miniature", "minute", "slight"],
        antonyms: ["big", "large", "huge"],
      },
    ],
  },
  happy: {
    word: "happy",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Feeling or showing pleasure or contentment",
        synonyms: ["glad", "joyful", "cheerful", "delighted", "pleased", "content", "ecstatic"],
        antonyms: ["sad", "unhappy", "sorrowful", "depressed"],
      },
    ],
  },
  sad: {
    word: "sad",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Feeling or showing sorrow; unhappy",
        synonyms: ["unhappy", "sorrowful", "gloomy", "depressed", "melancholy", "downcast"],
        antonyms: ["happy", "cheerful", "glad"],
      },
    ],
  },
  fast: {
    word: "fast",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Moving or capable of moving at high speed",
        synonyms: ["quick", "rapid", "swift", "speedy", "brisk", "prompt"],
        antonyms: ["slow", "sluggish"],
      },
      {
        partOfSpeech: "adv",
        definition: "At high speed; quickly",
        synonyms: ["quickly", "rapidly", "speedily", "swiftly"],
        antonyms: ["slowly"],
      },
    ],
  },
  slow: {
    word: "slow",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Moving or operating, or designed to do so, only at a low speed",
        synonyms: ["sluggish", "unhurried", "leisurely", "gradual", "delayed"],
        antonyms: ["fast", "quick", "rapid"],
      },
    ],
  },
  important: {
    word: "important",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Of great significance or value",
        synonyms: ["significant", "crucial", "vital", "essential", "critical", "key", "paramount"],
        antonyms: ["unimportant", "trivial", "minor"],
      },
    ],
  },
  beautiful: {
    word: "beautiful",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Pleasing the senses or mind aesthetically",
        synonyms: [
          "attractive",
          "pretty",
          "gorgeous",
          "lovely",
          "stunning",
          "handsome",
          "alluring",
        ],
        antonyms: ["ugly", "unattractive"],
      },
    ],
  },
  help: {
    word: "help",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Make it easier for (someone) to do something",
        synonyms: ["assist", "aid", "support", "serve", "relieve", "back"],
        antonyms: ["hinder", "impede", "prevent"],
      },
      {
        partOfSpeech: "noun",
        definition: "The action of helping someone to do something",
        synonyms: ["assistance", "aid", "support", "backing", "relief"],
        antonyms: ["hindrance", "obstruction"],
      },
    ],
  },
  start: {
    word: "start",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Begin or be reckoned from a particular point in time or space",
        synonyms: ["begin", "commence", "initiate", "launch", "open", "originate"],
        antonyms: ["finish", "end", "stop", "conclude"],
      },
    ],
  },
  begin: {
    word: "begin",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Perform or undergo the first part of (an action or activity)",
        synonyms: ["start", "commence", "initiate", "launch", "undertake"],
        antonyms: ["end", "stop", "finish"],
      },
    ],
  },
  stop: {
    word: "stop",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Come to an end; cease to happen",
        synonyms: ["cease", "halt", "end", "pause", "terminate", "discontinue"],
        antonyms: ["start", "continue", "resume"],
      },
    ],
  },
  make: {
    word: "make",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Form by putting parts together or combining substances",
        synonyms: ["create", "build", "construct", "produce", "generate", "fabricate", "form"],
        antonyms: ["destroy", "dismantle"],
      },
    ],
  },
  change: {
    word: "change",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Make or become different",
        synonyms: ["alter", "modify", "transform", "convert", "vary", "adjust"],
        antonyms: ["preserve", "maintain"],
      },
      {
        partOfSpeech: "noun",
        definition: "The act or instance of making or becoming different",
        synonyms: ["alteration", "modification", "transformation", "shift", "transition"],
        antonyms: ["stagnation", "permanence"],
      },
    ],
  },
  think: {
    word: "think",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Have a particular belief or idea",
        synonyms: [
          "believe",
          "consider",
          "ponder",
          "reflect",
          "contemplate",
          "reckon",
          "deliberate",
        ],
      },
    ],
  },
  show: {
    word: "show",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Allow or cause (something) to be visible",
        synonyms: [
          "display",
          "exhibit",
          "demonstrate",
          "reveal",
          "present",
          "indicate",
          "manifest",
        ],
        antonyms: ["hide", "conceal"],
      },
    ],
  },
  find: {
    word: "find",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Discover or perceive by chance or unexpectedly",
        synonyms: ["discover", "locate", "encounter", "detect", "uncover", "identify"],
        antonyms: ["lose", "misplace"],
      },
    ],
  },
  easy: {
    word: "easy",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Achieved without great effort; presenting few difficulties",
        synonyms: ["simple", "effortless", "straightforward", "uncomplicated", "painless"],
        antonyms: ["difficult", "hard", "complicated"],
      },
    ],
  },
  hard: {
    word: "hard",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Requiring a great deal of endurance or effort",
        synonyms: ["difficult", "challenging", "tough", "demanding", "arduous", "laborious"],
        antonyms: ["easy", "simple", "effortless"],
      },
    ],
  },
};

const ID_THESAURUS_DATA: Record<string, ThesaurusEntry> = {
  baik: {
    word: "baik",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Elok, patut, teratur, memberi faedah",
        synonyms: ["bagus", "elok", "positif", "unggul", "mantap", "prima"],
        antonyms: ["buruk", "jelek", "jahat"],
      },
    ],
  },
  buruk: {
    word: "buruk",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Tidak baik, jelek, tercela",
        synonyms: ["jelek", "cacat", "rusak", "negatif", "tercela"],
        antonyms: ["baik", "bagus", "elok"],
      },
    ],
  },
  besar: {
    word: "besar",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Lebih dari ukuran sedang; luas; tinggi",
        synonyms: ["luas", "akbar", "agung", "raksasa", "masif", "raya"],
        antonyms: ["kecil", "sempit"],
      },
    ],
  },
  kecil: {
    word: "kecil",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Kurang dari ukuran biasa",
        synonyms: ["mini", "mungil", "sempit", "sedikit", "ringan"],
        antonyms: ["besar", "luas", "raksasa"],
      },
    ],
  },
  cepat: {
    word: "cepat",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Dalam waktu singkat; lekas; tangkas",
        synonyms: ["lekas", "kilat", "gesit", "tangkas", "segera", "deras"],
        antonyms: ["lambat", "pelan", "lamban"],
      },
    ],
  },
  lambat: {
    word: "lambat",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Tidak cepat; perlahan-lahan",
        synonyms: ["pelan", "lamban", "alun", "tertunda"],
        antonyms: ["cepat", "lekas", "kilat"],
      },
    ],
  },
  senang: {
    word: "senang",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Puas hatinya; merasa gembira",
        synonyms: ["bahagia", "gembira", "suka", "lega", "girang", "ceria"],
        antonyms: ["sedih", "duka", "susah"],
      },
    ],
  },
  sedih: {
    word: "sedih",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Merasa susah hati; berduka cita",
        synonyms: ["duka", "pilu", "susah", "murung", "masygul"],
        antonyms: ["senang", "gembira", "bahagia"],
      },
    ],
  },
  penting: {
    word: "penting",
    meanings: [
      {
        partOfSpeech: "adj",
        definition: "Sangat berharga, berpengaruh, atau utama",
        synonyms: ["utama", "krusial", "vital", "esensial", "signifikan", "pokok"],
        antonyms: ["sepele", "remeh"],
      },
    ],
  },
  buat: {
    word: "buat",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Mengerjakan atau menghasilkan sesuatu",
        synonyms: ["bikin", "cipta", "susun", "hasilkan", "bangun", "produksi"],
        antonyms: ["hancurkan", "rusak"],
      },
    ],
  },
  mulai: {
    word: "mulai",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Mengawali atau bergerak dari permulaan",
        synonyms: ["awali", "buka", "rintis", "langkah", "start"],
        antonyms: ["selesai", "akhir", "henti", "tutup"],
      },
    ],
  },
  bantu: {
    word: "bantu",
    meanings: [
      {
        partOfSpeech: "verb",
        definition: "Menolong atau memberi sokongan",
        synonyms: ["tolong", "sokong", "dukung", "topang", "iringi"],
        antonyms: ["halang", "hambat", "ganggu"],
      },
    ],
  },
};

/** Look up a word in the offline thesaurus. */
export function lookupThesaurus(word: string, lang = "en"): ThesaurusEntry | null {
  const normWord = word.trim().toLowerCase();
  if (!normWord) return null;

  const isId = lang.toLowerCase().startsWith("id");
  const dataset = isId ? ID_THESAURUS_DATA : EN_THESAURUS_DATA;
  if (dataset[normWord]) {
    return dataset[normWord];
  }

  // Fallback: reverse lookup across synonyms
  const matchedSynonyms = new Set<string>();
  let fallbackPos: ThesaurusMeaning["partOfSpeech"] = "other";
  for (const entry of Object.values(dataset)) {
    for (const m of entry.meanings) {
      if (m.synonyms.some((s) => s.toLowerCase() === normWord)) {
        fallbackPos = m.partOfSpeech;
        matchedSynonyms.add(entry.word);
        for (const s of m.synonyms) {
          if (s.toLowerCase() !== normWord) matchedSynonyms.add(s);
        }
      }
    }
  }

  if (matchedSynonyms.size > 0) {
    return {
      word: normWord,
      meanings: [
        {
          partOfSpeech: fallbackPos,
          synonyms: Array.from(matchedSynonyms),
        },
      ],
    };
  }

  return null;
}

/** Get a flat list of synonyms for a given word. */
export function getSynonyms(word: string, lang = "en", limit = 6): string[] {
  const entry = lookupThesaurus(word, lang);
  if (!entry) return [];
  const list: string[] = [];
  for (const m of entry.meanings) {
    for (const syn of m.synonyms) {
      if (!list.includes(syn) && syn.toLowerCase() !== word.toLowerCase()) {
        list.push(syn);
      }
    }
  }
  return list.slice(0, limit);
}
