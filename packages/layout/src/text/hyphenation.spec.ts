import { describe, expect, it } from "vitest";

import { hyphenateEnglish, hyphenateText, hyphenateWord, syllabifyIndonesian } from "./hyphenation";

describe("hyphenation engine", () => {
  describe("Indonesian syllabification", () => {
    it("splits Indonesian words according to PUEBI rules", () => {
      expect(syllabifyIndonesian("sekolah")).toEqual(["se", "ko", "lah"]);
      expect(syllabifyIndonesian("pembahasan")).toEqual(["pem", "ba", "ha", "san"]);
      expect(syllabifyIndonesian("daun")).toEqual(["da", "un"]);
      expect(syllabifyIndonesian("bangun")).toEqual(["ba", "ngun"]);
    });

    it("inserts soft hyphens for Indonesian words", () => {
      const hyp = hyphenateWord("pembahasan", "id");
      expect(hyp).toContain("\u00AD");
      expect(hyp.replace(/\u00AD/g, "")).toBe("pembahasan");
    });
  });

  describe("English hyphenation (Liang patterns)", () => {
    it("hyphenates English words at syllable breaks", () => {
      const res = hyphenateEnglish("hyphenation");
      expect(res).toContain("\u00AD");
      expect(res.replace(/\u00AD/g, "")).toBe("hyphenation");

      const para = hyphenateEnglish("paragraph");
      expect(para).toContain("\u00AD");
      expect(para.replace(/\u00AD/g, "")).toBe("paragraph");
    });

    it("respects minimum prefix and suffix lengths", () => {
      const res = hyphenateEnglish("international", 3, 3);
      const parts = res.split("\u00AD");
      expect(parts[0]!.length).toBeGreaterThanOrEqual(3);
      expect(parts[parts.length - 1]!.length).toBeGreaterThanOrEqual(3);
    });
  });

  describe("options and full text hyphenation", () => {
    it("skips all-caps words when doNotHyphenateCaps is true", () => {
      const caps = hyphenateWord("IMPORTANT", "en", { doNotHyphenateCaps: true });
      expect(caps).toBe("IMPORTANT");

      const normal = hyphenateWord("important", "en", { doNotHyphenateCaps: true });
      expect(normal).toContain("\u00AD");
    });

    it("skips words shorter than minWordLength", () => {
      const short = hyphenateWord("word", "en", { minWordLength: 5 });
      expect(short).toBe("word");
    });

    it("hyphenates long words in text without disturbing punctuation or short words", () => {
      const text = "The international conference was very important for all participants.";
      const res = hyphenateText(text, "en");
      expect(res.replace(/\u00AD/g, "")).toBe(text);
      expect(res).toContain("\u00AD");
      // "was" and "for" and "all" should not be touched
      expect(res).toContain(" was ");
      expect(res).toContain(" for ");
      expect(res).toContain(" all ");
    });
  });
});
