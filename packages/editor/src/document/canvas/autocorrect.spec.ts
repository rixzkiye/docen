import { describe, expect, it } from "vitest";

import { defaultSettings } from "../settings";
import {
  autocorrectConfigOf,
  autocorrectOf,
  correctionOf,
  DEFAULT_AUTOCORRECT,
  DEFAULT_AUTOCORRECT_REPLACEMENTS,
  hyperlinkAtEnd,
  hyperlinkFix,
  smartQuoteOf,
  type AutocorrectConfig,
} from "./autocorrect";

/** A config with one rule toggled (defaults fill the rest). */
const cfg = (patch: Partial<AutocorrectConfig> = {}): AutocorrectConfig => ({
  ...DEFAULT_AUTOCORRECT,
  ...patch,
});

describe("smartQuoteOf", () => {
  it("opens after whitespace / start / opening punctuation", () => {
    expect(smartQuoteOf('"', "")).toBe("“");
    expect(smartQuoteOf('"', " ")).toBe("“");
    expect(smartQuoteOf("'", "\t")).toBe("‘");
    expect(smartQuoteOf('"', "(")).toBe("“");
    expect(smartQuoteOf('"', "—")).toBe("“");
  });

  it("closes after word characters and punctuation", () => {
    expect(smartQuoteOf('"', "o")).toBe("”");
    expect(smartQuoteOf("'", "n")).toBe("’");
    expect(smartQuoteOf('"', ",")).toBe("”");
    expect(smartQuoteOf('"', "”")).toBe("”");
  });

  it("leaves non-quotes alone", () => {
    expect(smartQuoteOf("a", "")).toBeNull();
    expect(smartQuoteOf("-", " ")).toBeNull();
  });
});

describe("correctionOf", () => {
  it("fixes known words and preserves case shape", () => {
    expect(correctionOf("teh")).toBe("the");
    expect(correctionOf("Teh")).toBe("The");
    expect(correctionOf("TEH")).toBe("THE");
    expect(correctionOf("recieve")).toBe("receive");
  });

  it("leaves unknown words alone", () => {
    expect(correctionOf("hello")).toBeNull();
  });
});

describe("autocorrectOf", () => {
  it("rewrites the word behind a boundary character, boundary trailing", () => {
    // "teh" + typed " " → "the " — the space stays as the word separator
    expect(autocorrectOf(" ", "teh")).toEqual({ text: "the ", back: 3 });
    expect(autocorrectOf(",", "Teh")).toEqual({ text: "The,", back: 3 });
  });

  it("replaces the quote in place with its curly side", () => {
    expect(autocorrectOf('"', "the ")).toEqual({ text: "“", back: 0 });
    expect(autocorrectOf("'", "don")).toEqual({ text: "’", back: 0 });
  });

  it("builds the em dash from the hyphen behind and ellipsis from two dots", () => {
    expect(autocorrectOf("-", "a-")).toEqual({ text: "—", back: 1 });
    expect(autocorrectOf(".", "a..")).toEqual({ text: "…", back: 2 });
  });

  it("returns null for ordinary characters", () => {
    expect(autocorrectOf("x", "teh ")).toBeNull();
    expect(autocorrectOf(" ", "hello")).toBeNull();
    expect(autocorrectOf("-", "a ")).toBeNull();
  });
});

describe("rule toggles", () => {
  it("leaves each rule off when its setting is off", () => {
    expect(autocorrectOf('"', "the ", cfg({ smartQuotes: false }))).toBeNull();
    expect(autocorrectOf("-", "a-", cfg({ emDash: false }))).toBeNull();
    expect(autocorrectOf(" ", "a -", cfg({ emDash: false }))).toBeNull();
    expect(autocorrectOf(".", "a..", cfg({ ellipsis: false }))).toBeNull();
    expect(autocorrectOf(" ", "teh", cfg({ replacements: [] }))).toBeNull();
    expect(autocorrectOf("t", "1s", cfg({ ordinalSuperscript: false }))).toBeNull();
    expect(autocorrectOf("w", "Hi. ", cfg({ capitalizeFirstLetter: false }))).toBeNull();
    expect(
      autocorrectOf(" ", "https://example.com", cfg({ hyperlinkAutoformat: false })),
    ).toBeNull();
  });

  it("keeps the default table at Word's 12 built-ins", () => {
    expect(DEFAULT_AUTOCORRECT_REPLACEMENTS).toHaveLength(12);
    expect(DEFAULT_AUTOCORRECT_REPLACEMENTS).toContainEqual({ from: "teh", to: "the" });
  });
});

describe("en dash", () => {
  it("turns a spaced hyphen completed by a space into an en dash", () => {
    expect(autocorrectOf(" ", "a -", cfg())).toEqual({ text: "– ", back: 1 });
    expect(autocorrectOf(" ", "word -", cfg())).toEqual({ text: "– ", back: 1 });
  });

  it("does not fire at the paragraph start or without a word before the hyphen", () => {
    expect(autocorrectOf(" ", " -", cfg())).toBeNull();
    expect(autocorrectOf(" ", "-", cfg())).toBeNull();
  });

  it("leaves a double hyphen to the em dash rule", () => {
    expect(autocorrectOf("-", "a --", cfg())).toEqual({ text: "—", back: 1 });
    expect(autocorrectOf(" ", "a --", cfg())).toBeNull();
  });
});

describe("ordinal superscript", () => {
  const config = cfg({ ordinalSuperscript: true });

  it("superscripts the completing suffix on valid ordinals", () => {
    expect(autocorrectOf("t", "1s", config)).toEqual({
      text: "st",
      back: 1,
      marks: [{ name: "superscript", start: 0, length: 2 }],
    });
    for (const [typed, before, suffix] of [
      ["d", "2n", "nd"],
      ["d", "3r", "rd"],
      ["h", "4t", "th"],
      ["h", "0t", "th"],
      ["h", "11t", "th"],
      ["t", "21s", "st"],
      ["d", "103r", "rd"],
    ] as const) {
      expect(autocorrectOf(typed, before, config)).toEqual({
        text: suffix,
        back: 1,
        marks: [{ name: "superscript", start: 0, length: 2 }],
      });
    }
  });

  it("rejects a suffix that does not match the number", () => {
    expect(autocorrectOf("h", "1t", config)).toBeNull();
    expect(autocorrectOf("d", "2r", config)).toBeNull();
    expect(autocorrectOf("t", "3s", config)).toBeNull();
  });

  it("is off by default", () => {
    expect(autocorrectOf("t", "1s")).toBeNull();
  });

  it("does not fire mid-word or on multi-character commits", () => {
    expect(autocorrectOf("t", "abc1s", config)).toBeNull();
    expect(autocorrectOf("st", "1", config)).toBeNull();
    expect(autocorrectOf("1", "1s", config)).toBeNull();
  });
});

describe("sentence capitalization", () => {
  it("capitalizes the first letter of a paragraph and of a sentence", () => {
    expect(autocorrectOf("w", "")).toEqual({ text: "W", back: 0 });
    expect(autocorrectOf("w", "   ")).toEqual({ text: "W", back: 0 });
    expect(autocorrectOf("w", "Hello. ")).toEqual({ text: "W", back: 0 });
    expect(autocorrectOf("w", "Done! ")).toEqual({ text: "W", back: 0 });
    expect(autocorrectOf("w", "Why? ")).toEqual({ text: "W", back: 0 });
    expect(autocorrectOf("w", "Wait… ")).toEqual({ text: "W", back: 0 });
    expect(autocorrectOf("w", 'He said "stop." ')).toEqual({ text: "W", back: 0 });
  });

  it("leaves mid-sentence and mid-word letters alone", () => {
    expect(autocorrectOf("w", "hello ")).toBeNull();
    expect(autocorrectOf("w", "hello.")).toBeNull();
    expect(autocorrectOf("w", "hello.w")).toBeNull();
    expect(autocorrectOf("W", "Hello. ")).toBeNull();
    expect(autocorrectOf("wo", "Hello. ")).toBeNull();
  });

  it("skips after common abbreviations and multi-dot initials", () => {
    for (const before of ["Dr. ", "etc. ", "vs. ", "Mr. ", "e.g. ", "i.e. ", "U.S. "]) {
      expect(autocorrectOf("s", before)).toBeNull();
    }
  });

  it("honors the user exceptions before the terminator", () => {
    expect(autocorrectOf("s", "zzz. ", cfg({ exceptions: ["zzz"] }))).toBeNull();
    expect(autocorrectOf("s", "zzz. ", cfg())).toEqual({ text: "S", back: 0 });
  });
});

describe("replacement table", () => {
  it("applies a custom table and preserves the case shape", () => {
    const config = cfg({ replacements: [{ from: "fo", to: "foo" }] });
    expect(autocorrectOf(" ", "fo", config)).toEqual({ text: "foo ", back: 2 });
    expect(autocorrectOf(" ", "Fo", config)).toEqual({ text: "Foo ", back: 2 });
    expect(autocorrectOf(" ", "FO", config)).toEqual({ text: "FOO ", back: 2 });
    expect(correctionOf("fo", config)).toBe("foo");
  });

  it("does not fall back to the built-ins for an explicitly empty table", () => {
    expect(autocorrectOf(" ", "teh", cfg({ replacements: [] }))).toBeNull();
    expect(correctionOf("teh", cfg({ replacements: [] }))).toBeNull();
  });

  it("leaves exception words alone, case-insensitively", () => {
    const config = cfg({ exceptions: ["Teh"] });
    expect(autocorrectOf(" ", "teh", config)).toBeNull();
    expect(autocorrectOf(" ", "TEH", config)).toBeNull();
    expect(correctionOf("Teh", config)).toBeNull();
  });

  it("does not rewrite words glued into a larger token", () => {
    expect(autocorrectOf(" ", "foo.teh")).toBeNull();
    expect(autocorrectOf(" ", "path/teh")).toBeNull();
    expect(autocorrectOf(" ", "a-teh")).toBeNull();
    expect(autocorrectOf(" ", "mail@teh")).toBeNull();
  });

  it("still corrects after opening punctuation", () => {
    expect(autocorrectOf(" ", "(teh")).toEqual({ text: "the ", back: 3 });
  });
});

describe("hyperlink autoformat", () => {
  it("linkifies a scheme URL completed by a typed space", () => {
    const fix = autocorrectOf(" ", "see https://example.com");
    expect(fix?.text).toBe(" ");
    expect(fix?.marks).toEqual([
      { name: "link", attrs: { href: "https://example.com" }, start: -19, length: 19 },
      { name: "textStyle", attrs: { style: "Hyperlink" }, start: -19, length: 19 },
    ]);
  });

  it("recognizes www hosts, emails and UNC paths with the right href", () => {
    expect(hyperlinkAtEnd("visit www.example.com")).toEqual({
      href: "https://www.example.com",
      text: "www.example.com",
    });
    expect(hyperlinkAtEnd("mail me@example.com")).toEqual({
      href: "mailto:me@example.com",
      text: "me@example.com",
    });
    expect(hyperlinkAtEnd(String.raw`open \\server\share`)).toEqual({
      href: String.raw`\\server\share`,
      text: String.raw`\\server\share`,
    });
  });

  it("strips sentence punctuation and wrapper brackets off the URL tail", () => {
    expect(hyperlinkAtEnd("see https://example.com.")).toEqual({
      href: "https://example.com",
      text: "https://example.com",
    });
    expect(hyperlinkAtEnd("see (https://example.com)")).toEqual({
      href: "https://example.com",
      text: "https://example.com",
    });
    expect(hyperlinkAtEnd("see https://en.wikipedia.org/wiki/Foo_(bar)")).toEqual({
      href: "https://en.wikipedia.org/wiki/Foo_(bar)",
      text: "https://en.wikipedia.org/wiki/Foo_(bar)",
    });
  });

  it("does not linkify ordinary words or bare file names", () => {
    expect(hyperlinkAtEnd("hello")).toBeNull();
    expect(hyperlinkAtEnd("document.docx")).toBeNull();
    expect(autocorrectOf(" ", "hello")).toBeNull();
  });

  it("does not re-linkify behind a space", () => {
    expect(autocorrectOf(" ", "https://example.com ")).toBeNull();
  });

  it("respects the toggle", () => {
    const off = cfg({ hyperlinkAutoformat: false });
    expect(autocorrectOf(" ", "https://example.com", off)).toBeNull();
    expect(hyperlinkFix("https://example.com", off)).toBeNull();
  });

  it("offers the Enter leg a marks-only fix", () => {
    expect(hyperlinkFix("www.example.com")).toEqual({
      text: "",
      back: 0,
      marks: [
        { name: "link", attrs: { href: "https://www.example.com" }, start: -15, length: 15 },
        { name: "textStyle", attrs: { style: "Hyperlink" }, start: -15, length: 15 },
      ],
    });
  });

  it("leaves URL characters to the symbol rules", () => {
    expect(autocorrectOf("-", "https://a.com/x-")).toBeNull();
    expect(autocorrectOf(".", "https://a.com/x..")).toBeNull();
    expect(autocorrectOf("-", "user@host-")).toBeNull();
  });
});

describe("multi-character safety", () => {
  it("never rewrites a multi-character commit (IME / autocomplete)", () => {
    expect(autocorrectOf("teh", "so ")).toBeNull();
    expect(autocorrectOf("hello", "Dr. ")).toBeNull();
    expect(autocorrectOf("st", "1")).toBeNull();
    expect(autocorrectOf("--", "a")).toBeNull();
    expect(autocorrectOf("..", "a")).toBeNull();
  });
});

describe("autocorrectConfigOf", () => {
  it("projects the persisted settings onto the rule config", () => {
    const settings = { ...defaultSettings().writing.autocorrect, ordinalSuperscript: true };
    const config = autocorrectConfigOf(settings);
    expect(config.replacements).toBe(DEFAULT_AUTOCORRECT_REPLACEMENTS);
    expect(config.exceptions).toEqual([]);
    expect(config.ordinalSuperscript).toBe(true);
    expect(config.smartQuotes).toBe(true);
  });

  it("uses the user table when present — even an explicitly empty one", () => {
    const base = defaultSettings().writing.autocorrect;
    const customized = autocorrectConfigOf({
      ...base,
      table: {
        version: 1,
        replacements: [{ from: "x", to: "y" }],
        exceptions: ["z"],
      },
    });
    expect(customized.replacements).toEqual([{ from: "x", to: "y" }]);
    expect(customized.exceptions).toEqual(["z"]);
    const emptied = autocorrectConfigOf({
      ...base,
      table: { version: 1, replacements: [], exceptions: [] },
    });
    expect(emptied.replacements).toEqual([]);
  });
});
