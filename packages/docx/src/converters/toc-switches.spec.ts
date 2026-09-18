import { describe, expect, it } from "vitest";

import {
  tokenizeTocInstruction,
  parseTocSwitches,
  generateTocInstruction,
  type TocSwitches,
} from "../index";

describe("W9.4 TOC Switches & Instruction Parser", () => {
  describe("tokenizeTocInstruction", () => {
    it("splits instructions respecting quoted arguments", () => {
      const instruction = 'TOC \\o "1-3" \\h \\z \\t "Title,1,Subtitle,2"';
      const tokens = tokenizeTocInstruction(instruction);
      expect(tokens).toEqual(["TOC", "\\o", "1-3", "\\h", "\\z", "\\t", "Title,1,Subtitle,2"]);
    });

    it('handles attached quotes (\\o"1-4") correctly', () => {
      const instruction = 'TOC \\o"1-4" \\u';
      const tokens = tokenizeTocInstruction(instruction);
      expect(tokens).toContain("TOC");
      expect(tokens).toContain("\\o");
      expect(tokens).toContain("1-4");
      expect(tokens).toContain("\\u");
    });
  });

  describe("parseTocSwitches", () => {
    it("parses all major switches: \\o, \\t, \\h, \\b, \\u, \\c, \\n, \\z, \\w, \\x, \\p", () => {
      const instruction =
        'TOC \\o "1-4" \\t "HeadingA,1,HeadingB,2" \\h \\b "ScopeBookmark" \\u \\c "Figure" \\n "1-2" \\z \\w \\x \\p " - "';
      const switches = parseTocSwitches(instruction);

      expect(switches.headingRange).toBe("1-4");
      expect(switches.customStyles).toBe("HeadingA,1,HeadingB,2");
      expect(switches.hyperlinks).toBe(true);
      expect(switches.bookmark).toBe("ScopeBookmark");
      expect(switches.outlineLevel).toBe(true);
      expect(switches.captionLabel).toBe("Figure");
      expect(switches.omitPageNumbers).toBe("1-2");
      expect(switches.hideTabLeaderInWeb).toBe(true);
      expect(switches.preserveTabs).toBe(true);
      expect(switches.preserveNewlines).toBe(true);
      expect(switches.pageNumberSeparator).toBe(" - ");
    });

    it("defaults \\o to 1-3 when switch present without arg", () => {
      const instruction = "TOC \\o \\h";
      const switches = parseTocSwitches(instruction);
      expect(switches.headingRange).toBe("1-3");
      expect(switches.hyperlinks).toBe(true);
    });
  });

  describe("generateTocInstruction", () => {
    it("generates formatted OOXML TOC instruction from TocSwitches object", () => {
      const switches: TocSwitches = {
        headingRange: "1-3",
        hyperlinks: true,
        outlineLevel: true,
        customStyles: "Style1,1",
        bookmark: "MyBookmark",
        captionLabel: "Table",
        hideTabLeaderInWeb: true,
        preserveTabs: true,
      };

      const instruction = generateTocInstruction(switches);
      expect(instruction).toContain('TOC \\o "1-3"');
      expect(instruction).toContain("\\h");
      expect(instruction).toContain("\\u");
      expect(instruction).toContain("\\z");
      expect(instruction).toContain('\\t "Style1,1"');
      expect(instruction).toContain('\\b "MyBookmark"');
      expect(instruction).toContain('\\c "Table"');
      expect(instruction).toContain("\\w");

      // Verify it round-trips through parseTocSwitches
      const reparsed = parseTocSwitches(instruction);
      expect(reparsed.headingRange).toBe("1-3");
      expect(reparsed.hyperlinks).toBe(true);
      expect(reparsed.outlineLevel).toBe(true);
      expect(reparsed.customStyles).toBe("Style1,1");
      expect(reparsed.bookmark).toBe("MyBookmark");
      expect(reparsed.captionLabel).toBe("Table");
      expect(reparsed.hideTabLeaderInWeb).toBe(true);
      expect(reparsed.preserveTabs).toBe(true);
    });
  });
});
