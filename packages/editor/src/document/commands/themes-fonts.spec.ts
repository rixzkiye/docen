import { describe, expect, it } from "vitest";

import { detectMissingFonts, getFallbackFontChain, resolveFontSubstitution } from "./fonts";
import { getColorScheme, getFontScheme, getTheme, THEMES } from "./themes";

describe("themes and fonts", () => {
  it("provides built-in Word themes with 12-color palettes and font pairs", () => {
    expect(THEMES.office).toBeDefined();
    expect(THEMES.facet).toBeDefined();
    expect(THEMES.integral).toBeDefined();

    const office = getTheme("office");
    expect(office.colors.accent1).toBe("4f81bd");
    expect(office.fonts.majorFont).toBe("Aptos Display");
    expect(office.fonts.minorFont).toBe("Aptos");

    const color = getColorScheme("facet");
    expect(color.accent1).toBe("91a14f");

    const fonts = getFontScheme("calibri");
    expect(fonts.majorFont).toBe("Calibri Light");
    expect(fonts.minorFont).toBe("Calibri");
  });

  it("substitutes Office fonts correctly to cross-platform fallbacks", () => {
    const aptosChain = resolveFontSubstitution("Aptos");
    expect(aptosChain).toContain("Calibri");
    expect(aptosChain).toContain("Arial");
    expect(aptosChain).toContain("sans-serif");

    const simsunChain = resolveFontSubstitution("SimSun");
    expect(simsunChain).toContain("Songti SC");

    const cssString = getFallbackFontChain("Calibri");
    expect(cssString).toContain('"Calibri"');
    expect(cssString).toContain('"Arial"');
  });

  it("detects missing fonts when compared against available system fonts", () => {
    const docFonts = ["Aptos", "Calibri", "Comic Sans Custom"];
    const available = ["Aptos", "Arial", "Calibri"];
    const missing = detectMissingFonts(docFonts, available);
    expect(missing).toEqual(["Comic Sans Custom"]);
  });
});
