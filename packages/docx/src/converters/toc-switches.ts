/**
 * OOXML TOC Field Instruction Parser and Generator.
 * Handles switches:
 *  \o : Heading levels range (e.g. "1-3")
 *  \t : Custom styles and levels mapping (e.g. "Title,1,Subtitle,2")
 *  \h : Hyperlinks to headings
 *  \b : Bookmark scope
 *  \u : Outline level from paragraph format
 *  \c : Caption SEQ label for Table of Figures
 *  \n : Omit page numbers
 *  \z : Hide tab leader in web layout
 *  \w : Preserve tab characters
 *  \x : Preserve newline characters
 *  \p : Page number separator
 */

export interface TocSwitches {
  headingRange?: string;
  customStyles?: string;
  hyperlinks?: boolean;
  bookmark?: string;
  outlineLevel?: boolean;
  captionLabel?: string;
  omitPageNumbers?: string | boolean;
  hideTabLeaderInWeb?: boolean;
  preserveTabs?: boolean;
  preserveNewlines?: boolean;
  pageNumberSeparator?: string;
  rawInstruction?: string;
}

/**
 * Tokenizes a TOC field instruction string, taking quotes and escaped characters into account.
 */
export function tokenizeTocInstruction(instruction: string): string[] {
  const tokens: string[] = [];
  let i = 0;
  const len = instruction.length;

  while (i < len) {
    while (i < len && /\s/.test(instruction[i]!)) i++;
    if (i >= len) break;

    if (instruction[i] === '"') {
      i++; // skip opening quote
      let token = "";
      while (i < len && instruction[i] !== '"') {
        if (instruction[i] === "\\" && i + 1 < len && instruction[i + 1] === '"') {
          token += '"';
          i += 2;
        } else {
          token += instruction[i]!;
          i++;
        }
      }
      if (i < len && instruction[i] === '"') i++; // skip closing quote
      tokens.push(token);
    } else {
      let token = "";
      while (i < len && !/\s/.test(instruction[i]!) && instruction[i] !== '"') {
        token += instruction[i]!;
        i++;
      }
      tokens.push(token);
    }
  }

  return tokens;
}

/**
 * Parses a TOC field instruction string into structured TocSwitches.
 */
export function parseTocSwitches(instruction: string): TocSwitches {
  const result: TocSwitches = { rawInstruction: instruction };
  const tokens = tokenizeTocInstruction(instruction);
  let idx = 0;

  if (tokens[0]?.toUpperCase() === "TOC") {
    idx = 1;
  }

  while (idx < tokens.length) {
    const token = tokens[idx]!;
    if (token.startsWith("\\")) {
      const switchChar = token.slice(1, 2).toLowerCase();
      const restOfToken = token.slice(2);
      let arg: string | undefined = restOfToken ? restOfToken.replace(/^"(.*)"$/, "$1") : undefined;

      const consumesArg = ["o", "t", "b", "c", "n", "p"].includes(switchChar);
      if (consumesArg && !arg && idx + 1 < tokens.length && !tokens[idx + 1]!.startsWith("\\")) {
        idx++;
        arg = tokens[idx]!;
      }

      switch (switchChar) {
        case "o":
          result.headingRange = arg || "1-3";
          break;
        case "t":
          result.customStyles = arg || "";
          break;
        case "h":
          result.hyperlinks = true;
          break;
        case "b":
          result.bookmark = arg || "";
          break;
        case "u":
          result.outlineLevel = true;
          break;
        case "c":
          result.captionLabel = arg || "Figure";
          break;
        case "n":
          result.omitPageNumbers = arg || true;
          break;
        case "z":
          result.hideTabLeaderInWeb = true;
          break;
        case "w":
          result.preserveTabs = true;
          break;
        case "x":
          result.preserveNewlines = true;
          break;
        case "p":
          result.pageNumberSeparator = arg || "";
          break;
      }
    }
    idx++;
  }

  return result;
}

/**
 * Generates a standard OOXML TOC field instruction string from TocSwitches.
 */
export function generateTocInstruction(switches: TocSwitches): string {
  const parts: string[] = ["TOC"];

  if (switches.headingRange) {
    parts.push(`\\o "${switches.headingRange}"`);
  }
  if (switches.hyperlinks) {
    parts.push("\\h");
  }
  if (switches.outlineLevel) {
    parts.push("\\u");
  }
  if (switches.hideTabLeaderInWeb) {
    parts.push("\\z");
  }
  if (switches.customStyles) {
    parts.push(`\\t "${switches.customStyles}"`);
  }
  if (switches.bookmark) {
    parts.push(`\\b "${switches.bookmark}"`);
  }
  if (switches.captionLabel) {
    parts.push(`\\c "${switches.captionLabel}"`);
  }
  if (switches.omitPageNumbers) {
    if (typeof switches.omitPageNumbers === "string") {
      parts.push(`\\n "${switches.omitPageNumbers}"`);
    } else {
      parts.push("\\n");
    }
  }
  if (switches.preserveTabs) {
    parts.push("\\w");
  }
  if (switches.preserveNewlines) {
    parts.push("\\x");
  }
  if (switches.pageNumberSeparator) {
    parts.push(`\\p "${switches.pageNumberSeparator}"`);
  }

  return parts.join(" ");
}
