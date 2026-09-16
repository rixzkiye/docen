import type { JSONContent } from "../core";

interface RtfState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  subscript: boolean;
  superscript: boolean;
  fontSize?: number; // in half-points
  colorIndex?: number;
  highlightIndex?: number;
  fontIndex?: number;
  align?: "left" | "center" | "right" | "justify";
}

interface ColorRGB {
  r: number;
  g: number;
  b: number;
}

function parseHexColor(hex: string): ColorRGB | null {
  const clean = hex.replace(/^#/, "").trim();
  if (clean.length === 3) {
    const r = parseInt(clean[0]! + clean[0]!, 16);
    const g = parseInt(clean[1]! + clean[1]!, 16);
    const b = parseInt(clean[2]! + clean[2]!, 16);
    return isNaN(r) || isNaN(g) || isNaN(b) ? null : { r, g, b };
  }
  if (clean.length === 6) {
    const r = parseInt(clean.slice(0, 2), 16);
    const g = parseInt(clean.slice(2, 4), 16);
    const b = parseInt(clean.slice(4, 6), 16);
    return isNaN(r) || isNaN(g) || isNaN(b) ? null : { r, g, b };
  }
  const rgbMatch = /^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i.exec(clean);
  if (rgbMatch) {
    return {
      r: Math.min(255, Math.max(0, parseInt(rgbMatch[1]!, 10))),
      g: Math.min(255, Math.max(0, parseInt(rgbMatch[2]!, 10))),
      b: Math.min(255, Math.max(0, parseInt(rgbMatch[3]!, 10))),
    };
  }
  return null;
}

function escapeRtfText(text: string): string {
  let out = "";
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    const ch = text[i]!;
    if (ch === "\\") {
      out += "\\\\";
    } else if (ch === "{") {
      out += "\\{";
    } else if (ch === "}") {
      out += "\\}";
    } else if (code < 128) {
      out += ch;
    } else {
      // RTF 16-bit signed unicode character
      const signed = code > 32767 ? code - 65536 : code;
      out += `\\u${signed}?`;
    }
  }
  return out;
}

/**
 * Generate a valid RTF 1.5 document string from Tiptap JSONContent.
 */
export function generateRTF(doc: JSONContent | JSONContent[]): string {
  const nodes = Array.isArray(doc) ? doc : (doc.content ?? []);
  const colorList: ColorRGB[] = [];
  const colorMap = new Map<string, number>();

  const getColorIndex = (colorStr?: string): number | undefined => {
    if (!colorStr) return undefined;
    const parsed = parseHexColor(colorStr);
    if (!parsed) return undefined;
    const key = `${parsed.r},${parsed.g},${parsed.b}`;
    let idx = colorMap.get(key);
    if (idx === undefined) {
      colorList.push(parsed);
      idx = colorList.length; // 1-indexed for RTF \cf1...
      colorMap.set(key, idx);
    }
    return idx;
  };

  // First pass: collect colors from nodes
  const collectColors = (items: readonly JSONContent[]): void => {
    for (const node of items) {
      if (node.marks) {
        for (const m of node.marks) {
          if (m.attrs?.color) getColorIndex(m.attrs.color);
        }
      }
      if (node.content) collectColors(node.content);
    }
  };
  collectColors(nodes);

  let body = "";

  const renderInline = (inline: JSONContent): string => {
    if (inline.type === "hardBreak") return "\\line\n";
    if (inline.type !== "text" || !inline.text) return "";

    const text = escapeRtfText(inline.text);
    let prefix = "";
    let suffix = "";

    if (inline.marks) {
      for (const m of inline.marks) {
        switch (m.type) {
          case "bold":
            prefix += "\\b ";
            suffix = "\\b0 " + suffix;
            break;
          case "italic":
            prefix += "\\i ";
            suffix = "\\i0 " + suffix;
            break;
          case "underline":
            prefix += "\\ul ";
            suffix = "\\ulnone " + suffix;
            break;
          case "strike":
            prefix += "\\strike ";
            suffix = "\\strike0 " + suffix;
            break;
          case "subscript":
            prefix += "\\sub ";
            suffix = "\\nosupersub " + suffix;
            break;
          case "superscript":
            prefix += "\\super ";
            suffix = "\\nosupersub " + suffix;
            break;
          case "textStyle":
            if (m.attrs?.color) {
              const cIdx = getColorIndex(m.attrs.color);
              if (cIdx != null) {
                prefix += `\\cf${cIdx} `;
                suffix = "\\cf0 " + suffix;
              }
            }
            if (m.attrs?.fontSize) {
              const size = parseInt(String(m.attrs.fontSize), 10);
              if (!isNaN(size) && size > 0) {
                prefix += `\\fs${size * 2} `;
                suffix = "\\fs22 " + suffix;
              }
            }
            break;
          case "highlight":
            if (m.attrs?.color) {
              const cIdx = getColorIndex(m.attrs.color);
              if (cIdx != null) {
                prefix += `\\highlight${cIdx} `;
                suffix = "\\highlight0 " + suffix;
              }
            }
            break;
        }
      }
    }

    return `${prefix}${text}${suffix}`;
  };

  const renderBlock = (block: JSONContent): void => {
    if (block.type === "paragraph" || block.type === "heading") {
      let alignCmd = "\\ql";
      const align = block.attrs?.textAlign;
      if (align === "center") alignCmd = "\\qc";
      else if (align === "right") alignCmd = "\\qr";
      else if (align === "justify") alignCmd = "\\qj";

      let headingPrefix = "";
      let headingSuffix = "";
      if (block.type === "heading") {
        const level = Number(block.attrs?.level) || 1;
        const fs = level === 1 ? 32 : level === 2 ? 28 : level === 3 ? 24 : 22;
        headingPrefix = `\\outlinelevel${level - 1}\\b\\fs${fs} `;
        headingSuffix = "\\b0\\fs22";
      }

      const inlines = (block.content ?? []).map(renderInline).join("");
      body += `\\pard\\plain\\f0\\fs22 ${alignCmd} ${headingPrefix}${inlines}${headingSuffix}\\par\n`;
    } else if (block.type === "bulletList" || block.type === "orderedList") {
      const items = block.content ?? [];
      items.forEach((li, idx) => {
        const prefix = block.type === "bulletList" ? "\\bullet\\tab " : `${idx + 1}.\\tab `;
        const pNodes = (li.content ?? []).filter((c) => c.type === "paragraph");
        if (pNodes.length > 0) {
          pNodes.forEach((p) => {
            const inlines = (p.content ?? []).map(renderInline).join("");
            body += `\\pard\\plain\\f0\\fs22\\li720\\fi-360 ${prefix}${inlines}\\par\n`;
          });
        } else {
          const inlines = (li.content ?? []).map(renderInline).join("");
          body += `\\pard\\plain\\f0\\fs22\\li720\\fi-360 ${prefix}${inlines}\\par\n`;
        }
      });
    } else if (block.type === "table") {
      const rows = block.content ?? [];
      for (const row of rows) {
        if (row.type !== "tableRow") continue;
        const cells = row.content ?? [];
        body += "\\trowd\\trgaph108\n";
        let cellRight = 0;
        cells.forEach(() => {
          cellRight += 2000;
          body += `\\cellx${cellRight}\n`;
        });
        for (const cell of cells) {
          const pNodes = cell.content ?? [];
          const text = pNodes
            .map((p) => (p.content ?? []).map(renderInline).join(""))
            .join("\\line ");
          body += `\\intbl ${text}\\cell\n`;
        }
        body += "\\row\n";
      }
    } else if (block.content) {
      for (const child of block.content) {
        renderBlock(child);
      }
    }
  };

  for (const node of nodes) {
    renderBlock(node);
  }

  const fontTable = "{\\fonttbl{\\f0\\fnil\\fcharset0 Calibri;}{\\f1\\fnil\\fcharset0 Arial;}}";
  let colorTable = "";
  if (colorList.length > 0) {
    colorTable =
      "{\\colortbl ;" +
      colorList.map((c) => `\\red${c.r}\\green${c.g}\\blue${c.b};`).join("") +
      "}";
  }

  return `{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\n${fontTable}\n${colorTable}\n${body}}`;
}

/**
 * Parse a standard RTF string into a clean Tiptap JSONContent document.
 */
export function parseRTF(rtf: string): JSONContent {
  const content: JSONContent[] = [];
  let currentInlines: JSONContent[] = [];
  const colorList: ColorRGB[] = [];

  const stateStack: RtfState[] = [];
  let state: RtfState = {
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    subscript: false,
    superscript: false,
  };

  const flushParagraph = (): void => {
    if (currentInlines.length > 0) {
      content.push({
        type: "paragraph",
        ...(state.align ? { attrs: { textAlign: state.align } } : {}),
        content: currentInlines,
      });
      currentInlines = [];
    } else if (content.length === 0 || content[content.length - 1]?.type === "paragraph") {
      content.push({ type: "paragraph", content: [] });
    }
  };

  const appendText = (text: string): void => {
    if (!text) return;
    const marks: NonNullable<JSONContent["marks"]> = [];
    if (state.bold) marks.push({ type: "bold" });
    if (state.italic) marks.push({ type: "italic" });
    if (state.underline) marks.push({ type: "underline" });
    if (state.strike) marks.push({ type: "strike" });
    if (state.subscript) marks.push({ type: "subscript" });
    if (state.superscript) marks.push({ type: "superscript" });

    if (state.colorIndex != null && state.colorIndex > 0 && state.colorIndex <= colorList.length) {
      const c = colorList[state.colorIndex - 1]!;
      const hex = `#${c.r.toString(16).padStart(2, "0")}${c.g.toString(16).padStart(2, "0")}${c.b.toString(16).padStart(2, "0")}`;
      marks.push({ type: "textStyle", attrs: { color: hex } });
    }
    if (state.fontSize != null && state.fontSize > 0) {
      const pt = Math.round(state.fontSize / 2);
      const existingStyle = marks.find((m) => m.type === "textStyle");
      if (existingStyle) {
        existingStyle.attrs = { ...existingStyle.attrs, fontSize: `${pt}pt` };
      } else {
        marks.push({ type: "textStyle", attrs: { fontSize: `${pt}pt` } });
      }
    }

    currentInlines.push({
      type: "text",
      text,
      ...(marks.length > 0 ? { marks } : {}),
    });
  };

  let i = 0;
  const len = rtf.length;
  let skipGroupDepth = 0;

  while (i < len) {
    const ch = rtf[i]!;

    if (ch === "{") {
      stateStack.push({ ...state });
      i++;
      if (rtf[i] === "\\" && rtf[i + 1] === "*") {
        skipGroupDepth++;
      } else if (skipGroupDepth > 0) {
        skipGroupDepth++;
      }
      continue;
    }

    if (ch === "}") {
      if (skipGroupDepth > 0) {
        skipGroupDepth--;
      }
      if (stateStack.length > 0) {
        state = stateStack.pop()!;
      }
      i++;
      continue;
    }

    if (skipGroupDepth > 0) {
      i++;
      continue;
    }

    if (ch === "\\") {
      i++;
      if (i >= len) break;
      const next = rtf[i]!;

      if (next === "\\" || next === "{" || next === "}") {
        appendText(next);
        i++;
        continue;
      }

      if (next === "'") {
        const hex = rtf.slice(i + 1, i + 3);
        const code = parseInt(hex, 16);
        if (!isNaN(code)) {
          appendText(String.fromCharCode(code));
          i += 3;
        } else {
          i++;
        }
        continue;
      }

      let word = "";
      while (i < len && /[a-zA-Z]/.test(rtf[i]!)) {
        word += rtf[i];
        i++;
      }

      let numStr = "";
      let hasNum = false;
      if (i < len && (rtf[i] === "-" || /[0-9]/.test(rtf[i]!))) {
        hasNum = true;
        numStr += rtf[i];
        i++;
        while (i < len && /[0-9]/.test(rtf[i]!)) {
          numStr += rtf[i];
          i++;
        }
      }

      const param = hasNum ? parseInt(numStr, 10) : undefined;

      if (i < len && rtf[i] === " ") {
        i++;
      }

      switch (word) {
        case "b":
          state.bold = param === undefined || param !== 0;
          break;
        case "i":
          state.italic = param === undefined || param !== 0;
          break;
        case "ul":
          state.underline = true;
          break;
        case "ulnone":
          state.underline = false;
          break;
        case "strike":
          state.strike = param === undefined || param !== 0;
          break;
        case "sub":
          state.subscript = true;
          state.superscript = false;
          break;
        case "super":
          state.superscript = true;
          state.subscript = false;
          break;
        case "nosupersub":
          state.subscript = false;
          state.superscript = false;
          break;
        case "fs":
          state.fontSize = param;
          break;
        case "cf":
          state.colorIndex = param;
          break;
        case "highlight":
          state.highlightIndex = param;
          break;
        case "ql":
          state.align = "left";
          break;
        case "qc":
          state.align = "center";
          break;
        case "qr":
          state.align = "right";
          break;
        case "qj":
          state.align = "justify";
          break;
        case "par":
        case "line":
          flushParagraph();
          break;
        case "tab":
          appendText("\t");
          break;
        case "bullet":
          appendText("•");
          break;
        case "u":
          if (param !== undefined) {
            const code = param < 0 ? param + 65536 : param;
            appendText(String.fromCharCode(code));
            if (i < len && rtf[i] !== "\\" && rtf[i] !== "{" && rtf[i] !== "}") {
              i++;
            }
          }
          break;
        case "fonttbl":
        case "stylesheet":
        case "info":
        case "themedata": {
          let depth = 1;
          while (i < len && depth > 0) {
            if (rtf[i] === "{") depth++;
            else if (rtf[i] === "}") depth--;
            i++;
          }
          if (stateStack.length > 0) state = stateStack.pop()!;
          break;
        }
        case "colortbl": {
          let tblStr = "";
          while (i < len && rtf[i] !== "}") {
            tblStr += rtf[i];
            i++;
          }
          const entries = tblStr.split(";");
          for (const entry of entries) {
            const rMatch = /\\red(\d+)/.exec(entry);
            const gMatch = /\\green(\d+)/.exec(entry);
            const bMatch = /\\blue(\d+)/.exec(entry);
            if (rMatch && gMatch && bMatch) {
              colorList.push({
                r: parseInt(rMatch[1]!, 10),
                g: parseInt(gMatch[1]!, 10),
                b: parseInt(bMatch[1]!, 10),
              });
            }
          }
          break;
        }
      }
      continue;
    }

    if (ch === "\r" || ch === "\n") {
      i++;
      continue;
    }

    let plainText = "";
    while (
      i < len &&
      rtf[i] !== "\\" &&
      rtf[i] !== "{" &&
      rtf[i] !== "}" &&
      rtf[i] !== "\r" &&
      rtf[i] !== "\n"
    ) {
      plainText += rtf[i];
      i++;
    }
    appendText(plainText);
  }

  if (currentInlines.length > 0) {
    flushParagraph();
  }

  if (content.length === 0) {
    content.push({ type: "paragraph", content: [] });
  }

  return { type: "doc", content };
}
