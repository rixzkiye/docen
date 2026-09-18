/**
 * LaTeX-like linear math ↔ office-open `MathInput` conversion.
 *
 * The linear editor label (MathInline.attrs.linear / HTML clipboard spans) is a
 * Word-style LaTeX subset; the structured payload is office-open's MathInput,
 * which serializes to OMML. This module is the full parser/serializer pair:
 *
 *  - `latexToMathInput` — recursive-descent parser: fractions, roots
 *    (indexed), super/subscripts (combined into subSuperScript), n-ary
 *    operators (Σ/∏/∫ with limits), function names, delimiters
 *    (`\left…\right`), matrices (`matrix`/`pmatrix`/`bmatrix`/`vmatrix`/
 *    `cases`), aligned environments (`aligned`/`align`/`gather` → eqArr),
 *    accents and bars, Greek letters and operator symbols, and literal text
 *    (`\text`/`\mathrm`/`\operatorname`).
 *  - `mathInputToLatex` — the inverse; unknown-but-parsed OMML nodes
 *    (box/phant/borderBox/groupChr, office-open `{text, properties}` runs)
 *    linearize through their children instead of vanishing.
 *
 * Both directions are canonical: `mathInputToLatex(latexToMathInput(x))` is
 * idempotent (a second parse yields the same MathInput). Constructs the parser
 * does not recognize stay literal text rather than becoming an empty shell —
 * no silent plain-run fallback for anything convertible.
 */

import type { MathInput } from "@office-open/docx";

/** One MathInput item (the parser works per item; sequences are arrays). */
type Item = MathInput;

// ── Symbol tables ──

/** LaTeX control sequence → Unicode symbol. The canonical spellings the
 *  serializer emits (first key wins when several map to the same char). */
const SYMBOLS: Record<string, string> = {
  // Greek lowercase
  alpha: "\u03b1",
  beta: "\u03b2",
  gamma: "\u03b3",
  delta: "\u03b4",
  epsilon: "\u03b5",
  varepsilon: "\u03f5",
  zeta: "\u03b6",
  eta: "\u03b7",
  theta: "\u03b8",
  vartheta: "\u03d1",
  iota: "\u03b9",
  kappa: "\u03ba",
  varkappa: "\u03f0",
  lambda: "\u03bb",
  mu: "\u03bc",
  nu: "\u03bd",
  xi: "\u03be",
  omicron: "\u03bf",
  pi: "\u03c0",
  varpi: "\u03d6",
  rho: "\u03c1",
  varrho: "\u03f1",
  sigma: "\u03c3",
  varsigma: "\u03c2",
  tau: "\u03c4",
  upsilon: "\u03c5",
  phi: "\u03d5",
  varphi: "\u03c6",
  chi: "\u03c7",
  psi: "\u03c8",
  omega: "\u03c9",
  // Greek uppercase
  Gamma: "\u0393",
  Delta: "\u0394",
  Theta: "\u0398",
  Lambda: "\u039b",
  Xi: "\u039e",
  Pi: "\u03a0",
  Sigma: "\u03a3",
  Upsilon: "\u03a5",
  Phi: "\u03a6",
  Psi: "\u03a8",
  Omega: "\u03a9",
  // Binary operators / relations
  times: "\u00d7",
  div: "\u00f7",
  pm: "\u00b1",
  mp: "\u2213",
  cdot: "\u22c5",
  ast: "\u2217",
  star: "\u22c6",
  circ: "\u2218",
  bullet: "\u2219",
  oplus: "\u2295",
  ominus: "\u2296",
  otimes: "\u2297",
  oslash: "\u2298",
  odot: "\u2299",
  leq: "\u2264",
  le: "\u2264",
  geq: "\u2265",
  ge: "\u2265",
  neq: "\u2260",
  ne: "\u2260",
  equiv: "\u2261",
  approx: "\u2248",
  sim: "\u223c",
  simeq: "\u2243",
  cong: "\u2245",
  propto: "\u221d",
  infty: "\u221e",
  partial: "\u2202",
  nabla: "\u2207",
  forall: "\u2200",
  exists: "\u2203",
  nexists: "\u2204",
  emptyset: "\u2205",
  varnothing: "\u2205",
  in: "\u2208",
  notin: "\u2209",
  ni: "\u220b",
  subset: "\u2282",
  supset: "\u2283",
  subseteq: "\u2286",
  supseteq: "\u2287",
  cup: "\u222a",
  cap: "\u2229",
  setminus: "\u2216",
  land: "\u2227",
  wedge: "\u2227",
  lor: "\u2228",
  vee: "\u2228",
  neg: "\u00ac",
  lnot: "\u00ac",
  to: "\u2192",
  rightarrow: "\u2192",
  leftarrow: "\u2190",
  leftrightarrow: "\u2194",
  Rightarrow: "\u21d2",
  Leftarrow: "\u21d0",
  Leftrightarrow: "\u21d4",
  mapsto: "\u21a6",
  longrightarrow: "\u27f6",
  longleftarrow: "\u27f5",
  uparrow: "\u2191",
  downarrow: "\u2193",
  implies: "\u27f9",
  iff: "\u27fa",
  angle: "\u2220",
  perp: "\u22a5",
  parallel: "\u2225",
  triangle: "\u25b3",
  square: "\u25a1",
  therefore: "\u2234",
  because: "\u2235",
  prime: "\u2032",
  ldots: "\u2026",
  cdots: "\u22ef",
  vdots: "\u22ee",
  ddots: "\u22f1",
  degree: "\u00b0",
  aleph: "\u2135",
  hbar: "\u210f",
  ell: "\u2113",
  Re: "\u211c",
  Im: "\u2111",
  wp: "\u2118",
  lfloor: "\u230a",
  rfloor: "\u230b",
  lceil: "\u2308",
  rceil: "\u2309",
  langle: "\u27e8",
  rangle: "\u27e9",
  lbrace: "{",
  rbrace: "}",
  vert: "|",
  Vert: "\u2016",
  backslash: "\u2216",
};

/** Reverse map — Unicode symbol → canonical LaTeX command. First spelling
 *  wins, so aliases (`le`/`ge`/`ne`) canonicalize to `leq`/`geq`/`neq`. */
const SYMBOL_COMMANDS = new Map<string, string>();
for (const [name, char] of Object.entries(SYMBOLS)) {
  if (!SYMBOL_COMMANDS.has(char)) SYMBOL_COMMANDS.set(char, name);
}
SYMBOL_COMMANDS.set("{", "\\{");
SYMBOL_COMMANDS.set("}", "\\}");
SYMBOL_COMMANDS.set("|", "|");
SYMBOL_COMMANDS.set("\u2016", "\\|");

/** N-ary operators: command → struct key, operator char, limit location. */
const NARY: Record<
  string,
  { key: "sum" | "integral"; char: string; limitLocation: "undOvr" | "subSup" }
> = {
  sum: { key: "sum", char: "\u2211", limitLocation: "undOvr" },
  prod: { key: "sum", char: "\u220f", limitLocation: "undOvr" },
  coprod: { key: "sum", char: "\u2210", limitLocation: "undOvr" },
  bigcup: { key: "sum", char: "\u22c3", limitLocation: "undOvr" },
  bigcap: { key: "sum", char: "\u22c2", limitLocation: "undOvr" },
  bigoplus: { key: "sum", char: "\u2a01", limitLocation: "undOvr" },
  bigotimes: { key: "sum", char: "\u2a02", limitLocation: "undOvr" },
  bigodot: { key: "sum", char: "\u2a00", limitLocation: "undOvr" },
  int: { key: "integral", char: "\u222b", limitLocation: "subSup" },
  iint: { key: "integral", char: "\u222c", limitLocation: "subSup" },
  iiint: { key: "integral", char: "\u222d", limitLocation: "subSup" },
  oint: { key: "integral", char: "\u222e", limitLocation: "subSup" },
};

/** Named functions rendered upright (`\sin` → `sin`). `\lim` family fills the
 *  limit structure instead. */
const FUNCTIONS = new Set([
  "sin",
  "cos",
  "tan",
  "cot",
  "sec",
  "csc",
  "arcsin",
  "arccos",
  "arctan",
  "sinh",
  "cosh",
  "tanh",
  "coth",
  "log",
  "ln",
  "lg",
  "exp",
  "max",
  "min",
  "sup",
  "inf",
  "det",
  "dim",
  "ker",
  "hom",
  "arg",
  "deg",
  "gcd",
]);

const LIMITS = new Set(["lim", "limsup", "liminf"]);

/** Accent commands → accent character. */
const ACCENTS: Record<string, string> = {
  hat: "\u0302",
  widehat: "\u0302",
  check: "\u030c",
  breve: "\u0306",
  acute: "\u0301",
  grave: "\u0300",
  dot: "\u0307",
  ddot: "\u0308",
  tilde: "\u0303",
  widetilde: "\u0303",
  vec: "\u20d7",
  bar: "\u0304",
};

/** Accent character → canonical command (serializer). */
const ACCENT_COMMANDS = new Map<string, string>();
for (const [name, char] of Object.entries(ACCENTS)) {
  if (!ACCENT_COMMANDS.has(char)) ACCENT_COMMANDS.set(char, name);
}

/** `\left`/`\right` delimiter token → bracket struct key + rendered pair. */
interface Delimiter {
  key: "roundBrackets" | "squareBrackets" | "curlyBrackets" | "angledBrackets";
  begin: string;
  end: string;
}

const DELIMITERS: Record<string, Delimiter> = {
  "(": { key: "roundBrackets", begin: "(", end: ")" },
  ")": { key: "roundBrackets", begin: "(", end: ")" },
  "[": { key: "squareBrackets", begin: "[", end: "]" },
  "]": { key: "squareBrackets", begin: "[", end: "]" },
  "\\{": { key: "curlyBrackets", begin: "{", end: "}" },
  "\\}": { key: "curlyBrackets", begin: "{", end: "}" },
  "\\lbrace": { key: "curlyBrackets", begin: "{", end: "}" },
  "\\rbrace": { key: "curlyBrackets", begin: "{", end: "}" },
  "\\langle": { key: "angledBrackets", begin: "\u27e8", end: "\u27e9" },
  "\\rangle": { key: "angledBrackets", begin: "\u27e8", end: "\u27e9" },
  "|": { key: "roundBrackets", begin: "|", end: "|" },
  "\\vert": { key: "roundBrackets", begin: "|", end: "|" },
  "\\|": { key: "roundBrackets", begin: "\u2016", end: "\u2016" },
  "\\Vert": { key: "roundBrackets", begin: "\u2016", end: "\u2016" },
  "\\lfloor": { key: "squareBrackets", begin: "\u230a", end: "\u230b" },
  "\\rfloor": { key: "squareBrackets", begin: "\u230a", end: "\u230b" },
  "\\lceil": { key: "squareBrackets", begin: "\u2308", end: "\u2309" },
  "\\rceil": { key: "squareBrackets", begin: "\u2308", end: "\u2309" },
  "\\.": { key: "roundBrackets", begin: "", end: "" },
  ".": { key: "roundBrackets", begin: "", end: "" },
};

const BRACKET_DEFAULTS: Record<Delimiter["key"], [string, string]> = {
  roundBrackets: ["(", ")"],
  squareBrackets: ["[", "]"],
  curlyBrackets: ["{", "}"],
  angledBrackets: ["\u27e8", "\u27e9"],
};

// ── LaTeX → MathInput ──

class LatexParser {
  private i = 0;
  /** A literal-character stop (the `]` of `\sqrt[n]`); null outside. */
  private stopChar: string | null = null;

  constructor(private readonly src: string) {}

  private get atEnd(): boolean {
    return this.i >= this.src.length;
  }

  private peek(offset = 0): string {
    return this.src[this.i + offset] ?? "";
  }

  /** Read a control sequence at the current `\`: the name (letters) or the
   *  single following character. */
  private readCommand(): string {
    if (this.peek() !== "\\") return "";
    this.i += 1;
    if (this.atEnd) return "";
    if (/[a-zA-Z]/.test(this.peek())) {
      let name = "";
      while (!this.atEnd && /[a-zA-Z]/.test(this.peek())) name += this.src[this.i++];
      return name;
    }
    return this.src[this.i++];
  }

  /** Peek the command name at the current `\` without consuming. */
  private peekCommand(): string {
    const save = this.i;
    const name = this.readCommand();
    this.i = save;
    return name;
  }

  /** Read a balanced `{…}` group as raw text (no math parsing inside). */
  private readRawGroup(): string {
    if (this.peek() !== "{") return this.atEnd ? "" : this.src[this.i++];
    let depth = 0;
    let out = "";
    while (!this.atEnd) {
      const ch = this.src[this.i++];
      if (ch === "{") {
        depth += 1;
        if (depth === 1) continue;
      } else if (ch === "}") {
        depth -= 1;
        if (depth === 0) break;
      }
      out += ch;
    }
    return out;
  }

  private skipSpaces(): void {
    while (this.peek() === " ") this.i += 1;
  }

  /** Parse one argument: a `{group}` (all items) or a single atom. */
  private parseArgument(): Item[] {
    this.skipSpaces();
    if (this.peek() === "{") {
      this.i += 1;
      const items = this.parseExpression(true);
      if (this.peek() === "}") this.i += 1;
      return items;
    }
    return this.parseAtom();
  }

  /** Parse an expression (sequence of items). Stops at `}` when `braceStop`,
   *  and always at `&`, `\\`, `\right`, `\end`, or EOF. */
  parseExpression(braceStop = false): Item[] {
    const out: Item[] = [];
    while (!this.atEnd) {
      const ch = this.peek();
      if (braceStop && ch === "}") break;
      if (ch === "&") break;
      if (ch === "\\") {
        if (this.peek(1) === "\\") break;
        const name = this.peekCommand();
        if (name === "right" || name === "end") break;
      }
      const atom = this.parseAtom();
      const applied = this.applyScripts(atom);
      // A n-ary operator or function claims the remainder of the expression as
      // its body (Word's linear format: `\sum_{i}^{n} x` → sum children ["x"]).
      const last = applied[applied.length - 1];
      if (applied.length === 1 && isNaryItem(last)) {
        const body = this.parseExpression(braceStop);
        setNaryChildren(last as Record<string, unknown>, body);
        out.push(last);
        break;
      }
      out.push(...applied);
    }
    return mergeTextItems(out);
  }

  /** Parse one atom: a group, a command, or literal characters. */
  private parseAtom(): Item[] {
    const ch = this.peek();
    if (ch === "{") {
      this.i += 1;
      const items = this.parseExpression(true);
      if (this.peek() === "}") this.i += 1;
      return items;
    }
    if (ch === "\\") return this.parseCommandAtom();
    // Literal characters accumulate until a special char.
    let text = "";
    while (!this.atEnd) {
      const c = this.peek();
      if (c === "\\" || c === "{" || c === "}" || c === "^" || c === "_" || c === "&") break;
      if (this.stopChar !== null && c === this.stopChar) break;
      text += c;
      this.i += 1;
    }
    if (text === "") {
      // A stray `^`/`_`/`&`/`}` with no base: consume so we always advance.
      this.i += 1;
      return [];
    }
    return [text];
  }

  /** Apply postfix `_`/`^` scripts to the atom just parsed. N-ary/limit
   *  structures receive the scripts in their own fields; everything else is
   *  wrapped (subSuperScript when both). */
  private applyScripts(atom: Item[]): Item[] {
    let sub: Item[] | undefined;
    let sup: Item[] | undefined;
    let changed = false;
    for (;;) {
      this.skipSpaces();
      const ch = this.peek();
      if (ch !== "^" && ch !== "_") break;
      this.i += 1;
      const arg = this.parseArgument();
      if (ch === "^") sup = arg;
      else sub = arg;
      changed = true;
    }
    if (!changed) return atom;
    const base: Item[] = atom.length > 0 ? atom : [{ text: "" }];
    const single = base.length === 1 ? (base[0] as Record<string, unknown>) : null;
    if (single && isNaryItem(single)) {
      setNaryChildren(single, undefined, sub, sup);
      return [single as Item];
    }
    if (single && isRecord(single.limitLower)) {
      // The limit structure's own children are the subscript condition; the
      // following expression stays outside (Word's limLow shape).
      const limit = single.limitLower as Record<string, unknown>;
      if (sub) limit.children = sub;
      if (sup) limit.container = sup;
      return [single as Item];
    }
    if (single && isRecord(single.function)) {
      const fn = single.function as Record<string, unknown>;
      const baseFn: Item = { function: { ...fn, children: fn.children ?? [] } } as Item;
      return [buildScript([baseFn], { sub, sup })];
    }
    return [buildScript(base, { sub, sup })];
  }

  private parseCommandAtom(): Item[] {
    const name = this.readCommand();
    if (!name) return [];

    if (name === "frac" || name === "dfrac" || name === "tfrac") {
      const numerator = this.parseArgument();
      const denominator = this.parseArgument();
      return [{ fraction: { numerator, denominator } } as Item];
    }
    if (name === "binom") {
      const numerator = this.parseArgument();
      const denominator = this.parseArgument();
      return [
        {
          roundBrackets: [{ fraction: { numerator, denominator } } as Item],
        } as Item,
      ];
    }
    if (name === "sqrt") {
      this.skipSpaces();
      let degree: Item[] | undefined;
      if (this.peek() === "[") {
        this.i += 1;
        degree = this.parseUntil("]");
        if (this.peek() === "]") this.i += 1;
      }
      const children = this.parseArgument();
      return [{ radical: { children, ...(degree ? { degree } : {}) } } as unknown as Item];
    }
    if (name === "text" || name === "mbox" || name === "textrm") {
      this.skipSpaces();
      // Word renders \text upright — carry the normal (roman) run style so
      // the distinction survives the linear label.
      return [{ text: this.readRawGroup(), properties: { normal: true } } as unknown as Item];
    }
    if (name === "mathrm" || name === "operatorname" || name === "textnormal") {
      this.skipSpaces();
      const raw = this.readRawGroup();
      return [{ text: raw, properties: { normal: true } } as unknown as Item];
    }
    if (name === "left") return this.parseDelimited();
    if (name === "begin") return this.parseEnvironment();
    if (name === "overbrace" || name === "underbrace") {
      const children = this.parseArgument();
      return [{ groupChr: { children, properties: { character: "\u23de" } } } as Item];
    }
    if (name === "overline")
      return [{ bar: { children: this.parseArgument(), type: "top" } } as Item];
    if (name === "underline")
      return [{ bar: { children: this.parseArgument(), type: "bot" } } as Item];
    if (name in ACCENTS) {
      return [
        { accent: { children: this.parseArgument(), accentCharacter: ACCENTS[name]! } } as Item,
      ];
    }
    if (name in NARY) {
      const nary = NARY[name]!;
      return [
        {
          [nary.key]: {
            children: [],
            properties: { limitLocation: nary.limitLocation },
          },
        } as unknown as Item,
      ];
    }
    if (LIMITS.has(name)) {
      return [
        {
          limitLower: { children: [], limit: [{ text: name, properties: { normal: true } }] },
        } as unknown as Item,
      ];
    }
    if (FUNCTIONS.has(name)) {
      return [
        {
          function: {
            name: [{ text: name, properties: { normal: true } } as Item],
            children: this.parseArgument(),
          },
        } as unknown as Item,
      ];
    }
    if (name in SYMBOLS) return [SYMBOLS[name]!];
    if (name === "quad" || name === "qquad") return ["\u2003"];
    if (name === "," || name === ";" || name === ":" || name === " ") return [" "];
    if (name === "!") return [];
    // Unknown command: keep the literal backslash spelling (no data loss).
    return [`\\${name}`];
  }

  /** Parse until the given closing character (not consuming it), respecting
   *  nested groups/commands. */
  private parseUntil(close: string): Item[] {
    const previous = this.stopChar;
    this.stopChar = close;
    const out: Item[] = [];
    while (!this.atEnd && this.peek() !== close) {
      const atom = this.parseAtom();
      out.push(...this.applyScripts(atom));
    }
    this.stopChar = previous;
    return out;
  }

  /** Parse a `\left…\right` delimiter pair. */
  private parseDelimited(): Item[] {
    const leftToken = this.readDelimiterToken();
    this.skipSpaces();
    const children = this.parseExpression(false);
    trimItems(children);
    this.skipSpaces();
    let rightToken = "";
    if (this.peek() === "\\" && this.peekCommand() === "right") {
      this.readCommand();
      rightToken = this.readDelimiterToken();
    }
    const left = DELIMITERS[leftToken] ?? { key: "roundBrackets", begin: "(", end: ")" };
    const right = DELIMITERS[rightToken] ?? left;
    const defaults = BRACKET_DEFAULTS[left.key];
    if (left.begin === defaults[0] && right.end === defaults[1]) {
      return [{ [left.key]: children } as Item];
    }
    return [
      {
        [left.key]: {
          children,
          properties: { beginCharacter: left.begin, endCharacter: right.end },
        },
      } as unknown as Item,
    ];
  }

  /** Read a delimiter token after `\left`/`\right`: a command name, a `\`
   *  escape, or a single character. */
  private readDelimiterToken(): string {
    this.skipSpaces();
    if (this.peek() === "\\") {
      const save = this.i;
      const name = this.readCommand();
      if (`\\${name}` in DELIMITERS) return `\\${name}`;
      this.i = save + 1;
      return "\\";
    }
    if (this.atEnd) return ".";
    return this.src[this.i++];
  }

  /** Parse a `\begin{env}…\end{env}` environment. */
  private parseEnvironment(): Item[] {
    this.skipSpaces();
    const env = this.readRawGroup();
    this.skipSpaces();
    // An array's column spec (`\begin{array}{cc}`) follows the env name.
    if ((env === "array" || env === "matrix*") && this.peek() === "{") this.readRawGroup();

    const rows: Item[][][] = [];
    let row: Item[][] = [];
    let cell: Item[] = [];
    const pushCell = (): void => {
      trimItems(cell);
      row.push(mergeTextItems(cell));
      cell = [];
    };
    const pushRow = (): void => {
      pushCell();
      rows.push(row);
      row = [];
    };

    while (!this.atEnd) {
      if (this.peek() === "\\") {
        if (this.peek(1) === "\\") {
          this.i += 2;
          pushRow();
          continue;
        }
        const name = this.peekCommand();
        if (name === "end") {
          this.readCommand();
          this.skipSpaces();
          this.readRawGroup();
          break;
        }
        if (name === "\\") {
          this.readCommand();
          pushRow();
          continue;
        }
      }
      if (this.peek() === "&") {
        this.i += 1;
        pushCell();
        continue;
      }
      const atom = this.parseAtom();
      cell.push(...this.applyScripts(atom));
    }
    pushRow();

    const matrix: Item = {
      matrix: { rows: rows.map((r) => r.map((c) => (c.length === 1 ? c[0]! : c))) },
    } as Item;
    if (env === "pmatrix") return [{ roundBrackets: [matrix] } as Item];
    if (env === "bmatrix") return [{ squareBrackets: [matrix] } as Item];
    if (env === "Bmatrix") return [{ curlyBrackets: [matrix] } as Item];
    if (env === "vmatrix" || env === "Vmatrix") {
      const line = env === "vmatrix" ? "|" : "\u2016";
      return [
        {
          roundBrackets: {
            children: [matrix],
            properties: { beginCharacter: line, endCharacter: line },
          },
        } as Item,
      ];
    }
    if (env === "cases") {
      return [
        {
          curlyBrackets: [{ eqArr: { rows: rows.map(flattenRow) } } as Item],
        } as Item,
      ];
    }
    if (
      env === "aligned" ||
      env === "align" ||
      env === "align*" ||
      env === "gather" ||
      env === "gathered" ||
      env === "eqnarray" ||
      env === "split"
    ) {
      return [{ eqArr: { rows: rows.map(flattenRow) } } as Item];
    }
    return [matrix];
  }
}

/** True for an item the parser fills with the following expression body —
 *  n-ary operator structs. */
function isNaryItem(item: unknown): boolean {
  return (
    isRecord(item) &&
    (isRecord((item as { sum?: unknown }).sum) ||
      isRecord((item as { integral?: unknown }).integral))
  );
}

/** Trim cell/bracket edge whitespace off a parsed item sequence, so
 *  `a & b` cells canonicalize to ["a"] / ["b"]. */
function trimItems(items: Item[]): void {
  for (const end of ["start", "finish"] as const) {
    for (;;) {
      const item = end === "start" ? items[0] : items[items.length - 1];
      if (typeof item === "string") {
        const trimmed = end === "start" ? item.replace(/^\s+/, "") : item.replace(/\s+$/, "");
        if (trimmed === item) break;
        if (end === "start") items[0] = trimmed;
        else items[items.length - 1] = trimmed;
        if (trimmed !== "") break;
        if (end === "start") items.shift();
        else items.pop();
        if (items.length === 0) break;
        continue;
      }
      const rec = item as { text?: unknown };
      if (isRecord(item) && typeof rec.text === "string") {
        const trimmed =
          end === "start" ? rec.text.replace(/^\s+/, "") : rec.text.replace(/\s+$/, "");
        if (trimmed === rec.text) break;
        rec.text = trimmed;
        if (trimmed !== "") break;
        if (end === "start") items.shift();
        else items.pop();
        if (items.length === 0) break;
        continue;
      }
      break;
    }
  }
}

/** Set a n-ary struct's children/scripts in place. */
function setNaryChildren(
  item: Record<string, unknown>,
  body?: Item[],
  sub?: Item[],
  sup?: Item[],
): void {
  const value = (item.sum ?? item.integral) as Record<string, unknown>;
  if (body !== undefined) value.children = body;
  if (sub !== undefined) value.subScript = sub;
  if (sup !== undefined) value.superScript = sup;
}

/** The literal text an item contributes, when it is a text run. */
function textOf(item: unknown): string | undefined {
  if (typeof item === "string") return item;
  const rec = item as { text?: unknown } | undefined;
  return rec && typeof rec.text === "string" ? rec.text : undefined;
}

/** Merge adjacent plain-string items into one run. MathInput run granularity
 *  is not semantic, and merging makes the canonical linear form a fixed point
 *  (cell boundaries an aligned environment lost in flattening re-merge). */
function mergeTextItems(items: Item[]): Item[] {
  const out: Item[] = [];
  for (const item of items) {
    const prev = out[out.length - 1];
    if (typeof item === "string" && typeof prev === "string") {
      out[out.length - 1] = prev + item;
    } else {
      out.push(item);
    }
  }
  return out;
}

/** Flatten environment cells into one eqArr row, keeping a space between two
 *  cells so distinct columns never word-join (`x` & `x > 0` → `x x > 0`). */
function flattenRow(cells: Item[][]): Item[] {
  const out: Item[] = [];
  for (const cell of cells) {
    if (cell.length === 0) continue;
    const last = out[out.length - 1];
    const first = cell[0];
    const endsSpace = typeof last === "string" && /\s$/.test(last);
    const firstText = textOf(first);
    const startsSpace = firstText !== undefined && /^\s/.test(firstText);
    if (out.length > 0 && !endsSpace && !startsSpace) out.push(" ");
    out.push(...cell);
  }
  return mergeTextItems(out);
}

/** Combine a base + scripts into the right MathInput struct (subSuperScript
 *  when both are present). */
function buildScript(base: Item[], scripts: { sub?: Item[]; sup?: Item[] }): Item {
  if (scripts.sub && scripts.sup) {
    return {
      subSuperScript: { children: base, subScript: scripts.sub, superScript: scripts.sup },
    } as Item;
  }
  if (scripts.sup) {
    return { superScript: { children: base, superScript: scripts.sup } } as Item;
  }
  return { subScript: { children: base, subScript: scripts.sub ?? [] } } as Item;
}

/**
 * Parse a LaTeX-like linear expression into a MathInput sequence. The result
 * is the sequence (not the office-open paragraph wrapper): callers place it in
 * `math.children`.
 */
export function latexToMathInput(linear: string): Item[] {
  return new LatexParser(linear).parseExpression(false);
}

/** Backward-compatible single-value shape: one item when the input is a single
 *  struct, else `{ children: […] }` (the office-open math paragraph shape). */
export function convertLinearToOMML(linear: string): Record<string, unknown> {
  const items = latexToMathInput(linear);
  if (items.length === 1 && typeof items[0] !== "string") {
    return items[0] as unknown as Record<string, unknown>;
  }
  return { children: items };
}

// ── MathInput → LaTeX ──

export function mathInputToLatex(math: unknown): string {
  return linearizeItem(math);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function linearizeItems(items: readonly unknown[] | undefined): string {
  if (!Array.isArray(items)) return "";
  let out = "";
  for (const item of items) {
    const part = linearizeItem(item);
    // A control sequence ending in letters followed by a letter would merge
    // into one unknown command (`\leq` + `b` → `\leqb`); keep a separator.
    if (/[a-zA-Z]$/.test(out) && /^[a-zA-Z]/.test(part)) out += " ";
    out += part;
  }
  return out;
}

/** Serialize one item, escaping/quoting symbols so the output re-parses to an
 *  equivalent MathInput. */
function linearizeItem(item: unknown): string {
  if (typeof item === "string") return quoteText(item);
  if (!isRecord(item)) return "";

  // A math run: bare `{ text }` or office-open's `{ text, properties }`.
  if (typeof item.text === "string") {
    return item.properties && isRecord(item.properties) && item.properties.normal === true
      ? `\\mathrm{${escapeRaw(item.text)}}`
      : quoteText(item.text);
  }

  if (isRecord(item.fraction)) {
    const f = item.fraction;
    return `\\frac{${linearizeItems(f.numerator as unknown[])}}{${linearizeItems(f.denominator as unknown[])}}`;
  }
  if (isRecord(item.radical)) {
    const r = item.radical;
    const inner = linearizeItems(r.children as unknown[]);
    const degree = Array.isArray(r.degree) ? linearizeItems(r.degree) : "";
    return degree ? `\\sqrt[${degree}]{${inner}}` : `\\sqrt{${inner}}`;
  }
  if (isRecord(item.superScript)) {
    const s = item.superScript;
    return `${linearizeItems(s.children as unknown[])}^{${linearizeItems(s.superScript as unknown[])}}`;
  }
  if (isRecord(item.subScript)) {
    const s = item.subScript;
    return `${linearizeItems(s.children as unknown[])}_{${linearizeItems(s.subScript as unknown[])}}`;
  }
  if (isRecord(item.subSuperScript)) {
    const s = item.subSuperScript;
    return `${linearizeItems(s.children as unknown[])}_{${linearizeItems(s.subScript as unknown[])}}^{${linearizeItems(s.superScript as unknown[])}}`;
  }
  if (isRecord(item.preSubSuperScript)) {
    const s = item.preSubSuperScript;
    return `{}_{${linearizeItems(s.subScript as unknown[])}}^{${linearizeItems(s.superScript as unknown[])}}${linearizeItems(s.children as unknown[])}`;
  }
  if (isRecord(item.sum)) return linearizeNary("\\sum", item.sum);
  if (isRecord(item.integral)) return linearizeNary("\\int", item.integral);
  if (isRecord(item.limitLower)) return linearizeLimit("lower", item.limitLower);
  if (isRecord(item.limitUpper)) return linearizeLimit("upper", item.limitUpper);
  if (isRecord(item.function)) {
    const f = item.function;
    const name = rawText(f.name as unknown[]);
    const body = linearizeItems(f.children as unknown[]);
    if (FUNCTIONS.has(name)) {
      return `\\${name}${/^[a-zA-Z]$/.test(body.slice(0, 1)) ? " " : ""}${body}`;
    }
    return `\\operatorname{${escapeRaw(name)}}${body}`;
  }
  if (isRecord(item.matrix)) {
    const rows = (item.matrix.rows ?? []) as unknown[][];
    const body = rows
      .map((row) =>
        row.map((cell) => linearizeItems(Array.isArray(cell) ? cell : [cell])).join(" & "),
      )
      .join(" \\\\ ");
    return `\\begin{matrix}${body}\\end{matrix}`;
  }
  if (isRecord(item.eqArr)) {
    const rows = (item.eqArr.rows ?? []) as unknown[][];
    const body = rows.map((row) => linearizeItems(row)).join(" \\\\ ");
    return `\\begin{aligned}${body}\\end{aligned}`;
  }
  for (const key of [
    "roundBrackets",
    "squareBrackets",
    "curlyBrackets",
    "angledBrackets",
  ] as const) {
    if (key in item) {
      const { children, properties } = bracketParts(item[key]);
      const defaults = BRACKET_DEFAULTS[key];
      const custom = isRecord(properties)
        ? {
            begin:
              typeof properties.beginCharacter === "string"
                ? properties.beginCharacter
                : defaults[0],
            end:
              typeof properties.endCharacter === "string" ? properties.endCharacter : defaults[1],
          }
        : { begin: defaults[0], end: defaults[1] };
      return `\\left${delimiterCommand(custom.begin)}${linearizeItems(children)}\\right${delimiterCommand(custom.end)}`;
    }
  }
  if (isRecord(item.accent)) {
    const a = item.accent;
    const char = typeof a.accentCharacter === "string" ? a.accentCharacter : "\u0302";
    const command = ACCENT_COMMANDS.get(char) ?? "hat";
    return `\\${command}{${linearizeItems(a.children as unknown[])}}`;
  }
  if (isRecord(item.bar)) {
    const b = item.bar;
    const command = b.type === "bot" ? "underline" : "overline";
    return `\\${command}{${linearizeItems(b.children as unknown[])}}`;
  }
  if (isRecord(item.groupChr)) return linearizeItems(item.groupChr.children as unknown[]);
  if (isRecord(item.borderBox)) return linearizeItems(item.borderBox.children as unknown[]);
  if (isRecord(item.box)) return linearizeItems(item.box.children as unknown[]);
  if (isRecord(item.phant)) return linearizeItems(item.phant.children as unknown[]);
  if (isRecord(item.argumentControlProperties)) return "";
  if (Array.isArray(item.children)) return linearizeItems(item.children);
  return "";
}

/** Children + properties of a bracket item (array sugar or full shape). */
function bracketParts(value: unknown): { children: unknown[]; properties?: unknown } {
  if (Array.isArray(value)) return { children: value };
  if (isRecord(value)) {
    if (Array.isArray(value.elements)) {
      const flat: unknown[] = [];
      for (const element of value.elements) {
        if (Array.isArray(element)) flat.push(...element);
      }
      return { children: flat, properties: value.properties };
    }
    return {
      children: Array.isArray(value.children) ? value.children : [],
      properties: value.properties,
    };
  }
  return { children: [] };
}

function linearizeNary(op: string, value: unknown): string {
  const n = isRecord(value) ? value : {};
  const sub = Array.isArray(n.subScript) ? linearizeItems(n.subScript) : "";
  const sup = Array.isArray(n.superScript) ? linearizeItems(n.superScript) : "";
  const body = linearizeItems(n.children as unknown[]);
  const limits = `${sub ? `_{${sub}}` : ""}${sup ? `^{${sup}}` : ""}`;
  return `${op}${limits}${body}`;
}

/** The raw text of a MathInput sequence (function/limit names). */
function rawText(items: unknown): string {
  if (!Array.isArray(items)) return "";
  let out = "";
  for (const item of items) {
    if (typeof item === "string") out += item;
    else if (isRecord(item) && typeof item.text === "string") out += item.text;
  }
  return out;
}

/** `\lim`/`\limsup`/`\liminf` (and any limitUpper/Lower) back to linear. */
function linearizeLimit(kind: "lower" | "upper", value: unknown): string {
  const l = isRecord(value) ? value : {};
  const name = rawText(l.limit);
  const command = LIMITS.has(name) ? `\\${name}` : `\\operatorname{${escapeRaw(name || "lim")}}`;
  const inner = linearizeItems(
    (kind === "lower" ? l.children : (l.container ?? l.children)) as unknown[],
  );
  return kind === "lower" ? `${command}_{${inner}}` : `${command}^{${inner}}`;
}

/** Delimiter char → `\left` token spelling. */
function delimiterCommand(delimiter: string): string {
  switch (delimiter) {
    case "(":
    case ")":
    case "[":
    case "]":
      return delimiter;
    case "{":
      return "\\{";
    case "}":
      return "\\}";
    case "\u230a":
      return "\\lfloor";
    case "\u230b":
      return "\\rfloor";
    case "\u2308":
      return "\\lceil";
    case "\u2309":
      return "\\rceil";
    case "\u27e8":
      return "\\langle";
    case "\u27e9":
      return "\\rangle";
    case "\u2016":
      return "\\|";
    case "":
      return ".";
    default:
      return delimiter;
  }
}

/** Serialize a text run, converting symbol chars to their commands. A letter
 *  following a symbol command gets a separating space (else it re-parses as a
 *  longer unknown command). */
function quoteText(text: string): string {
  let out = "";
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    const command = SYMBOL_COMMANDS.get(ch);
    if (command) {
      out += /^[a-zA-Z]+$/.test(command) ? `\\${command}` : command;
      if (/^[a-zA-Z]$/.test(text[i + 1] ?? "")) out += " ";
    } else {
      out += ch;
    }
  }
  return out;
}

/** Escape literal text inside `\mathrm{…}` (only braces/backslash need care). */
function escapeRaw(text: string): string {
  return text.replace(/[{}]/g, (ch) => `\\${ch}`);
}
