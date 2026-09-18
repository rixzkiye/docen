import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { unzipSync } from "@office-open/core";
import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCXSync, type JSONContent } from "../index";
import { convertLinearToOMML, convertOMMLToLinear } from "./math";

/**
 * Item-5 fixture matrix. Each case asserts three independent things:
 *
 *  1. the linear input parses into the expected MathInput structure
 *     (fractions/roots/scripts/Greek/operators/n-ary/matrices/aligned);
 *  2. linear → OMML → linear is canonical and idempotent — the second cycle
 *     is identical to the first (no plain-run fallback for convertible
 *     constructs, no structure drift);
 *  3. the OMML survives DOCX generation and re-parsing, and LibreOffice
 *     renders the formula (PDF text oracle) when soffice is available.
 */

/** Run one linear expression through the canonical round-trip. */
function cycle(linear: string): { omml: Record<string, unknown>; linear: string } {
  const omml = convertLinearToOMML(linear);
  const once = convertOMMLToLinear(omml);
  return { omml, linear: once };
}

/** Assert canonical idempotence: parse→linearize twice is a fixed point. */
function expectCanonical(linear: string): { omml: Record<string, unknown>; linear: string } {
  const first = cycle(linear);
  const second = cycle(first.linear);
  expect(second.linear).toBe(first.linear);
  expect(second.omml).toEqual(first.omml);
  return first;
}

describe("linear → OMML → linear canonical round-trip", () => {
  const cases = [
    // fractions
    "\\frac{a}{b}",
    "\\frac{x+1}{y-1}",
    // roots
    "\\sqrt{x}",
    "\\sqrt[3]{8}",
    // scripts
    "x^{2}",
    "a_{i}",
    "x_{1}^{2}",
    // Greek + operators
    "\\alpha + \\beta = \\gamma",
    "a \\leq b \\neq c \\times d \\pm \\infty",
    // n-ary
    "\\sum_{i=1}^{n} x_{i}",
    "\\prod_{k=0}^{m} a_{k}",
    "\\int_{0}^{1} f(x) dx",
    // functions + limits
    "\\sin x",
    "\\lim_{x \\to 0} \\frac{\\sin x}{x}",
    // delimiters
    "\\left( \\frac{a}{b} \\right)",
    "\\left[ x^{2} \\right]",
    "\\left| y \\right|",
    // matrices / aligned
    "\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}",
    "\\begin{matrix} 1 & 2 \\\\ 3 & 4 \\end{matrix}",
    "\\begin{aligned} x &= 1 \\\\ y &= 2 \\end{aligned}",
    "\\begin{cases} x & x > 0 \\\\ -x & x \\leq 0 \\end{cases}",
    // accents / bars
    "\\hat{x}",
    "\\bar{x}",
    "\\vec{v}",
    "\\overline{AB}",
    // literal text
    "\\text{if } x > 0",
    "\\mathrm{d}x",
    // binom
    "\\binom{n}{k}",
  ];

  for (const linear of cases) {
    it(`round-trips ${linear}`, () => {
      expectCanonical(linear);
    });
  }
});

describe("structured MathInput shapes", () => {
  it("parses Greek letters and operators as symbol runs", () => {
    expect(convertLinearToOMML("\\alpha")).toEqual({ children: ["\u03b1"] });
    expect(convertLinearToOMML("\\Omega")).toEqual({ children: ["\u03a9"] });
    // Adjacent text runs merge into one canonical run (run granularity is not
    // semantic in MathInput).
    expect(convertLinearToOMML("a \\leq b")).toEqual({ children: ["a \u2264b"] });
    expect(convertLinearToOMML("a \\times b \\pm c")).toEqual({
      children: ["a \u00d7b \u00b1c"],
    });
  });

  it("parses n-ary limits into the sum structure", () => {
    const omml = convertLinearToOMML("\\sum_{i=1}^{n} x_{i}") as {
      sum?: { children: unknown[]; subScript: unknown[]; superScript: unknown[] };
    };
    expect(omml.sum?.subScript).toEqual(["i=1"]);
    expect(omml.sum?.superScript).toEqual(["n"]);
    expect(omml.sum?.children).toEqual([{ subScript: { children: ["x"], subScript: ["i"] } }]);
    expect(convertOMMLToLinear(omml)).toBe("\\sum_{i=1}^{n}x_{i}");
  });

  it("parses matrices and bracket-wrapped matrices", () => {
    const omml = convertLinearToOMML("\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}") as {
      roundBrackets?: Array<{ matrix?: { rows: unknown[][] } }>;
    };
    const rows = omml.roundBrackets?.[0]?.matrix?.rows;
    expect(rows).toHaveLength(2);
    expect(rows?.[0]).toEqual(["a", "b"]);
    expect(convertOMMLToLinear(omml)).toBe(
      "\\left(\\begin{matrix}a & b \\\\ c & d\\end{matrix}\\right)",
    );
  });

  it("parses aligned environments into eqArr", () => {
    const omml = convertLinearToOMML("\\begin{aligned} x &= 1 \\\\ y &= 2 \\end{aligned}") as {
      eqArr?: { rows: unknown[][] };
    };
    expect(omml.eqArr?.rows).toHaveLength(2);
  });

  it("keeps unknown commands literal instead of an empty shell", () => {
    expect(convertLinearToOMML("\\unknowncmd")).toEqual({ children: ["\\unknowncmd"] });
    expect(convertOMMLToLinear(convertLinearToOMML("\\unknowncmd"))).toBe("\\unknowncmd");
  });

  it("linearizes parsed office-open run properties without losing text", () => {
    const omml = {
      children: [
        {
          fraction: {
            numerator: [{ text: "a", properties: { style: "italic" } }],
            denominator: ["b"],
          },
        },
        { text: " + " },
        { text: "sin", properties: { normal: true } },
      ],
    };
    expect(convertOMMLToLinear(omml)).toBe("\\frac{a}{b} + \\mathrm{sin}");
  });
});

// ── DOCX / LibreOffice oracle ──

const doctFor = (specs: string[]): JSONContent => ({
  type: "doc",
  content: specs.map((linear, i) => ({
    type: "paragraph",
    content: [
      { type: "text", text: `eq${i}: ` },
      {
        type: "mathInline",
        attrs: { math: convertLinearToOMML(linear), linear },
      },
    ],
  })),
});

describe("math fixture matrix through DOCX", { timeout: 20000 }, () => {
  const specs: Array<[linear: string, expectedKey: RegExp]> = [
    ["\\frac{a}{b}", /fraction/],
    ["\\sqrt[3]{8}", /radical/],
    ["x_{1}^{2}", /subSuperScript/],
    ["\\alpha + \\beta", /children/],
    ["\\sum_{i=1}^{n} x_{i}", /sum/],
    ["\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}", /roundBrackets/],
  ];

  it("keeps every formula as m:oMath through generate → parse", () => {
    const bytes = generateDOCXSync(doctFor(specs.map(([linear]) => linear)), {
      prepare: false,
    }) as Uint8Array;
    const reparsed = parseDOCXSync(bytes);
    const maths = (reparsed.content ?? []).flatMap((p) =>
      (p.content ?? []).filter((c) => c.type === "mathInline"),
    );
    expect(maths).toHaveLength(specs.length);
    // Every reparsed formula keeps its structure (not an empty shell).
    maths.forEach((math, index) => {
      expect(JSON.stringify(math.attrs?.math)).toMatch(specs[index]![1]);
    });
  });

  it("LibreOffice renders and re-emits the formulas (independent oracle)", () => {
    if (!existsSync("/usr/bin/soffice")) {
      console.log("⏭ soffice missing — LibreOffice oracle skipped");
      return;
    }
    const dir = mkdtempSync(join(tmpdir(), "docen-math-"));
    try {
      const docx = generateDOCXSync(doctFor(specs.map(([linear]) => linear)), {
        prepare: false,
      }) as Uint8Array;
      const docxPath = join(dir, "math.docx");
      writeFileSync(docxPath, Buffer.from(docx));
      const loArgs = [
        "--headless",
        `-env:UserInstallation=file://${join(dir, "lo-profile")}`,
        "--convert-to",
      ] as const;
      const childEnv = { ...process.env, TMPDIR: dir };
      execFileSync("/usr/bin/soffice", [...loArgs, "pdf", "--outdir", dir, docxPath], {
        stdio: "ignore",
        timeout: 120000,
        env: childEnv,
      });
      const text = execFileSync("pdftotext", [join(dir, "math.pdf"), "-"], {
        encoding: "utf8",
      });
      // The non-formula document text renders (proves the file opened).
      expect(text).toContain("eq0:");
      // Symbol runs inside m:oMath land in the text layer.
      expect(text).toContain("\u03b1 + \u03b2");
      // LibreOffice re-serializes the DOCX: every m:oMath must survive its
      // independent OMML parse/write cycle.
      execFileSync(
        "/usr/bin/soffice",
        [...loArgs, "docx", "--outdir", join(dir, "resave"), docxPath],
        { stdio: "ignore", timeout: 120000, env: childEnv },
      );
      const resaved = readFileSync(join(dir, "resave", "math.docx"));
      const xml = new TextDecoder().decode(unzipSync(new Uint8Array(resaved))["word/document.xml"]);
      const count = [...xml.matchAll(/<m:oMath[ >]/g)].length;
      expect(count).toBe(specs.length);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
