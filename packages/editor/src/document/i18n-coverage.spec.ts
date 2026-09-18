// @vitest-environment happy-dom
/**
 * Locale-table compatibility guard: every literal `t("key")` call must resolve
 * from a registered translation table (no raw keys in the UI), and the en and
 * zh-CN business tables must stay key-symmetric.
 */
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join, resolve } from "node:path";

import { describe, expect, it } from "vitest";

import { ribbonEn, ribbonZhCN } from "./i18n";

const SRC_ROOT = resolve(process.cwd(), "packages/editor/src");

function walk(dir: string, out: string[] = []): string[] {
  for (const entry of readdirSync(dir)) {
    const path = join(dir, entry);
    if (statSync(path).isDirectory()) walk(path, out);
    else if (/\.tsx?$/.test(path) && !path.endsWith(".spec.ts")) out.push(path);
  }
  return out;
}

function codeLines(file: string): string {
  return readFileSync(file, "utf8")
    .split("\n")
    .filter((line) => {
      const trimmed = line.trim();
      return !trimmed.startsWith("//") && !trimmed.startsWith("*");
    })
    .join("\n");
}

/** Every `"key.with.dots":` property name across the editor sources — this
 *  covers the business tables (i18n.ts), component-local tables, and the
 *  localize module seeds in one pass. */
function definedKeys(files: string[]): Set<string> {
  const keys = new Set<string>();
  for (const file of files) {
    for (const match of codeLines(file).matchAll(/"([A-Za-z][\w.-]*\.[\w.-]+)"\s*:/g)) {
      keys.add(match[1]);
    }
  }
  return keys;
}

/** Literal `t("key")` / `t("key", el)` usages; template literals and
 *  concatenations are dynamic and excluded. */
function usedLiteralKeys(files: string[]): Map<string, string> {
  const used = new Map<string, string>();
  for (const file of files) {
    for (const match of codeLines(file).matchAll(/\bt\(\s*"([^"]+)"\s*(?:,[^)]*)?\)/g)) {
      used.set(match[1], file);
    }
  }
  return used;
}

describe("editor i18n coverage", () => {
  const files = walk(SRC_ROOT);
  const defined = definedKeys(files);

  it("resolves every literal t() key from a translation table", () => {
    const missing = [...usedLiteralKeys(files)]
      .filter(([key]) => !defined.has(key))
      .map(([key, file]) => `${key} (${file.slice(SRC_ROOT.length + 1)})`);
    expect(missing).toEqual([]);
  });

  it("keeps en and zh-CN business tables key-symmetric", () => {
    const en = Object.keys(ribbonEn.translations).sort();
    const zh = Object.keys(ribbonZhCN.translations).sort();
    expect(en.filter((k) => !zh.includes(k))).toEqual([]);
    expect(zh.filter((k) => !en.includes(k))).toEqual([]);
  });

  it("has no empty translations", () => {
    const empty = Object.entries(ribbonEn.translations)
      .filter(([, value]) => value.trim() === "")
      .map(([key]) => key);
    expect(empty).toEqual([]);
  });
});
