import { Document, Paragraph } from "@docen/docx";
import { Editor, Node as TextNode, type Editor as EditorType } from "@docen/docx/core";
import { beforeEach, describe, expect, it } from "vitest";

import { CustomDictionaryManager } from "./custom-dictionary";
import { checkGrammar } from "./grammar";
import { ProofingPackManager } from "./pack-manager";
import { getSynonyms, lookupThesaurus } from "./thesaurus";

const Text = TextNode.create({ name: "text", group: "inline" });

const build = (...paragraphs: string[]): EditorType =>
  new Editor({
    element: null,
    extensions: [Document, Paragraph, Text],
    content: {
      type: "doc",
      content: paragraphs.map((text) => ({
        type: "paragraph",
        content: [{ type: "text", text }],
      })),
    },
  });

describe("CustomDictionaryManager", () => {
  it("adds and checks words case-insensitively", () => {
    const dict = new CustomDictionaryManager("test-dict-1");
    dict.clear();
    expect(dict.has("CustomWord")).toBe(false);
    dict.add("CustomWord");
    expect(dict.has("customword")).toBe(true);
    expect(dict.has("CUSTOMWORD")).toBe(true);
    expect(dict.size()).toBe(1);
  });

  it("removes words", () => {
    const dict = new CustomDictionaryManager("test-dict-2");
    dict.clear();
    dict.add("sample");
    expect(dict.has("sample")).toBe(true);
    dict.remove("SAMPLE");
    expect(dict.has("sample")).toBe(false);
  });

  it("returns sorted all words", () => {
    const dict = new CustomDictionaryManager("test-dict-3");
    dict.clear();
    dict.add("zebra");
    dict.add("apple");
    dict.add("mango");
    expect(dict.all()).toEqual(["apple", "mango", "zebra"]);
  });
});

describe("checkGrammar", () => {
  it("detects repeated consecutive words", () => {
    const editor = build("This is the the best day.");
    const issues = checkGrammar(editor.state.doc);
    expect(issues.some((i) => i.ruleId === "repeated-words")).toBe(true);
    const issue = issues.find((i) => i.ruleId === "repeated-words")!;
    expect(issue.replacements).toEqual(["the"]);
  });

  it("detects unnecessary space before punctuation", () => {
    const editor = build("Hello world , how are you?");
    const issues = checkGrammar(editor.state.doc);
    expect(issues.some((i) => i.ruleId === "space-before-punctuation")).toBe(true);
  });

  it("detects multiple consecutive spaces", () => {
    const editor = build("Multiple   spaces here.");
    const issues = checkGrammar(editor.state.doc);
    expect(issues.some((i) => i.ruleId === "multiple-spaces")).toBe(true);
  });

  it("detects uncapitalized sentence beginnings", () => {
    const editor = build("First sentence. second sentence.");
    const issues = checkGrammar(editor.state.doc);
    expect(issues.some((i) => i.ruleId === "sentence-capitalization")).toBe(true);
    const issue = issues.find((i) => i.ruleId === "sentence-capitalization")!;
    expect(issue.replacements).toEqual(["Second"]);
  });

  it("detects improper indefinite articles", () => {
    const editor = build("He ate a apple and an banana.");
    const issues = checkGrammar(editor.state.doc);
    expect(issues.some((i) => i.ruleId === "indefinite-articles")).toBe(true);
  });

  it("detects frequently confused homophones", () => {
    const editor = build("They forgot they're homework.");
    const issues = checkGrammar(editor.state.doc);
    expect(issues.some((i) => i.ruleId === "frequently-confused")).toBe(true);
    const issue = issues.find((i) => i.ruleId === "frequently-confused")!;
    expect(issue.replacements).toEqual(["their"]);
  });

  it("respects options to disable grammar check", () => {
    const editor = build("This is the the best day.");
    expect(checkGrammar(editor.state.doc, { checkGrammar: false })).toEqual([]);
    expect(checkGrammar(editor.state.doc, { hideGrammarErrors: true })).toEqual([]);
  });
});

describe("Thesaurus", () => {
  it("looks up English words and retrieves synonyms and antonyms", () => {
    const res = lookupThesaurus("good", "en");
    expect(res).not.toBeNull();
    expect(res?.word).toBe("good");
    expect(res?.meanings[0].synonyms).toContain("great");
    expect(res?.meanings[0].antonyms).toContain("bad");
  });

  it("looks up Indonesian words", () => {
    const res = lookupThesaurus("baik", "id");
    expect(res).not.toBeNull();
    expect(res?.word).toBe("baik");
    expect(res?.meanings[0].synonyms).toContain("bagus");
  });

  it("performs reverse lookup when searching a synonym", () => {
    const res = lookupThesaurus("great", "en");
    expect(res).not.toBeNull();
    expect(res?.meanings[0].synonyms).toContain("good");
  });

  it("retrieves flat synonyms list via getSynonyms", () => {
    const syns = getSynonyms("good", "en");
    expect(syns.length).toBeGreaterThan(0);
    expect(syns).not.toContain("good");
  });
});

describe("ProofingPackManager", () => {
  let mgr: ProofingPackManager;

  beforeEach(() => {
    mgr = new ProofingPackManager();
  });

  it("has preinstalled packs for English and Indonesian", () => {
    expect(mgr.isInstalled("en")).toBe(true);
    expect(mgr.isInstalled("id")).toBe(true);
    expect(mgr.getPack("en")?.name).toBe("English");
    expect(mgr.getPack("id")?.name).toBe("Bahasa Indonesia");
  });

  it("resolves regional language tags by prefix", () => {
    expect(mgr.getPack("en-US")?.id).toBe("en");
    expect(mgr.getPack("id-ID")?.id).toBe("id");
  });

  it("prevents uninstalling default English pack", () => {
    expect(mgr.uninstallPack("en")).toBe(false);
    expect(mgr.isInstalled("en")).toBe(true);
  });

  it("installs and uninstalls custom pack", () => {
    mgr.installPack({
      id: "fr",
      name: "Français",
      spellWords: new Set(["bonjour", "monde"]),
    });
    expect(mgr.isInstalled("fr")).toBe(true);
    expect(mgr.getPack("fr")?.name).toBe("Français");
    expect(mgr.uninstallPack("fr")).toBe(true);
    expect(mgr.isInstalled("fr")).toBe(false);
  });
});
