import type { Editor } from "@docen/docx/core";
import type { Transaction } from "@tiptap/pm/state";

import {
  addSpellWord,
  checkGrammar,
  checkSpelling,
  ignoreSpellOnce,
  ignoreSpellWord,
  spellSuggestions,
  type GrammarIssue,
  type SpellingIssue,
} from "../spelling";

/** The spelling commands' view of the host — resolved per call so the
 *  controller can be built before a document opens. */
export interface SpellingHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — squiggle overlay feeds and jump scrolling. */
  bridge():
    | {
        setSpellingIssues(issues: SpellingIssue[]): void;
        setGrammarIssues?(issues: GrammarIssue[]): void;
        scrollIntoView(pos: number): void;
      }
    | undefined;
  /** The host element — the shadow-DOM root for the proofing pane and the
   *  status-bar book. */
  element(): HTMLElement;
}

interface UnifiedProofingEntry {
  kind: "spelling" | "grammar";
  from: number;
  to: number;
  word: string;
  suggestions: string[];
  category: "spelling" | "grammar" | "style";
  message?: string;
  ruleId?: string;
}

/**
 * The spelling and grammar domain, split out of the host element: the debounced
 * check after every render, the issue list that feeds the squiggle overlays, the
 * proofing pane, and the status-bar book, plus the replace/ignore/navigation
 * interactions.
 */
export class SpellingCommands {
  constructor(private readonly host: SpellingHost) {}

  #issues: SpellingIssue[] = [];
  #grammarIssues: GrammarIssue[] = [];
  #unifiedEntries: UnifiedProofingEntry[] = [];
  #ignoredGrammar = new Set<string>();

  /** The pane's active issue (document order); -1 = nothing selected. */
  #active = -1;
  #timer: ReturnType<typeof setTimeout> | null = null;
  /** Word's "Check spelling as you type" (File → Options → Proofing): off
   *  stops the checks and clears the squiggles. */
  #enabled = true;

  /** Tear down pending timers (the host's disconnectedCallback). */
  dispose(): void {
    if (this.#timer != null) clearTimeout(this.#timer);
  }

  /** Re-check after the user pauses (debounced; driven by every render). */
  schedule(): void {
    if (!this.#enabled) return;
    if (this.#timer != null) clearTimeout(this.#timer);
    this.#timer = setTimeout(() => {
      this.#timer = null;
      this.run();
    }, 400);
  }

  /** The checker's on/off (the options dialog's Proofing section). Turning
   *  it off clears every visible squiggle immediately; turning it back on
   *  re-checks at once. */
  setEnabled(on: boolean): void {
    if (on === this.#enabled) return;
    this.#enabled = on;
    this.#active = -1;
    if (on) this.run();
    else {
      if (this.#timer != null) clearTimeout(this.#timer);
      this.#timer = null;
      this.#issues = [];
      this.#grammarIssues = [];
      this.#unifiedEntries = [];
      this.host.bridge()?.setSpellingIssues([]);
      this.host.bridge()?.setGrammarIssues?.([]);
      const bar = this.host.element().shadowRoot?.querySelector("docen-status-bar");
      bar?.setAttribute("proofing", "ok");
      this.#syncPane();
    }
  }

  /** Whether the checker runs (the options dialog pre-fills from here). */
  enabled(): boolean {
    return this.#enabled;
  }

  /** The live spelling issue list (the ribbon's spell-check command reads it). */
  issues(): SpellingIssue[] {
    return this.#issues;
  }

  /** The live grammar issue list. */
  grammarIssues(): GrammarIssue[] {
    return this.#grammarIssues;
  }

  /** Carry the issue ranges through a doc-changing transaction — the same
   *  mapping PM applies to selections, applied the moment the edit lands. */
  mapThrough(tr: Transaction): void {
    if (this.#issues.length > 0) {
      const mapped: SpellingIssue[] = [];
      for (const issue of this.#issues) {
        const from = tr.mapping.map(issue.from, -1);
        const to = tr.mapping.map(issue.to, 1);
        if (from >= to) continue;
        mapped.push(from === issue.from && to === issue.to ? issue : { ...issue, from, to });
      }
      this.#issues = mapped;
      this.host.bridge()?.setSpellingIssues(this.#issues);
    }

    if (this.#grammarIssues.length > 0) {
      const mappedG: GrammarIssue[] = [];
      for (const issue of this.#grammarIssues) {
        const from = tr.mapping.map(issue.from, -1);
        const to = tr.mapping.map(issue.to, 1);
        if (from >= to) continue;
        mappedG.push(from === issue.from && to === issue.to ? issue : { ...issue, from, to });
      }
      this.#grammarIssues = mappedG;
      this.host.bridge()?.setGrammarIssues?.(this.#grammarIssues);
    }
  }

  run(): void {
    const editor = this.host.editor();
    if (!editor) return;
    this.#issues = checkSpelling(editor.state.doc);
    const rawGrammar = checkGrammar(editor.state.doc);
    this.#grammarIssues = rawGrammar.filter(
      (g) => !this.#ignoredGrammar.has(`${g.from}:${g.ruleId}`),
    );

    this.host.bridge()?.setSpellingIssues(this.#issues);
    this.host.bridge()?.setGrammarIssues?.(this.#grammarIssues);

    this.#buildUnifiedEntries();
    this.#active = this.#unifiedEntries.length ? 0 : -1;

    const bar = this.host.element().shadowRoot?.querySelector("docen-status-bar");
    bar?.setAttribute(
      "proofing",
      this.#issues.length || this.#grammarIssues.length ? "issues" : "ok",
    );
    this.#syncPane();
  }

  #buildUnifiedEntries(): void {
    const editor = this.host.editor();
    const list: UnifiedProofingEntry[] = [];

    for (const issue of this.#issues) {
      list.push({
        kind: "spelling",
        from: issue.from,
        to: issue.to,
        word: issue.word,
        suggestions: spellSuggestions(issue.word, 5, issue.lang),
        category: "spelling",
      });
    }

    for (const g of this.#grammarIssues) {
      const word = editor
        ? editor.state.doc.textBetween(
            Math.min(g.from, editor.state.doc.content.size),
            Math.min(g.to, editor.state.doc.content.size),
          )
        : "";
      list.push({
        kind: "grammar",
        from: g.from,
        to: g.to,
        word: word || "Grammar error",
        suggestions: g.replacements,
        category: g.category,
        message: g.message,
        ruleId: g.ruleId,
      });
    }

    list.sort((a, b) => a.from - b.from);
    this.#unifiedEntries = list;
  }

  /** Push the active issue into the proofing pane when it's open. */
  #syncPane(): void {
    const pane = this.host.element().shadowRoot?.querySelector("docen-spelling-pane") as
      | (HTMLElement & {
          entries: Array<{
            word: string;
            suggestions: string[];
            category?: "spelling" | "grammar" | "style";
            message?: string;
          }>;
          active: number;
          total: number;
        })
      | null;
    if (!pane) return;
    const entries = this.#unifiedEntries;
    pane.total = entries.length;
    pane.entries = entries.map((e) => ({
      word: e.word,
      suggestions: e.suggestions,
      category: e.category,
      message: e.message,
    }));
    if (this.#active >= 0 && this.#active < entries.length) {
      pane.active = this.#active;
    } else {
      pane.active = -1;
    }
  }

  /** Select and scroll to a proofing issue (the pane / command navigation). */
  goto(index: number): void {
    const entries = this.#unifiedEntries;
    if (!entries.length) return;
    this.#active = ((index % entries.length) + entries.length) % entries.length;
    const entry = entries[this.#active];
    this.host.editor()?.commands.setTextSelection({ from: entry.from, to: entry.to });
    this.host.bridge()?.scrollIntoView(entry.from);
    this.#syncPane();
  }

  /** Make the spelling issue covering `pos` the active one (no scroll). */
  activateAt(pos: number): SpellingIssue | null {
    const hit = (p: number) => this.#issues.findIndex((i) => i.from <= p && p < i.to);
    const index = hit(pos) >= 0 ? hit(pos) : hit(pos - 1);
    if (index < 0) return null;
    const issue = this.#issues[index];
    const uIdx = this.#unifiedEntries.findIndex((e) => e.from === issue.from && e.to === issue.to);
    if (uIdx >= 0) this.#active = uIdx;
    this.host.editor()?.commands.setTextSelection({ from: issue.from, to: issue.to });
    this.#syncPane();
    return issue;
  }

  /** Make the grammar issue covering `pos` the active one (no scroll). */
  activateGrammarAt(pos: number): GrammarIssue | null {
    const hit = (p: number) => this.#grammarIssues.findIndex((i) => i.from <= p && p < i.to);
    const index = hit(pos) >= 0 ? hit(pos) : hit(pos - 1);
    if (index < 0) return null;
    const issue = this.#grammarIssues[index];
    const uIdx = this.#unifiedEntries.findIndex((e) => e.from === issue.from && e.to === issue.to);
    if (uIdx >= 0) this.#active = uIdx;
    this.host.editor()?.commands.setTextSelection({ from: issue.from, to: issue.to });
    this.#syncPane();
    return issue;
  }

  /** The pane's active issue index (the nav stepper offsets from it). */
  activeIndex(): number {
    return this.#active;
  }

  /** Replace the active issue's text with a suggestion. */
  replace(replacement: string): void {
    const entry = this.#unifiedEntries[this.#active];
    const editor = this.host.editor();
    if (!entry || !editor) return;
    editor.commands.insertContentAt({ from: entry.from, to: entry.to }, replacement);
    editor.commands.setTextSelection({ from: entry.from, to: entry.from + replacement.length });
  }

  /** The ignore levels (Word's pane and context menu). */
  ignore(mode: "once" | "ignore" | "add"): void {
    const entry = this.#unifiedEntries[this.#active];
    if (!entry) return;

    if (entry.kind === "spelling") {
      if (mode === "add") addSpellWord(entry.word);
      else if (mode === "once") ignoreSpellOnce({ from: entry.from, word: entry.word });
      else ignoreSpellWord(entry.word);
    } else if (entry.kind === "grammar") {
      this.#ignoredGrammar.add(`${entry.from}:${entry.ruleId}`);
    }
    this.run();
  }

  /** Ignore a specific grammar issue once. */
  ignoreGrammarOnce(issue: GrammarIssue): void {
    this.#ignoredGrammar.add(`${issue.from}:${issue.ruleId}`);
    this.run();
  }

  /** The pane's Ignore All: clears every issue for this session. */
  ignoreAll(): void {
    for (const issue of this.#issues) ignoreSpellWord(issue.word);
    for (const g of this.#grammarIssues) this.#ignoredGrammar.add(`${g.from}:${g.ruleId}`);
    this.run();
  }
}
