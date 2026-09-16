import type { Editor } from "@docen/docx/core";

import type { SpellingIssue } from "../../spelling";
import type { HostCommandDomain } from "./registry";

/** The proofing/language domain's view of the host — only what its command
 *  bodies touch. */
export interface ProofingLanguageHostView {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** Open the Word Count statistics dialog. */
  showWordCount(): void;
  /** Re-run the spelling check and rebuild the proofing pane. */
  spellingRun(): void;
  /** Open/close a task pane (the proofing pane). */
  setTaskpane(id: "proofing", open: boolean): void;
  spellingIssues(): SpellingIssue[];
  spellingGoto(index: number): void;
  spellingReplace(replacement: string): void;
  spellingIgnore(mode: "once" | "ignore" | "add"): void;
  /** Open the proofing-language dialog for the selection. */
  openLanguageDialog(): void;
}

/**
 * Proofing and language commands split out of the host element: Word Count,
 * Spelling & Grammar (check/replace/ignore), and the proofing-language dialog.
 */
export class ProofingLanguageHostCommands implements HostCommandDomain {
  constructor(private readonly host: ProofingLanguageHostView) {}

  readonly chrome: readonly string[] = ["word-count"];

  readonly editor: readonly string[] = [
    "spell-check",
    "spell-pick",
    "spell-ignore-once",
    "spell-ignore-all",
    "spell-add",
    "language",
  ];

  run(event: string, value?: string): boolean {
    // Word Count (ribbon Review → Proofing) → the statistics dialog.
    if (event === "word-count") {
      this.host.showWordCount();
      return true;
    }
    const editor = this.host.editor();
    if (!editor) return false;
    // Spelling (Review → Spelling & Grammar, F7, the status-bar book):
    // re-check now, open the pane, and start at the first issue at/after the
    // caret (Word starts checking from the insertion point).
    if (event === "spell-check") {
      this.host.spellingRun();
      this.host.setTaskpane("proofing", true);
      const from = editor.state.selection.from;
      const issues = this.host.spellingIssues();
      if (issues.length) {
        const first = issues.find((issue) => issue.from >= from) ?? issues[0];
        this.host.spellingGoto(issues.indexOf(first));
      }
      return true;
    }
    // The context menu's spelling group: replace with the picked suggestion
    // or apply one of the three ignore levels (the issue was activated when
    // the menu was built).
    if (event === "spell-pick") {
      this.host.spellingReplace(value ?? "");
      return true;
    }
    if (event === "spell-ignore-once" || event === "spell-ignore-all" || event === "spell-add") {
      this.host.spellingIgnore(
        event === "spell-ignore-once" ? "once" : event === "spell-ignore-all" ? "ignore" : "add",
      );
      return true;
    }
    // Language (Review → Language, the status-bar language item): the
    // proofing-language dialog for the selection.
    if (event === "language") {
      this.host.openLanguageDialog();
      return true;
    }
    return false;
  }
}
