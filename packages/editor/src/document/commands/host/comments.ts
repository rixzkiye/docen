import type { HostCommandDomain } from "./registry";

/** The comments domain's view of the host — only what its command bodies touch. */
export interface CommentsHostView {
  /** The comment commands (anchor/insert/edit/delete/jump). */
  insertComment(): void;
  editComment(): void;
  deleteComment(): void;
  jumpComment(direction: "previous" | "next"): void;
  /** Toggle a task pane (Word's comment sidebar). */
  togglePane(id: "comments"): void;
  /** Open/close a task pane without the toggle flip. */
  setTaskpane(id: "comments", open: boolean): void;
  getTaskpaneState(id: "comments"): boolean;
}

/**
 * Comment commands split out of the host element: insert/edit/delete on the
 * selection, prev/next jumps, and the pane toggles (the trailing title-bar
 * button and Review → Show Comments).
 */
export class CommentsHostCommands implements HostCommandDomain {
  constructor(private readonly host: CommentsHostView) {}

  readonly chrome: readonly string[] = [];

  readonly editor: readonly string[] = [
    "new-comment",
    "comment",
    "edit-comment",
    "delete-comment",
    "previous-comment",
    "next-comment",
    "show-comments",
  ];

  run(event: string, _value?: string): boolean {
    // New Comment — anchor the selection (or the word at the caret) with a
    // Word comment; Edit/Delete operate on the comment covering the selection.
    if (event === "new-comment") {
      this.host.insertComment();
      return true;
    }
    // The trailing title-bar "comment" button toggles the comments pane
    // (Word's sidebar) — it lists every comment, it does not create one.
    if (event === "comment") {
      this.host.togglePane("comments");
      return true;
    }
    if (event === "edit-comment") {
      this.host.editComment();
      return true;
    }
    if (event === "delete-comment") {
      this.host.deleteComment();
      return true;
    }
    if (event === "previous-comment") {
      this.host.jumpComment("previous");
      return true;
    }
    if (event === "next-comment") {
      this.host.jumpComment("next");
      return true;
    }
    // Review → Show Comments: toggle the comments pane (Word's sidebar).
    if (event === "show-comments") {
      this.host.setTaskpane("comments", !this.host.getTaskpaneState("comments"));
      return true;
    }
    return false;
  }
}
