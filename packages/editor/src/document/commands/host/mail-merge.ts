import type { MergeRecipients, MergeType } from "../mail-merge";
import type { HostCommandDomain } from "./registry";

/** The mail-merge domain's view of the host — only what its command bodies
 *  touch. */
export interface MailMergeHostView {
  /** The host element — the shadow-DOM root for the merge dialogs. */
  element(): HTMLElement;
  /** The pasted recipients grid (null before Select Recipients). */
  recipients(): MergeRecipients | null;
  insertAddressBlock(): void;
  insertGreetingLine(): void;
  togglePreview(): void;
  firstRecord(): void;
  lastRecord(): void;
  setMergeType(type: MergeType): void;
  /** Finish & Merge — assemble the output document. */
  finishMerge(mode: "edit" | "print" | "email"): void;
}

/**
 * Mail-merge commands split out of the host element: the recipients dialogs,
 * the merge-field seeds, the preview pass, and Finish & Merge.
 */
export class MailMergeHostCommands implements HostCommandDomain {
  constructor(private readonly host: MailMergeHostView) {}

  readonly chrome: readonly string[] = [];

  readonly editor: readonly string[] = [
    "select-recipients",
    "edit-recipients",
    "merge-field",
    "address-block",
    "greeting-line",
    "preview-results",
    "first-record",
    "last-record",
    "start-merge",
    "finish-merge",
  ];

  run(event: string, value?: string): boolean {
    // Mail merge — the recipients dialogs, the merge-field seeds, the preview
    // pass, and the Finish & Merge document (#merge / #finishMerge).
    if (event === "select-recipients" || event === "edit-recipients") {
      const root = this.host.element().shadowRoot;
      const mergeDialog = root?.querySelector("docen-merge-recipients-dialog") as {
        show(recipients: unknown): void;
      } | null;
      if (mergeDialog) {
        mergeDialog.show(this.host.recipients());
      } else {
        (
          root?.querySelector("docen-recipients-dialog") as {
            show(recipients: unknown): void;
          } | null
        )?.show(this.host.recipients());
      }
      return true;
    }
    if (event === "merge-field") {
      const recipients = this.host.recipients();
      (
        this.host.element().shadowRoot?.querySelector("docen-merge-field-dialog") as {
          show(headers: string[]): void;
        } | null
      )?.show(recipients?.headers ?? []);
      return true;
    }
    if (event === "address-block") {
      this.host.insertAddressBlock();
      return true;
    }
    if (event === "greeting-line") {
      this.host.insertGreetingLine();
      return true;
    }
    if (event === "preview-results") {
      this.host.togglePreview();
      return true;
    }
    if (event === "first-record") {
      this.host.firstRecord();
      return true;
    }
    if (event === "last-record") {
      this.host.lastRecord();
      return true;
    }
    if (event === "start-merge") {
      // The menu picks record the document kind; the face opens the
      // recipients dialog (the merge's first step).
      if (value === "letters" || value === "directory") this.host.setMergeType(value);
      else
        (
          this.host.element().shadowRoot?.querySelector("docen-recipients-dialog") as {
            show(recipients: unknown): void;
          } | null
        )?.show(this.host.recipients());
      return true;
    }
    if (event === "finish-merge") {
      this.host.finishMerge(value === "edit" || !value ? "edit" : (value as "print" | "email"));
      return true;
    }
    return false;
  }
}
