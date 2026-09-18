import type { HostCommandDomain } from "./registry";

/** The file-I/O domain's view of the host — only what its command bodies touch. */
export interface FileIoHostView {
  /** Emit a cancelable host event; true = the host took the action over. */
  emitCancelable(name: "docen:save" | "docen:open" | "docen:new" | "docen:print"): boolean;
  saveAs(): Promise<void>;
  pickFile(): void;
  print(): Promise<void>;
  insertFileText(): void;
  /** File → New: reset to a blank document carrying the stored Set as Default
   *  formatting (called when no host handles `docen:new`). */
  newDocument(): void;
}

/**
 * File commands split out of the host element: save/new/open/print (the QAT
 * buttons re-emit the filename menu's actions) and Insert → Object → Text
 * from File.
 */
export class FileIoHostCommands implements HostCommandDomain {
  constructor(private readonly host: FileIoHostView) {}

  readonly chrome: readonly string[] = [];

  readonly editor: readonly string[] = ["save", "new", "open", "print", "insert-file-text"];

  run(event: string, _value?: string): boolean {
    if (event === "save") {
      if (!this.host.emitCancelable("docen:save")) void this.host.saveAs();
      return true;
    }
    // The QAT buttons re-emit the filename menu's file actions as commands
    // (those menu items ride change events instead) — same bodies as the
    // matching #onChange cases.
    if (event === "new") {
      if (!this.host.emitCancelable("docen:new")) this.host.newDocument();
      return true;
    }
    if (event === "open") {
      if (!this.host.emitCancelable("docen:open")) this.host.pickFile();
      return true;
    }
    if (event === "print") {
      if (!this.host.emitCancelable("docen:print")) void this.host.print();
      return true;
    }
    if (event === "insert-file-text") {
      this.host.insertFileText();
      return true;
    }
    return false;
  }
}
