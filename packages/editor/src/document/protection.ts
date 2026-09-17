import type { Editor } from "@docen/docx/core";

/**
 * Document-protection logic shared by the host handlers and the tests. The
 * view interface is the host surface the logic drives (DocenDocument
 * implements it with its editor + documentExtras channel; tests with a real
 * Editor and a real restrict-editing pane), so the behavior is exercised
 * through production code instead of a re-implementation.
 */

/** Whether the selection sits inside an editable content control (SDT): an
 *  sdtBlock/sdtInline ancestor (or any node carrying control properties) that
 *  fully contains the selection and is not locked with cannotEdit. */
export function isInsideEditableSdt(editor: Editor): boolean {
  const { $from, $to } = editor.state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const node = $from.node(d);
    if (node.type.name === "sdtBlock" || node.type.name === "sdtInline" || node.attrs?.properties) {
      if ($to.pos < $from.start(d) || $to.pos > $from.end(d)) return false;
      const props = (node.attrs?.properties ?? {}) as Record<string, unknown>;
      if (props.cannotEdit === true) return false;
      return true;
    }
  }
  return false;
}

/** The restrict-editing pane surface the logic syncs. */
export interface ProtectionPane {
  setProtectionState?(
    state: { isEnforced: boolean; type: string; formattingRestricted: boolean },
    hash?: string,
  ): void;
}

/** The host surface protection drives. */
export interface ProtectionHostView {
  editor(): Editor | undefined;
  /** The open document's settings slice (documentExtras.settings). */
  settings(): Record<string, unknown>;
  /** Persist a whole settings slice onto the document. */
  commitSettings(settings: Record<string, unknown>): void;
  /** Store the active restriction type + protected flag and re-derive
   *  editability/ribbon state. */
  setMode(mode: string | undefined, docProtected: boolean): void;
  pane(): ProtectionPane | null;
}

/** Fold a new protection mode into a settings slice WITHOUT dropping the
 *  existing hash/formatting fields (Options-dialog consistency). */
export function withProtection(
  settings: Record<string, unknown>,
  protection: string,
): Record<string, unknown> {
  const next = { ...settings };
  if (protection === "none") {
    delete next.documentProtection;
  } else {
    next.documentProtection = {
      ...(settings.documentProtection as object | undefined),
      edit: protection,
    };
  }
  return next;
}

/** Restrict Editing → "Yes, Start Enforcing Protection". */
export function enforceProtection(
  view: ProtectionHostView,
  detail: { type?: string; formattingRestricted?: boolean; passwordHash?: string },
): void {
  const type = detail?.type ?? "readOnly";
  const editor = view.editor();
  if (editor) {
    view.commitSettings({
      ...view.settings(),
      documentProtection: {
        edit: type,
        hash: detail?.passwordHash,
        formatting: Boolean(detail?.formattingRestricted),
      },
    });
    if (type === "trackedChanges") {
      const commands = editor.commands as unknown as Record<
        string,
        ((v?: unknown) => void) | undefined
      >;
      commands["track-changes"]?.(true);
    }
  }
  view.setMode(type, type === "readOnly" || type === "comments");
}

/** Restrict Editing → "Stop Protection" (password already verified by the pane). */
export function stopProtection(view: ProtectionHostView): void {
  const editor = view.editor();
  if (editor) {
    const settings = { ...view.settings() };
    delete settings.documentProtection;
    view.commitSettings(settings);
  }
  view.setMode(undefined, false);
}

/** Re-derive the mode/flags and re-sync the pane from the document settings —
 *  shared by the load boundary and the Options → Document path. */
export function applyProtectionMode(view: ProtectionHostView, protection: string): void {
  view.setMode(
    protection !== "none" ? protection : undefined,
    protection === "readOnly" || protection === "comments",
  );
  const pane = view.pane();
  if (!pane) return;
  const docProt = view.settings().documentProtection as
    | { formatting?: boolean; hash?: string }
    | undefined;
  const isEnforced = protection !== "none";
  pane.setProtectionState?.(
    {
      isEnforced,
      type: isEnforced ? protection : "trackedChanges",
      formattingRestricted: Boolean(docProt?.formatting),
    },
    docProt?.hash,
  );
}
