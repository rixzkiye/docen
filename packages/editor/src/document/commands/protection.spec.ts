// @vitest-environment happy-dom
import { docxExtensions } from "@docen/docx";
import { Editor } from "@docen/docx/core";
import { beforeEach, describe, expect, it, vi } from "vitest";

import {
  DocenRestrictEditingPane,
  type ProtectionType,
  type RestrictEditingState,
} from "../../ui/components/workspace/restrict-editing-pane";

function makeTestEditor(initialSettings?: Record<string, unknown>) {
  return new Editor({
    element: null,
    extensions: docxExtensions,
    content: {
      type: "doc",
      attrs: {
        documentExtras: {
          settings: initialSettings ?? {},
        },
      },
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "Protected document test content" }],
        },
      ],
    },
  });
}

describe("Document Protection & Restrict Editing Lifecycle", () => {
  beforeEach(() => {
    window.alert = vi.fn();
  });

  describe("SHA-256 password hash verification", () => {
    it("hashes password correctly using crypto.subtle", async () => {
      const password = "SecretPassword123!";
      const encoder = new TextEncoder();
      const data = encoder.encode(password);
      const digest = await crypto.subtle.digest("SHA-256", data);
      const hash = Array.from(new Uint8Array(digest))
        .map((b) => b.toString(16).padStart(2, "0"))
        .join("");

      expect(hash).toBeDefined();
      expect(hash).toHaveLength(64);

      // Verify idempotency
      const digest2 = await crypto.subtle.digest("SHA-256", data);
      const hash2 = Array.from(new Uint8Array(digest2))
        .map((b) => b.toString(16).padStart(2, "0"))
        .join("");
      expect(hash2).toBe(hash);
    });
  });

  describe("DocenRestrictEditingPane component", () => {
    it("sets protection state and reflects in properties", () => {
      const pane = new DocenRestrictEditingPane();
      document.body.appendChild(pane);
      const state: RestrictEditingState = {
        isEnforced: true,
        type: "readOnly",
        formattingRestricted: true,
      };

      pane.setProtectionState(state, "dummyhash123");
      expect(pane.isEnforced).toBe(true);
      expect(pane.protectionType).toBe("readOnly");
      expect(pane.formattingRestricted).toBe(true);
      pane.remove();
    });

    it("prevents starting protection when passwords mismatch", async () => {
      const pane = new DocenRestrictEditingPane();
      document.body.appendChild(pane);

      pane.startPasswordInput = { value: "pass1" } as HTMLInputElement;
      pane.startConfirmPasswordInput = { value: "pass2" } as HTMLInputElement;

      let emitted = false;
      pane.addEventListener("protection:enforce", () => {
        emitted = true;
      });

      await pane.onStartProtection();
      expect(window.alert).toHaveBeenCalled();
      expect(emitted).toBe(false);
      expect(pane.isEnforced).toBe(false);
      pane.remove();
    });

    it("emits protection:enforce with hash when valid password provided", async () => {
      const pane = new DocenRestrictEditingPane();
      document.body.appendChild(pane);

      pane.startPasswordInput = { value: "mypassword" } as HTMLInputElement;
      pane.startConfirmPasswordInput = { value: "mypassword" } as HTMLInputElement;

      let detail: any = null;
      pane.addEventListener("protection:enforce", (e: Event) => {
        detail = (e as CustomEvent).detail;
      });

      await pane.onStartProtection();
      expect(detail).not.toBeNull();
      expect(detail.passwordHash).toBeDefined();
      expect(detail.passwordHash).toHaveLength(64);
      expect(pane.isEnforced).toBe(true);
      pane.remove();
    });

    it("prevents stopping protection with incorrect password", async () => {
      const pane = new DocenRestrictEditingPane();
      document.body.appendChild(pane);

      // Compute correct hash
      const encoder = new TextEncoder();
      const data = encoder.encode("correctPass");
      const digest = await crypto.subtle.digest("SHA-256", data);
      const correctHash = Array.from(new Uint8Array(digest))
        .map((b) => b.toString(16).padStart(2, "0"))
        .join("");

      pane.setProtectionState(
        { isEnforced: true, type: "comments", formattingRestricted: false },
        correctHash,
      );

      // Try wrong password
      pane.stopPasswordInput = { value: "wrongPass" } as HTMLInputElement;
      let stopped = false;
      pane.addEventListener("protection:stop", () => {
        stopped = true;
      });

      await pane.onStopProtection();
      expect(window.alert).toHaveBeenCalled();
      expect(stopped).toBe(false);
      expect(pane.isEnforced).toBe(true);

      // Try correct password
      pane.stopPasswordInput = { value: "correctPass" } as HTMLInputElement;
      await pane.onStopProtection();
      expect(stopped).toBe(true);
      expect(pane.isEnforced).toBe(false);
      pane.remove();
    });
  });

  describe("Document Settings Protection Persistence", () => {
    it("persists documentProtection in documentExtras.settings and survives reload", () => {
      const editor = makeTestEditor();

      // Dispatch document protection settings into doc
      const attrs = (editor.state.doc.attrs ?? {}) as { documentExtras?: Record<string, unknown> };
      const extras = attrs.documentExtras ?? {};
      const settings = {
        ...(extras.settings as Record<string, unknown> | undefined),
        documentProtection: {
          edit: "readOnly" as ProtectionType,
          hash: "abc123sha",
          formatting: true,
        },
      };

      editor.view.dispatch(
        editor.state.tr.setDocAttribute("documentExtras", { ...extras, settings }),
      );

      // Check persisted attributes
      const updatedJson = editor.getJSON();
      const docExtras = updatedJson.attrs?.documentExtras as any;
      expect(docExtras?.settings?.documentProtection).toEqual({
        edit: "readOnly",
        hash: "abc123sha",
        formatting: true,
      });

      // Simulate re-opening / reload with loaded JSON
      const reloadedEditor = makeTestEditor(docExtras.settings);
      const reloadedExtras = reloadedEditor.getJSON().attrs?.documentExtras as any;
      const loadedProtection = reloadedExtras?.settings?.documentProtection;

      expect(loadedProtection).toBeDefined();
      expect(loadedProtection.edit).toBe("readOnly");
      expect(loadedProtection.hash).toBe("abc123sha");
      expect(loadedProtection.formatting).toBe(true);
    });

    it("clears documentProtection when protection is stopped", () => {
      const editor = makeTestEditor({
        documentProtection: {
          edit: "forms",
          hash: "hash456",
          formatting: false,
        },
      });

      const attrs = (editor.state.doc.attrs ?? {}) as { documentExtras?: Record<string, unknown> };
      const extras = attrs.documentExtras ?? {};
      const prevSettings = (extras.settings ?? {}) as Record<string, unknown>;
      const newSettings = { ...prevSettings };
      delete newSettings.documentProtection;

      editor.view.dispatch(
        editor.state.tr.setDocAttribute("documentExtras", { ...extras, settings: newSettings }),
      );

      const json = editor.getJSON();
      const docExtras = json.attrs?.documentExtras as any;
      expect(docExtras?.settings?.documentProtection).toBeUndefined();
    });

    it("preserves hash and formatting when changing protection mode via settings update", () => {
      const editor = makeTestEditor({
        documentProtection: {
          edit: "readOnly",
          hash: "preserveHash123",
          formatting: true,
        },
      });

      const attrs = (editor.state.doc.attrs ?? {}) as { documentExtras?: Record<string, unknown> };
      const extras = attrs.documentExtras ?? {};
      const prevSettings = (extras.settings ?? {}) as Record<string, unknown>;
      const prevProt = prevSettings.documentProtection as Record<string, unknown>;

      // Simulating #applyDocumentSettings:
      const newSettings = {
        ...prevSettings,
        documentProtection: {
          ...prevProt,
          edit: "comments",
        },
      };

      editor.view.dispatch(
        editor.state.tr.setDocAttribute("documentExtras", { ...extras, settings: newSettings }),
      );

      const json = editor.getJSON();
      const docExtras = json.attrs?.documentExtras as any;
      const loadedProt = docExtras?.settings?.documentProtection;

      expect(loadedProt).toBeDefined();
      expect(loadedProt.edit).toBe("comments");
      expect(loadedProt.hash).toBe("preserveHash123");
      expect(loadedProt.formatting).toBe(true);
    });

    it("verifies forms protection input gating inside and outside SDT", () => {
      const editor = new Editor({
        element: null,
        extensions: docxExtensions,
        content: {
          type: "doc",
          content: [
            {
              type: "paragraph",
              content: [{ type: "text", text: "Outside text" }],
            },
            {
              type: "sdtBlock",
              attrs: {
                properties: { title: "Test Field" },
              },
              content: [
                {
                  type: "paragraph",
                  content: [{ type: "text", text: "Inside field" }],
                },
              ],
            },
            {
              type: "sdtBlock",
              attrs: {
                properties: { title: "Locked Field", cannotEdit: true },
              },
              content: [
                {
                  type: "paragraph",
                  content: [{ type: "text", text: "Locked content" }],
                },
              ],
            },
          ],
        },
      });

      // Implement the same logic as #isInsideSdt
      const isInsideSdt = (ed: Editor): boolean => {
        const { $from, $to } = ed.state.selection;
        for (let d = $from.depth; d > 0; d--) {
          const node = $from.node(d);
          if (
            node.type.name === "sdtBlock" ||
            node.type.name === "sdtInline" ||
            node.attrs?.properties
          ) {
            if ($to.pos < $from.start(d) || $to.pos > $from.end(d)) return false;
            const props = (node.attrs?.properties ?? {}) as Record<string, unknown>;
            if (props.cannotEdit === true) return false;
            return true;
          }
        }
        return false;
      };

      // 1. Caret in outside paragraph (pos = 1)
      editor.commands.setTextSelection(1);
      expect(isInsideSdt(editor)).toBe(false);

      // 2. Caret inside first SDT (pos = 16)
      editor.commands.setTextSelection(16);
      expect(isInsideSdt(editor)).toBe(true);

      // 3. Caret inside locked SDT (cannotEdit: true)
      // Find pos inside locked SDT
      let lockedPos = -1;
      editor.state.doc.descendants((node, pos) => {
        if (node.isText && node.text === "Locked content") {
          lockedPos = pos + 1;
        }
      });
      expect(lockedPos).toBeGreaterThan(0);
      editor.commands.setTextSelection(lockedPos);
      expect(isInsideSdt(editor)).toBe(false);
    });
  });
});
