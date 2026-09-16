import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export type ProtectionType = "readOnly" | "trackedChanges" | "comments" | "forms";

export interface RestrictEditingState {
  isEnforced: boolean;
  type: ProtectionType;
  formattingRestricted: boolean;
}

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    height: 100%;
    padding: 12px;
    box-sizing: border-box;
    font-size: 12px;
    color: var(--docen-color-foreground, #242424);
    overflow-y: auto;
    gap: 16px;
  }
  .section {
    display: flex;
    flex-direction: column;
    gap: 8px;
  }
  .section-header {
    font-weight: 600;
    font-size: 13px;
    color: var(--docen-color-foreground, #242424);
    border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
    padding-bottom: 4px;
  }
  .num-step {
    font-weight: 700;
    margin-right: 4px;
  }
  .dropdown-select {
    padding: 6px 8px;
    border: 1px solid var(--docen-color-divider, #d1d1d1);
    border-radius: 4px;
    font-size: 12px;
    background: var(--docen-color-canvas, #ffffff);
    color: inherit;
    width: 100%;
    box-sizing: border-box;
  }
  .password-box {
    display: flex;
    flex-direction: column;
    gap: 6px;
    padding: 8px;
    background: var(--docen-color-neutral-background, #f7f7f7);
    border-radius: 4px;
    border: 1px solid var(--docen-color-divider, #e2e2e2);
  }
  .password-input {
    padding: 4px 8px;
    font-size: 12px;
    border: 1px solid var(--docen-color-divider, #ccc);
    border-radius: 3px;
  }
  .status-badge {
    padding: 4px 8px;
    border-radius: 3px;
    font-weight: 500;
    display: inline-block;
  }
  .status-badge.active {
    background: #def6e3;
    color: #0e700e;
  }
  .status-badge.inactive {
    background: #f0f0f0;
    color: #616161;
  }
`;

const template = html<DocenRestrictEditingPane>`
  <div class="section">
    <div class="section-header">
      <span class="num-step">1.</span> ${(x) => t("protect.formattingRestrictions", x)}
    </div>
    <fluent-checkbox
      ${ref("limitFormattingCheckbox")}
      ?disabled="${(x) => x.isEnforced}"
      ?checked="${(x) => x.formattingRestricted}"
    >
      ${(x) => t("protect.limitFormatting", x)}
    </fluent-checkbox>
  </div>

  <div class="section">
    <div class="section-header">
      <span class="num-step">2.</span> ${(x) => t("protect.editingRestrictions", x)}
    </div>
    <fluent-checkbox ${ref("allowEditingCheckbox")} ?disabled="${(x) => x.isEnforced}" checked>
      ${(x) => t("protect.allowOnlyThisEditing", x)}
    </fluent-checkbox>

    <select
      class="dropdown-select"
      ${ref("protectionTypeSelect")}
      ?disabled="${(x) => x.isEnforced}"
    >
      <option value="trackedChanges">${(x) => t("protect.trackedChanges", x)}</option>
      <option value="comments">${(x) => t("protect.comments", x)}</option>
      <option value="forms">${(x) => t("protect.fillingInForms", x)}</option>
      <option value="readOnly">${(x) => t("protect.noChangesReadOnly", x)}</option>
    </select>
  </div>

  <div class="section">
    <div class="section-header">
      <span class="num-step">3.</span> ${(x) => t("protect.startEnforcement", x)}
    </div>

    ${(x) =>
      x.isEnforced
        ? html<DocenRestrictEditingPane>`
            <div class="status-badge active">${(p) => t("protect.protectionActive", p)}</div>
            <p style="margin: 4px 0;">${(p) => t("protect.enforcedDesc", p)}</p>
            <div class="password-box">
              <label>${(p) => t("protect.enterPasswordToStop", p)}:</label>
              <input
                type="password"
                class="password-input"
                ${ref("stopPasswordInput")}
                placeholder="Password"
              />
            </div>
            <fluent-button appearance="accent" @click="${(p) => p.onStopProtection()}">
              ${(p) => t("protect.stopProtection", p)}
            </fluent-button>
          `
        : html<DocenRestrictEditingPane>`
            <div class="password-box">
              <label>${(p) => t("protect.passwordOptional", p)}:</label>
              <input
                type="password"
                class="password-input"
                ${ref("startPasswordInput")}
                placeholder="Password"
              />
              <input
                type="password"
                class="password-input"
                ${ref("startConfirmPasswordInput")}
                placeholder="Confirm password"
              />
            </div>
            <fluent-button appearance="accent" @click="${(p) => p.onStartProtection()}">
              ${(p) => t("protect.startEnforcingBtn", p)}
            </fluent-button>
          `}
  </div>
`;

@customElement({ name: "docen-restrict-editing-pane", template, styles })
export class DocenRestrictEditingPane extends FASTElement {
  @observable isEnforced = false;
  @observable protectionType: ProtectionType = "trackedChanges";
  @observable formattingRestricted = false;

  @observable limitFormattingCheckbox?: HTMLElement & { checked: boolean };
  @observable allowEditingCheckbox?: HTMLElement & { checked: boolean };
  @observable protectionTypeSelect?: HTMLSelectElement;
  @observable startPasswordInput?: HTMLInputElement;
  @observable startConfirmPasswordInput?: HTMLInputElement;
  @observable stopPasswordInput?: HTMLInputElement;

  #passwordHash?: string;
  #unsubscribe?: () => void;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#unsubscribe = observeLang(() => {});
  }

  override disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  setProtectionState(state: RestrictEditingState, passwordHash?: string): void {
    this.isEnforced = state.isEnforced;
    this.protectionType = state.type;
    this.formattingRestricted = state.formattingRestricted;
    this.#passwordHash = passwordHash;
    if (this.protectionTypeSelect) {
      this.protectionTypeSelect.value = state.type;
    }
  }

  async onStartProtection(): Promise<void> {
    const p1 = this.startPasswordInput?.value ?? "";
    const p2 = this.startConfirmPasswordInput?.value ?? "";

    if (p1 && p1 !== p2) {
      alert("Passwords do not match!");
      return;
    }

    let hash: string | undefined;
    if (p1) {
      // Use native Web Crypto API
      const encoder = new TextEncoder();
      const data = encoder.encode(p1);
      const digest = await crypto.subtle.digest("SHA-256", data);
      hash = Array.from(new Uint8Array(digest))
        .map((b) => b.toString(16).padStart(2, "0"))
        .join("");
    }

    const type = (this.protectionTypeSelect?.value as ProtectionType) || "trackedChanges";
    const formattingRestricted = Boolean(this.limitFormattingCheckbox?.checked);

    this.isEnforced = true;
    this.protectionType = type;
    this.formattingRestricted = formattingRestricted;
    this.#passwordHash = hash;

    this.$emit("protection:enforce", {
      type,
      formattingRestricted,
      passwordHash: hash,
    });
  }

  async onStopProtection(): Promise<void> {
    const entered = this.stopPasswordInput?.value ?? "";
    if (this.#passwordHash) {
      const encoder = new TextEncoder();
      const data = encoder.encode(entered);
      const digest = await crypto.subtle.digest("SHA-256", data);
      const hash = Array.from(new Uint8Array(digest))
        .map((b) => b.toString(16).padStart(2, "0"))
        .join("");

      if (hash !== this.#passwordHash) {
        alert("Incorrect password!");
        return;
      }
    }

    this.isEnforced = false;
    this.#passwordHash = undefined;

    this.$emit("protection:stop", {});
  }
}
