import { DELEGATION_NOTICE } from "@docen/docx";
import { FASTElement, css, customElement, html, observable, when } from "@microsoft/fast-element";

export { DELEGATION_NOTICE };

export interface AltTextTarget {
  kind?: string;
  title?: string;
  descr?: string;
}

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    flex: 1;
    min-height: 0;
    box-sizing: border-box;
    font-size: 12px;
    font-family: var(--docen-font-family, system-ui, -apple-system, sans-serif);
    color: var(--docen-color-foreground, #242424);
    padding: 16px;
    gap: 16px;
    overflow-y: auto;
  }

  .pane-header {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .pane-title {
    font-size: 14px;
    font-weight: 600;
    color: var(--docen-color-foreground, #242424);
  }

  .pane-help {
    font-size: 11px;
    line-height: 1.4;
    color: var(--docen-color-foreground-secondary, #616161);
  }

  .notice-box {
    padding: 10px 12px;
    background-color: var(--docen-color-info-subtle, #f0f4f9);
    border: 1px solid var(--docen-color-info-border, #d0e0f5);
    border-radius: 4px;
    font-size: 11px;
    line-height: 1.45;
    color: var(--docen-color-foreground, #1b3a57);
    display: flex;
    gap: 8px;
    align-items: flex-start;
  }

  .notice-icon {
    font-size: 14px;
    line-height: 1;
    flex-shrink: 0;
    margin-top: 1px;
  }

  .field-group {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .field-label {
    font-weight: 600;
    font-size: 12px;
    color: var(--docen-color-foreground, #242424);
  }

  .text-input,
  .text-area {
    width: 100%;
    box-sizing: border-box;
    border: 1px solid var(--docen-color-stroke-1, #c7c7c7);
    border-radius: 4px;
    padding: 6px 8px;
    font-size: 12px;
    font-family: inherit;
    background: var(--docen-color-background, #fff);
    color: var(--docen-color-foreground, #242424);
    outline: none;
    transition: border-color 0.15s ease-in-out;
  }

  .text-input:focus,
  .text-area:focus {
    border-color: var(--docen-color-accent, #0f6cbd);
  }

  .text-area {
    resize: vertical;
    min-height: 72px;
  }

  .actions {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 4px;
  }

  .apply-btn {
    padding: 6px 14px;
    background-color: var(--docen-color-accent, #0f6cbd);
    color: #fff;
    border: none;
    border-radius: 4px;
    font-size: 12px;
    font-weight: 500;
    cursor: pointer;
  }

  .apply-btn:hover {
    background-color: var(--docen-color-accent-hover, #115ea3);
  }
`;

const template = html<DocenAltTextPane>`
  <div class="pane-header">
    <div class="pane-title">Alt Text</div>
    <div class="pane-help">
      How would you describe this object and its context to someone who is blind or has low vision?
    </div>
  </div>

  ${when(
    (x) => x.showNotice,
    html<DocenAltTextPane>`
      <div class="notice-box" role="status">
        <span class="notice-icon" aria-hidden="true">ℹ️</span>
        <span class="notice-text">${() => DELEGATION_NOTICE}</span>
      </div>
    `,
  )}

  <div class="field-group">
    <label class="field-label" for="alt-title-input">Title</label>
    <input
      id="alt-title-input"
      class="text-input"
      type="text"
      placeholder="Enter a brief title"
      :value="${(x) => x.titleText}"
      @input="${(x, c) => x.handleTitleInput(c.event)}"
    />
  </div>

  <div class="field-group">
    <label class="field-label" for="alt-descr-input">Description</label>
    <textarea
      id="alt-descr-input"
      class="text-area"
      rows="4"
      placeholder="Detailed description of the drawing"
      :value="${(x) => x.descrText}"
      @input="${(x, c) => x.handleDescrInput(c.event)}"
    ></textarea>
  </div>

  <div class="actions">
    <button class="apply-btn" type="button" @click="${(x) => x.apply()}">Apply</button>
  </div>
`;

@customElement({ name: "docen-alt-text-pane", template, styles })
export class DocenAltTextPane extends FASTElement {
  @observable titleText = "";
  @observable descrText = "";
  @observable targetKind = "";
  @observable showNotice = false;

  get title(): string {
    return this.titleText;
  }

  set title(val: string) {
    this.titleText = val ?? "";
  }

  get description(): string {
    return this.descrText;
  }

  set description(val: string) {
    this.descrText = val ?? "";
  }

  get descr(): string {
    return this.descrText;
  }

  set descr(val: string) {
    this.descrText = val ?? "";
  }

  setTarget(target: AltTextTarget | null): void {
    if (!target) {
      this.targetKind = "";
      this.titleText = "";
      this.descrText = "";
      this.showNotice = false;
      return;
    }
    this.targetKind = target.kind ?? "";
    this.titleText = target.title ?? "";
    this.descrText = target.descr ?? "";
    this.showNotice = target.kind === "model3d" || target.kind === "ink";
  }

  handleTitleInput(e: Event): void {
    const target = e.target as HTMLInputElement;
    this.titleText = target.value;
    this.$emit("alt-text:change", {
      title: this.titleText,
      descr: this.descrText,
      kind: this.targetKind,
    });
  }

  handleDescrInput(e: Event): void {
    const target = e.target as HTMLTextAreaElement;
    this.descrText = target.value;
    this.$emit("alt-text:change", {
      title: this.titleText,
      descr: this.descrText,
      kind: this.targetKind,
    });
  }

  apply(): void {
    this.$emit("alt-text:apply", {
      title: this.titleText,
      descr: this.descrText,
      kind: this.targetKind,
    });
  }
}

export default DocenAltTextPane;
