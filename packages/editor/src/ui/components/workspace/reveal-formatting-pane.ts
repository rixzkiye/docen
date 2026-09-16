import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang } from "../../i18n/localize";

export interface FormattingInfo {
  sampleText?: string;
  font?: {
    family: string;
    size: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
  };
  paragraph?: {
    alignment: string;
    indentLeft?: string;
    indentRight?: string;
    indentFirstLine?: string;
    spaceBefore?: string;
    spaceAfter?: string;
    lineSpacing?: string;
  };
  section?: {
    margins?: string;
    orientation?: string;
    paperSize?: string;
  };
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
  }
  .sample-box {
    padding: 8px;
    border: 1px solid var(--docen-color-divider, #e0e0e0);
    border-radius: 4px;
    background: var(--docen-color-neutral-background, #fafafa);
    margin-bottom: 12px;
    font-size: 13px;
    min-height: 24px;
    word-break: break-all;
  }
  .category {
    margin-bottom: 12px;
  }
  .cat-header {
    font-weight: 600;
    color: var(--docen-color-accent, #0f6cbd);
    border-bottom: 1px solid var(--docen-color-divider, #e0e0e0);
    padding-bottom: 2px;
    margin-bottom: 6px;
    text-transform: uppercase;
    font-size: 11px;
    letter-spacing: 0.5px;
  }
  .prop-row {
    display: flex;
    justify-content: space-between;
    padding: 2px 0;
  }
  .prop-name {
    color: var(--docen-color-foreground-secondary, #616161);
  }
  .prop-val {
    font-weight: 500;
  }
  .compare-toggle {
    margin-top: 8px;
    padding-top: 8px;
    border-top: 1px dashed var(--docen-color-divider, #ccc);
  }
`;

const template = html<DocenRevealFormattingPane>`
  <div class="sample-box" part="sample-box">
    ${(x) => x.formatting?.sampleText || "Selected Text"}
  </div>

  <div class="category">
    <div class="cat-header">Font</div>
    <div class="prop-row">
      <span class="prop-name">Font Family:</span>
      <span class="prop-val">${(x) => x.formatting?.font?.family ?? "Calibri"}</span>
    </div>
    <div class="prop-row">
      <span class="prop-name">Font Size:</span>
      <span class="prop-val">${(x) => x.formatting?.font?.size ?? "11 pt"}</span>
    </div>
    <div class="prop-row">
      <span class="prop-name">Styles:</span>
      <span class="prop-val">
        ${(x) =>
          [
            x.formatting?.font?.bold ? "Bold" : "",
            x.formatting?.font?.italic ? "Italic" : "",
            x.formatting?.font?.underline ? "Underline" : "",
          ]
            .filter(Boolean)
            .join(", ") || "Regular"}
      </span>
    </div>
    <div class="prop-row">
      <span class="prop-name">Color:</span>
      <span class="prop-val">${(x) => x.formatting?.font?.color ?? "Auto"}</span>
    </div>
  </div>

  <div class="category">
    <div class="cat-header">Paragraph</div>
    <div class="prop-row">
      <span class="prop-name">Alignment:</span>
      <span class="prop-val">${(x) => x.formatting?.paragraph?.alignment ?? "Left"}</span>
    </div>
    <div class="prop-row">
      <span class="prop-name">Indents:</span>
      <span class="prop-val">Left: ${(x) => x.formatting?.paragraph?.indentLeft ?? "0 pt"}</span>
    </div>
    <div class="prop-row">
      <span class="prop-name">Line Spacing:</span>
      <span class="prop-val"
        >${(x) => x.formatting?.paragraph?.lineSpacing ?? "Multiple 1.15"}</span
      >
    </div>
  </div>

  <div class="category">
    <div class="cat-header">Section</div>
    <div class="prop-row">
      <span class="prop-name">Margins:</span>
      <span class="prop-val">${(x) => x.formatting?.section?.margins ?? "Normal (1 in)"}</span>
    </div>
    <div class="prop-row">
      <span class="prop-name">Orientation:</span>
      <span class="prop-val">${(x) => x.formatting?.section?.orientation ?? "Portrait"}</span>
    </div>
  </div>

  <div class="compare-toggle">
    <fluent-checkbox ${ref("compareCheckbox")}> Compare to another selection </fluent-checkbox>
  </div>
`;

/**
 * `<docen-reveal-formatting-pane>` — Reveal Formatting task pane inspecting font,
 * paragraph, and section properties of the current caret or selection.
 */
@customElement({ name: "docen-reveal-formatting-pane", template, styles })
export class DocenRevealFormattingPane extends FASTElement {
  @observable formatting?: FormattingInfo;
  @observable compareCheckbox?: HTMLElement & { checked: boolean };

  #unsubscribe?: () => void;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#unsubscribe = observeLang(() => {});
  }

  override disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  setFormatting(info: FormattingInfo): void {
    this.formatting = { ...info };
  }
}
