import type { JSONContent } from "@docen/docx";
import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export interface DocumentVersionSnapshot {
  id: string;
  timestamp: string;
  author: string;
  isAutosave: boolean;
  doc: JSONContent;
  summary?: string;
}

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(720px, 94vw);
    height: 520px;
  }
  .history-container {
    display: flex;
    gap: 12px;
    height: 380px;
    padding: 8px 4px;
    font-size: 13px;
  }
  .version-list {
    width: 240px;
    flex-shrink: 0;
    overflow-y: auto;
    border: 1px solid var(--docen-color-divider, #e0e0e0);
    border-radius: 4px;
    display: flex;
    flex-direction: column;
  }
  .version-card {
    padding: 10px 12px;
    border-bottom: 1px solid var(--docen-color-divider, #f0f0f0);
    cursor: pointer;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .version-card:hover {
    background: var(--colorNeutralBackground1Hover, #f5f5f5);
  }
  .version-card.selected {
    background: var(--docen-color-accent-subtle, #ebf3fc);
    border-left: 3px solid var(--docen-color-accent, #0f6cbd);
  }
  .version-time {
    font-weight: 600;
    font-size: 12px;
  }
  .version-meta {
    font-size: 11px;
    color: var(--docen-color-foreground-secondary, #616161);
    display: flex;
    justify-content: space-between;
  }
  .version-tag {
    font-size: 10px;
    padding: 1px 4px;
    border-radius: 2px;
    background: #e1dfdd;
  }
  .preview-pane {
    flex: 1;
    overflow-y: auto;
    border: 1px solid var(--docen-color-divider, #e0e0e0);
    border-radius: 4px;
    padding: 14px;
    background: var(--docen-color-canvas, #ffffff);
    font-family: inherit;
    line-height: 1.5;
  }
  .preview-empty {
    color: #888;
    text-align: center;
    margin-top: 60px;
  }
  .footer {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 12px;
  }
`;

const template = html<DocenVersionHistoryDialog>`
  <docen-dialog ${ref("dialog")} modal heading="${(x) => t("history.dialogTitle", x)}">
    <div class="history-container">
      <div class="version-list" ${ref("listEl")}></div>
      <div class="preview-pane" ${ref("previewEl")}>
        <div class="preview-empty">${(x) => t("history.selectVersionToPreview", x)}</div>
      </div>
    </div>
    <div class="footer" slot="action">
      <fluent-button
        appearance="accent"
        ?disabled="${(x) => !x.selectedVersion}"
        @click="${(x) => x.onRestore()}"
      >
        ${(x) => t("history.restoreBtn", x)}
      </fluent-button>
      <fluent-button @click="${(x) => x.close()}">${(x) => t("dialog.close", x)}</fluent-button>
    </div>
  </docen-dialog>
`;

@customElement({ name: "docen-version-history-dialog", template, styles })
export class DocenVersionHistoryDialog extends FASTElement {
  @observable dialog?: HTMLElement & { open: boolean };
  @observable listEl?: HTMLElement;
  @observable previewEl?: HTMLElement;
  @observable selectedVersion?: DocumentVersionSnapshot;

  #versions: DocumentVersionSnapshot[] = [];
  #unsubscribe?: () => void;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#unsubscribe = observeLang(() => {});
  }

  override disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  show(versions: DocumentVersionSnapshot[]): void {
    this.#versions = [...versions].sort(
      (a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime(),
    );
    this.selectedVersion = this.#versions[0];
    this.#renderList();
    this.#renderPreview();
    if (this.dialog) this.dialog.open = true;
  }

  close(): void {
    if (this.dialog) this.dialog.open = false;
  }

  #selectVersion(v: DocumentVersionSnapshot): void {
    this.selectedVersion = v;
    this.#renderList();
    this.#renderPreview();
  }

  #renderList(): void {
    if (!this.listEl) return;
    this.listEl.replaceChildren();

    for (const v of this.#versions) {
      const card = document.createElement("div");
      card.className = `version-card ${this.selectedVersion?.id === v.id ? "selected" : ""}`;

      const timeEl = document.createElement("div");
      timeEl.className = "version-time";
      try {
        timeEl.textContent = new Date(v.timestamp).toLocaleString();
      } catch {
        timeEl.textContent = v.timestamp;
      }

      const metaEl = document.createElement("div");
      metaEl.className = "version-meta";
      metaEl.textContent = v.author || "User";

      const tag = document.createElement("span");
      tag.className = "version-tag";
      tag.textContent = v.isAutosave ? t("history.autosave", this) : t("history.manualSave", this);
      metaEl.appendChild(tag);

      card.appendChild(timeEl);
      card.appendChild(metaEl);

      card.addEventListener("click", () => this.#selectVersion(v));
      this.listEl.appendChild(card);
    }
  }

  #renderPreview(): void {
    if (!this.previewEl) return;
    this.previewEl.replaceChildren();

    if (!this.selectedVersion) {
      const empty = document.createElement("div");
      empty.className = "preview-empty";
      empty.textContent = t("history.selectVersionToPreview", this);
      this.previewEl.appendChild(empty);
      return;
    }

    const doc = this.selectedVersion.doc;
    const paras = doc.content ?? [];
    for (const node of paras) {
      if (node.type === "paragraph") {
        const p = document.createElement("p");
        p.style.margin = "0 0 8px 0";
        p.textContent = (node.content ?? []).map((c) => c.text ?? "").join("");
        this.previewEl.appendChild(p);
      }
    }
  }

  onRestore(): void {
    if (!this.selectedVersion) return;
    this.$emit("version:restore", {
      versionId: this.selectedVersion.id,
      doc: this.selectedVersion.doc,
    });
    this.close();
  }
}
