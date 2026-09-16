import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export type GoToKind =
  | "page"
  | "section"
  | "line"
  | "bookmark"
  | "footnote"
  | "endnote"
  | "heading";

export interface GoToPayload {
  kind: GoToKind;
  target: string;
  direction?: -1 | 1;
}

type FluentTextInput = HTMLElement & { value: string };

const GO_TO_KINDS: readonly GoToKind[] = [
  "page",
  "section",
  "line",
  "bookmark",
  "footnote",
  "endnote",
  "heading",
];

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(460px, 92vw);
  }
  .goto-body {
    padding: 8px 4px 4px;
    display: flex;
    gap: 16px;
    font-size: 13px;
  }
  .kinds-pane {
    width: 140px;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .pane-label {
    font-weight: 600;
    margin-bottom: 2px;
  }
  .kinds-list {
    height: 180px;
    border: 1px solid var(--colorNeutralStroke2, #d1d1d1);
    border-radius: 4px;
    overflow-y: auto;
    background: var(--colorNeutralBackground1, #fff);
    display: flex;
    flex-direction: column;
  }
  .kind-item {
    padding: 6px 10px;
    cursor: pointer;
    user-select: none;
    font-size: 13px;
  }
  .kind-item:hover {
    background: var(--colorNeutralBackground1Hover, #f0f0f0);
  }
  .kind-item.selected {
    background: var(--colorBrandBackground2, #e0eafc);
    font-weight: 600;
  }
  .input-pane {
    flex: 1 1 0;
    min-width: 0;
    display: flex;
    flex-direction: column;
    gap: 8px;
  }
  fluent-text-input {
    width: 100%;
  }
  .hint-text {
    font-size: 12px;
    color: var(--colorNeutralForeground3, #666);
    line-height: 1.4;
  }
`;

const template = html<DocenGoToDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="goto-body">
      <div class="kinds-pane">
        <div class="pane-label" ${ref("whatLabel")}></div>
        <div class="kinds-list" ${ref("kindsContainer")}></div>
      </div>

      <div class="input-pane">
        <label class="pane-label" ${ref("enterLabel")}></label>
        <fluent-text-input
          ${ref("targetInput")}
          type="text"
          @keydown="${(x, c) => x.onKeydown(c.event as KeyboardEvent)}"
        ></fluent-text-input>
        <div class="hint-text" ${ref("hintEl")}></div>
      </div>
    </div>

    <div slot="action">
      <fluent-button
        appearance="accent"
        ${ref("gotoBtn")}
        @click="${(x) => x.onGoTo()}"
      ></fluent-button>
      <fluent-button ${ref("prevBtn")} @click="${(x) => x.onNavigate(-1)}"></fluent-button>
      <fluent-button ${ref("nextBtn")} @click="${(x) => x.onNavigate(1)}"></fluent-button>
      <fluent-button ${ref("closeBtn")} @click="${(x) => x.hide()}"></fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-go-to-dialog>` — Word's Go To dialog (Ctrl+G, Find & Replace Go To tab).
 * Navigates by Page, Section, Line, Bookmark, Footnote, Endnote, or Heading.
 */
@customElement({ name: "docen-go-to-dialog", template, styles })
export class DocenGoToDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable whatLabel?: HTMLElement;
  @observable kindsContainer?: HTMLElement;
  @observable enterLabel?: HTMLElement;
  @observable targetInput?: FluentTextInput;
  @observable hintEl?: HTMLElement;
  @observable gotoBtn?: HTMLElement;
  @observable prevBtn?: HTMLElement;
  @observable nextBtn?: HTMLElement;
  @observable closeBtn?: HTMLElement;

  @observable activeKind: GoToKind = "page";

  #unobserveLang?: () => void;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#renderKinds();
    this.#unobserveLang = observeLang(() => {
      this.#applyLabels();
      this.#renderKinds();
    });
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  show(initialKind: GoToKind = "page"): void {
    this.activeKind = initialKind;
    if (this.targetInput) this.targetInput.value = "";
    this.#applyTargetLabels();
    this.#renderKinds();
    this.dialogEl?.show();
    setTimeout(() => {
      this.targetInput?.focus();
    }, 50);
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  selectKind(kind: GoToKind): void {
    this.activeKind = kind;
    this.#applyTargetLabels();
    this.#highlightSelectedKind();
    this.targetInput?.focus();
  }

  onKeydown(e: KeyboardEvent): void {
    if (e.key === "Enter") {
      e.preventDefault();
      this.onGoTo();
    }
  }

  onGoTo(): void {
    const target = this.targetInput?.value.trim() ?? "";
    this.$emit("goto:navigate", {
      kind: this.activeKind,
      target,
    } satisfies GoToPayload);
  }

  onNavigate(direction: -1 | 1): void {
    const target = this.targetInput?.value.trim() ?? "";
    this.$emit("goto:navigate", {
      kind: this.activeKind,
      target,
      direction,
    } satisfies GoToPayload);
  }

  #renderKinds(): void {
    if (!this.kindsContainer) return;
    this.kindsContainer.innerHTML = "";
    for (const kind of GO_TO_KINDS) {
      const el = document.createElement("div");
      el.className = "kind-item" + (kind === this.activeKind ? " selected" : "");
      el.textContent = t(`goto.${kind}`, this);
      el.addEventListener("click", () => this.selectKind(kind));
      this.kindsContainer.appendChild(el);
    }
  }

  #highlightSelectedKind(): void {
    if (!this.kindsContainer) return;
    const items = this.kindsContainer.querySelectorAll<HTMLElement>(".kind-item");
    items.forEach((el, i) => {
      el.classList.toggle("selected", GO_TO_KINDS[i] === this.activeKind);
    });
  }

  #applyTargetLabels(): void {
    if (this.enterLabel) {
      this.enterLabel.textContent = t(`goto.enter-${this.activeKind}`, this);
    }
    if (this.hintEl) {
      if (
        this.activeKind === "page" ||
        this.activeKind === "section" ||
        this.activeKind === "line" ||
        this.activeKind === "footnote" ||
        this.activeKind === "endnote" ||
        this.activeKind === "heading"
      ) {
        this.hintEl.textContent =
          "+N moves forward N, -N moves backward N relative to the current position.";
      } else {
        this.hintEl.textContent = "";
      }
    }
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("goto.title", this);
    if (this.whatLabel) this.whatLabel.textContent = t("goto.what", this);
    if (this.gotoBtn) this.gotoBtn.textContent = t("goto.btn-goto", this);
    if (this.prevBtn) this.prevBtn.textContent = t("goto.btn-prev", this);
    if (this.nextBtn) this.nextBtn.textContent = t("goto.btn-next", this);
    if (this.closeBtn) this.closeBtn.textContent = t("goto.btn-close", this);
    this.#applyTargetLabels();
  }
}

export default DocenGoToDialog;
