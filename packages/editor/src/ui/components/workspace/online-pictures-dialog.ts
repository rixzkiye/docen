import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { resolveOnlinePicture, type OnlinePictureFailure } from "../../../document/online-pictures";
import { probeImageSize } from "../../../document/watermark";
import { observeLang, t } from "../../i18n/localize";

/** A `fluent-text-input` widget plus its string value accessor. */
type FluentTextInput = HTMLElement & { value: string };

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(460px, 92vw);
  }
  .op-body {
    padding: 8px 4px 4px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    font-size: 13px;
  }
  .field {
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .field > label {
    white-space: nowrap;
  }
  fluent-text-input {
    min-width: 0;
    flex: 1 1 auto;
  }
  .op-preview {
    height: 180px;
    display: flex;
    align-items: center;
    justify-content: center;
    border: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    border-radius: 4px;
    background: var(--colorNeutralBackground2, #fafafa);
    overflow: hidden;
  }
  .op-preview img {
    max-width: 100%;
    max-height: 100%;
    object-fit: contain;
  }
  .op-status {
    min-height: 18px;
    color: var(--colorPaletteRedForeground1, #b10e1c);
  }
  .op-status[data-kind="hint"] {
    color: var(--colorNeutralForeground2, #444);
  }
`;

const template = html<DocenOnlinePicturesDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="op-body">
      <div class="field">
        <label ${ref("urlLabel")}></label>
        <fluent-text-input
          ${ref("urlInput")}
          @input="${(x) => x.preview()}"
          @keydown="${(x, c) => x.onKeyDown(c.event as KeyboardEvent)}"
        ></fluent-text-input>
      </div>
      <div class="op-preview">
        <img ${ref("previewEl")} hidden alt="" />
      </div>
      <span class="op-status" ${ref("statusEl")} data-kind="hint"></span>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button appearance="accent" ${ref("insertBtn")} @click="${(x) => x.insert()}">
      </fluent-button>
    </div>
  </docen-dialog>
`;

/**
 * `<docen-online-pictures-dialog>` — Insert → Pictures → Online Pictures.
 *
 * Word opens a Bing image search here; the browser element has no picture
 * service, so the dialog takes an image address instead (http/https or a
 * `data:image/…` URL), live-previews it, and downloads it at insert time with
 * size/type guards. A successful insert emits `online-picture:ok` with the
 * embedded picture; failures render as inline messages and the dialog stays
 * open so the address can be corrected.
 */
@customElement({ name: "docen-online-pictures-dialog", template, styles })
class DocenOnlinePicturesDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable urlLabel?: HTMLElement;
  @observable urlInput?: FluentTextInput;
  @observable previewEl?: HTMLImageElement;
  @observable statusEl?: HTMLElement;
  @observable insertBtn?: HTMLElement & { disabled?: boolean };
  @observable cancelBtn?: HTMLElement;

  #unobserveLang?: () => void;
  #busy = false;

  connectedCallback(): void {
    super.connectedCallback();
    this.#applyLabels();
    this.#unobserveLang = observeLang(() => this.#applyLabels());
  }

  disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    super.disconnectedCallback();
  }

  /** Open the dialog. Kept minimal — opening with no prefill is the rule; the
   *  previous address is cleared so each open starts clean. */
  show(): void {
    if (this.urlInput) this.urlInput.value = "";
    this.#busy = false;
    if (this.statusEl) {
      this.statusEl.textContent = t("onlinePictures.hint", this);
      this.statusEl.dataset.kind = "hint";
    }
    if (this.previewEl) {
      this.previewEl.hidden = true;
      this.previewEl.removeAttribute("src");
    }
    this.syncInsertEnabled();
    this.dialogEl?.show();
    this.urlInput?.focus?.();
  }

  hide(): void {
    if (this.#busy) return;
    this.dialogEl?.hide();
  }

  /** Live preview as the address changes — the browser's own image loader
   *  decides whether the URL resolves to something displayable. */
  preview(): void {
    const src = (this.urlInput?.value ?? "").trim();
    if (!this.previewEl) return;
    if (!src) {
      this.previewEl.hidden = true;
      this.previewEl.removeAttribute("src");
    } else {
      this.previewEl.hidden = false;
      this.previewEl.src = src;
    }
    this.syncInsertEnabled();
    this.#setStatus("", "hint");
  }

  /** Template-visible Enter handling (FAST cannot bind private methods). */
  onKeyDown(event: KeyboardEvent): void {
    if (event.key === "Enter") {
      event.preventDefault();
      void this.insert();
    }
  }

  /** Download + embed the address; emit it for the host to insert. */
  async insert(): Promise<void> {
    if (this.#busy) return;
    const url = (this.urlInput?.value ?? "").trim();
    if (!url) {
      this.#setStatus(t("onlinePictures.error.empty-url", this));
      return;
    }
    this.#busy = true;
    this.#applyLabels();
    this.#setStatus(t("onlinePictures.checking", this), "hint");
    const result = await resolveOnlinePicture(url, { probeSize: probeImageSize });
    this.#busy = false;
    this.#applyLabels();
    if (!result.ok) {
      this.#setStatus(this.#errorMessage(result.reason));
      return;
    }
    this.dialogEl?.hide();
    this.$emit("online-picture:ok", result.picture);
  }

  #errorMessage(reason: OnlinePictureFailure): string {
    return t(`onlinePictures.error.${reason}`, this);
  }

  #setStatus(text: string, kind: "hint" | "error" = "error"): void {
    if (!this.statusEl) return;
    this.statusEl.textContent = text;
    this.statusEl.dataset.kind = kind;
  }

  syncInsertEnabled(): void {
    if (!this.insertBtn) return;
    const empty = (this.urlInput?.value ?? "").trim() === "";
    this.insertBtn.toggleAttribute("disabled", empty || this.#busy);
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("onlinePictures.title", this);
    if (this.urlLabel) this.urlLabel.textContent = t("onlinePictures.address", this);
    if (this.urlInput) {
      this.urlInput.setAttribute("placeholder", t("onlinePictures.placeholder", this));
    }
    if (this.insertBtn) {
      this.insertBtn.textContent = this.#busy
        ? t("onlinePictures.inserting", this)
        : t("onlinePictures.insert", this);
    }
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    this.syncInsertEnabled();
  }
}

export default DocenOnlinePicturesDialog;
