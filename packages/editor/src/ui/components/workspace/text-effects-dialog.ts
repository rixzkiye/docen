import type { TextEffects } from "@docen/docx";
import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

const BEVEL_PRESETS = [
  "circle",
  "relaxedInset",
  "cross",
  "coolSlant",
  "angle",
  "softRound",
  "convex",
  "slope",
] as const;

const styles = css`
  :host {
    display: contents;
  }
  docen-dialog::part(dialog) {
    width: min(560px, 94vw);
  }
  .te-body {
    padding: 4px 4px 8px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 13px;
    max-height: min(70vh, 640px);
    overflow: auto;
  }
  .te-section {
    border: 1px solid var(--colorNeutralStroke2, #e0e0e0);
    border-radius: 4px;
    padding: 8px 10px;
  }
  .te-section > label.te-enable {
    display: flex;
    align-items: center;
    gap: 6px;
    font-weight: 600;
    margin-bottom: 6px;
  }
  .te-row {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
  }
  .te-row > label {
    color: var(--colorNeutralForeground2, #444);
  }
  input[type="number"] {
    width: 64px;
  }
  input[type="color"] {
    width: 36px;
    height: 24px;
    padding: 0;
    border: 1px solid var(--colorNeutralStroke1, #ccc);
    background: none;
  }
  input[type="range"] {
    width: 110px;
  }
  select {
    min-width: 90px;
  }
`;

const template = html<DocenTextEffectsDialog>`
  <docen-dialog ${ref("dialogEl")}>
    <div class="te-body">
      <section class="te-section">
        <label class="te-enable">
          <input type="checkbox" data-field="outline-on" />
          <span data-i18n="textEffects.outline"></span>
        </label>
        <div class="te-row">
          <label data-i18n="textEffects.color"></label>
          <input type="color" data-field="outline-color" value="#000000" />
          <label data-i18n="textEffects.width"></label>
          <input
            type="number"
            data-field="outline-width"
            min="0.25"
            max="20"
            step="0.25"
            value="1"
          />
          <label data-i18n="textEffects.dash"></label>
          <select data-field="outline-dash">
            <option value="solid" data-i18n="textEffects.dash-solid"></option>
            <option value="dash" data-i18n="textEffects.dash-dash"></option>
            <option value="dot" data-i18n="textEffects.dash-dot"></option>
          </select>
        </div>
      </section>

      <section class="te-section">
        <label class="te-enable">
          <input type="checkbox" data-field="shadow-on" />
          <span data-i18n="textEffects.shadow"></span>
        </label>
        <div class="te-row">
          <label data-i18n="textEffects.color"></label>
          <input type="color" data-field="shadow-color" value="#000000" />
          <label data-i18n="textEffects.transparency"></label>
          <input type="range" data-field="shadow-transparency" min="0" max="100" value="40" />
          <label data-i18n="textEffects.blur"></label>
          <input type="number" data-field="shadow-blur" min="0" max="50" step="0.5" value="2" />
          <label data-i18n="textEffects.distance"></label>
          <input type="number" data-field="shadow-distance" min="0" max="50" step="0.5" value="3" />
          <label data-i18n="textEffects.angle"></label>
          <input type="number" data-field="shadow-angle" min="0" max="359" step="1" value="45" />
        </div>
      </section>

      <section class="te-section">
        <label class="te-enable">
          <input type="checkbox" data-field="reflection-on" />
          <span data-i18n="textEffects.reflection"></span>
        </label>
        <div class="te-row">
          <label data-i18n="textEffects.transparency"></label>
          <input type="range" data-field="reflection-transparency" min="0" max="100" value="50" />
          <label data-i18n="textEffects.blur"></label>
          <input
            type="number"
            data-field="reflection-blur"
            min="0"
            max="50"
            step="0.5"
            value="0.5"
          />
          <label data-i18n="textEffects.distance"></label>
          <input
            type="number"
            data-field="reflection-distance"
            min="0"
            max="50"
            step="0.5"
            value="0"
          />
        </div>
      </section>

      <section class="te-section">
        <label class="te-enable">
          <input type="checkbox" data-field="glow-on" />
          <span data-i18n="textEffects.glow"></span>
        </label>
        <div class="te-row">
          <label data-i18n="textEffects.color"></label>
          <input type="color" data-field="glow-color" value="#4472c4" />
          <label data-i18n="textEffects.transparency"></label>
          <input type="range" data-field="glow-transparency" min="0" max="100" value="40" />
          <label data-i18n="textEffects.size"></label>
          <input type="number" data-field="glow-size" min="1" max="100" step="0.5" value="5" />
        </div>
      </section>

      <section class="te-section">
        <label class="te-enable">
          <input type="checkbox" data-field="bevel-on" />
          <span data-i18n="textEffects.bevel"></span>
        </label>
        <div class="te-row">
          <label data-i18n="textEffects.preset"></label>
          <select data-field="bevel-preset">
            ${BEVEL_PRESETS.map((preset) => html`<option value="${preset}">${preset}</option>`)}
          </select>
          <label data-i18n="textEffects.width"></label>
          <input type="number" data-field="bevel-width" min="0" max="100" step="0.5" value="5.5" />
          <label data-i18n="textEffects.height"></label>
          <input type="number" data-field="bevel-height" min="0" max="100" step="0.5" value="5.5" />
        </div>
      </section>

      <section class="te-section">
        <label class="te-enable">
          <input type="checkbox" data-field="rotation-on" />
          <span data-i18n="textEffects.rotation"></span>
        </label>
        <div class="te-row">
          <label data-i18n="textEffects.x"></label>
          <input type="number" data-field="rotation-x" min="-180" max="180" step="1" value="0" />
          <label data-i18n="textEffects.y"></label>
          <input type="number" data-field="rotation-y" min="-180" max="180" step="1" value="0" />
          <label data-i18n="textEffects.z"></label>
          <input type="number" data-field="rotation-z" min="-180" max="180" step="1" value="0" />
        </div>
      </section>
    </div>
    <div slot="action">
      <fluent-button ${ref("cancelBtn")} @click="${(x) => x.hide()}"></fluent-button>
      <fluent-button appearance="accent" ${ref("okBtn")} @click="${(x) => x.submit()}">
      </fluent-button>
    </div>
  </docen-dialog>
`;

const hexOf = (color: string): string => color.replace(/^#/, "").toUpperCase();

/** pt → px (Word's dialog units are points; the model stores px). */
const ptToPx = (pt: number): number => (pt * 96) / 72;

/**
 * `<docen-text-effects-dialog>` — Word's "Format Text Effects" surface for the
 * Font group's Text Effects gallery. The dialog edits all six families
 * (outline, shadow, reflection, glow, bevel, 3-D rotation) and emits
 * `text-effects:ok` with a per-family patch (`null` clears); the host applies
 * it to the selected runs through the `text-effects-apply` command.
 */
@customElement({ name: "docen-text-effects-dialog", template, styles })
class DocenTextEffectsDialog extends FASTElement {
  @observable dialogEl?: HTMLElement & { heading?: string; show(): void; hide(): void };
  @observable okBtn?: HTMLElement;
  @observable cancelBtn?: HTMLElement;

  #unobserveLang?: () => void;

  /** Template query that tolerates the pre-render window (happy-dom upgrades). */
  #query<T extends Element>(selector: string): T | null {
    return this.shadowRoot?.querySelector<T>(selector) ?? null;
  }

  #queryAll<T extends Element>(selector: string): T[] {
    return [...(this.shadowRoot?.querySelectorAll<T>(selector) ?? [])];
  }

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

  /** Prefill from the selection's current effects and open. */
  show(current: TextEffects = {}): void {
    const field = (name: string): HTMLInputElement | HTMLSelectElement | null =>
      this.#query<HTMLInputElement | HTMLSelectElement>(`[data-field="${name}"]`);
    const check = (name: string, on: boolean): void => {
      const input = field(name) as HTMLInputElement | null;
      if (input) input.checked = on;
    };
    const number = (name: string, value: number | undefined, fallback: number): void => {
      const input = field(name) as HTMLInputElement | null;
      if (input)
        input.value = String(Math.round(((value ?? fallback) + Number.EPSILON) * 100) / 100);
    };
    const color = (name: string, hex: string): void => {
      const input = field(name) as HTMLInputElement | null;
      if (input) input.value = `#${hex}`;
    };

    check("outline-on", !!current.outline);
    color("outline-color", current.outline?.color ?? "000000");
    number("outline-width", current.outline ? current.outline.widthPx / (96 / 72) : undefined, 1);
    {
      const dash = field("outline-dash") as HTMLSelectElement | null;
      if (dash) dash.value = current.outline?.dash ?? "solid";
    }

    check("shadow-on", !!current.shadow);
    color("shadow-color", current.shadow?.color ?? "000000");
    number(
      "shadow-transparency",
      current.shadow ? (1 - current.shadow.opacity) * 100 : undefined,
      40,
    );
    number("shadow-blur", current.shadow?.blurPx, 2);
    number("shadow-distance", current.shadow?.distPx, 3);
    number("shadow-angle", current.shadow?.dirDeg, 45);

    check("reflection-on", !!current.reflection);
    number(
      "reflection-transparency",
      current.reflection ? (1 - current.reflection.startOpacity) * 100 : undefined,
      50,
    );
    number("reflection-blur", current.reflection?.blurPx, 0.5);
    number("reflection-distance", current.reflection?.distPx, 0);

    check("glow-on", !!current.glow);
    color("glow-color", current.glow?.color ?? "4472C4");
    number("glow-transparency", current.glow ? (1 - current.glow.opacity) * 100 : undefined, 40);
    number("glow-size", current.glow?.radiusPx, 5);

    check("bevel-on", !!current.bevel?.top);
    {
      const preset = field("bevel-preset") as HTMLSelectElement | null;
      if (preset) preset.value = current.bevel?.top?.preset ?? "circle";
    }
    number("bevel-width", current.bevel?.top?.widthPx, 5.5);
    number("bevel-height", current.bevel?.top?.heightPx, 5.5);

    check("rotation-on", !!current.rotation);
    number("rotation-x", current.rotation?.x, 0);
    number("rotation-y", current.rotation?.y, 0);
    number("rotation-z", current.rotation?.z, 0);

    this.dialogEl?.show();
  }

  hide(): void {
    this.dialogEl?.hide();
  }

  /** Build the per-family patch and emit it for the host to apply. */
  submit(): void {
    const value = (name: string): string => {
      const input = this.#query<HTMLInputElement | HTMLSelectElement>(`[data-field="${name}"]`);
      return input?.value ?? "";
    };
    const checked = (name: string): boolean =>
      this.#query<HTMLInputElement>(`[data-field="${name}"]`)?.checked === true;
    const num = (name: string, fallback = 0): number => {
      const parsed = Number(value(name));
      return Number.isFinite(parsed) ? parsed : fallback;
    };

    const patch: Record<string, unknown> = {};
    patch.outline = checked("outline-on")
      ? {
          color: hexOf(value("outline-color")),
          widthPx: ptToPx(num("outline-width", 1)),
          dash: value("outline-dash") || "solid",
        }
      : null;
    patch.shadow = checked("shadow-on")
      ? {
          color: hexOf(value("shadow-color")),
          opacity: 1 - num("shadow-transparency") / 100,
          blurPx: num("shadow-blur"),
          distPx: num("shadow-distance"),
          dirDeg: num("shadow-angle", 45),
        }
      : null;
    patch.reflection = checked("reflection-on")
      ? {
          blurPx: num("reflection-blur"),
          startOpacity: 1 - num("reflection-transparency") / 100,
          endOpacity: (1 - num("reflection-transparency") / 100) * 0.15,
          distPx: num("reflection-distance"),
        }
      : null;
    patch.glow = checked("glow-on")
      ? {
          color: hexOf(value("glow-color")),
          opacity: 1 - num("glow-transparency") / 100,
          radiusPx: ptToPx(num("glow-size", 5)),
        }
      : null;
    patch.bevel = checked("bevel-on")
      ? {
          top: {
            widthPx: num("bevel-width", 5.5),
            heightPx: num("bevel-height", 5.5),
            preset: value("bevel-preset") || "circle",
          },
        }
      : null;
    patch.rotation = checked("rotation-on")
      ? { x: num("rotation-x"), y: num("rotation-y"), z: num("rotation-z") }
      : null;

    this.dialogEl?.hide();
    this.$emit("text-effects:ok", patch);
  }

  #applyLabels(): void {
    if (this.dialogEl) this.dialogEl.heading = t("textEffects.title", this);
    if (this.okBtn) this.okBtn.textContent = t("options.ok", this);
    if (this.cancelBtn) this.cancelBtn.textContent = t("options.cancel", this);
    for (const el of this.#queryAll<HTMLElement>("[data-i18n]")) {
      const key = el.getAttribute("data-i18n");
      if (key) el.textContent = t(key, this);
    }
  }
}

export default DocenTextEffectsDialog;
