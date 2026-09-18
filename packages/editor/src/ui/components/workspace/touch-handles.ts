import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import type { EditBridge } from "../../../document/canvas/edit-bridge";
import { observeLang, resolveDir } from "../../i18n/localize";

export type TouchHandleKind = "start" | "end";

export interface TouchHandleDragDetail {
  handle: TouchHandleKind;
  clientX: number;
  clientY: number;
}

const styles = css`
  :host {
    position: absolute;
    inset: 0;
    pointer-events: none;
    z-index: 1000;
    overflow: visible;
  }

  :host([hidden]) {
    display: none !important;
  }

  .handle {
    position: absolute;
    pointer-events: auto;
    touch-action: none;
    user-select: none;
    display: flex;
    flex-direction: column;
    align-items: center;
    cursor: pointer;
    z-index: 1001;
    /* 40x40px minimum touch hit area centered on caret point */
    width: 40px;
    height: 48px;
    margin-left: -20px;
    justify-content: flex-start;
  }

  .stem {
    width: 2px;
    height: 12px;
    background-color: var(--docen-selection-handle-color, #0f6cbd);
    border-radius: 1px;
    flex-shrink: 0;
  }

  .teardrop {
    width: 22px;
    height: 22px;
    background-color: var(--docen-selection-handle-color, #0f6cbd);
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.35);
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    cursor: grab;
    transition:
      transform 0.1s ease,
      background-color 0.15s ease;
  }

  .teardrop:active,
  .handle.dragging .teardrop {
    cursor: grabbing;
    transform: scale(1.18);
    background-color: var(--docen-selection-handle-active-color, #115ea3);
  }

  /* Teardrop geometry: start points top-right, end points top-left */
  .teardrop-start {
    border-radius: 0 50% 50% 50%;
    transform: rotate(45deg);
  }

  .teardrop-start:active,
  .handle.dragging .teardrop-start {
    transform: rotate(45deg) scale(1.18);
  }

  .teardrop-end {
    border-radius: 50% 0 50% 50%;
    transform: rotate(-45deg);
  }

  .teardrop-end:active,
  .handle.dragging .teardrop-end {
    transform: rotate(-45deg) scale(1.18);
  }

  :host([dir="rtl"]) .teardrop-start {
    border-radius: 50% 0 50% 50%;
    transform: rotate(-45deg);
  }

  :host([dir="rtl"]) .teardrop-start:active,
  :host([dir="rtl"]) .handle.dragging .teardrop-start {
    transform: rotate(-45deg) scale(1.18);
  }

  :host([dir="rtl"]) .teardrop-end {
    border-radius: 0 50% 50% 50%;
    transform: rotate(45deg);
  }

  :host([dir="rtl"]) .teardrop-end:active,
  :host([dir="rtl"]) .handle.dragging .teardrop-end {
    transform: rotate(45deg) scale(1.18);
  }

  .inner-dot {
    width: 6px;
    height: 6px;
    background-color: #ffffff;
    border-radius: 50%;
    pointer-events: none;
  }
`;

const template = html<DocenTouchHandles>`
  <div
    class="handle handle-start ${(x) => (x.activeHandle === "start" ? "dragging" : "")}"
    part="handle-start"
    ${ref("startHandle")}
    style="left: ${(x) => x.startX}px; top: ${(x) => x.startY + x.startHeight}px; display: ${(x) =>
      x.visible && x.hasSelection ? "flex" : "none"};"
    @pointerdown="${(x, c) => x.onHandlePointerDown("start", c.event as PointerEvent)}"
  >
    <div class="stem" part="stem-start"></div>
    <div class="teardrop teardrop-start" part="teardrop-start">
      <div class="inner-dot"></div>
    </div>
  </div>

  <div
    class="handle handle-end ${(x) => (x.activeHandle === "end" ? "dragging" : "")}"
    part="handle-end"
    ${ref("endHandle")}
    style="left: ${(x) => x.endX}px; top: ${(x) => x.endY + x.endHeight}px; display: ${(x) =>
      x.visible && x.hasSelection ? "flex" : "none"};"
    @pointerdown="${(x, c) => x.onHandlePointerDown("end", c.event as PointerEvent)}"
  >
    <div class="stem" part="stem-end"></div>
    <div class="teardrop teardrop-end" part="teardrop-end">
      <div class="inner-dot"></div>
    </div>
  </div>
`;

/**
 * `<docen-touch-handles>`
 * Selection handles for touch/pen input:
 * - Start & end teardrop handles with ≥40x40px hit area
 * - Draggable handles moving selection anchor / head
 * - RTL aware orientation
 */
@customElement({ name: "docen-touch-handles", template, styles })
export class DocenTouchHandles extends FASTElement {
  @attr({ mode: "boolean" }) visible = false;
  @attr dir: "ltr" | "rtl" = "ltr";

  @observable startX = 0;
  @observable startY = 0;
  @observable startHeight = 20;

  @observable endX = 0;
  @observable endY = 0;
  @observable endHeight = 20;

  @observable hasSelection = false;
  @observable isDragging = false;
  @observable activeHandle: TouchHandleKind | null = null;

  @observable startHandle?: HTMLElement;
  @observable endHandle?: HTMLElement;

  #unobserveLang?: () => void;
  #activePointerId: number | null = null;

  get isRtl(): boolean {
    return this.dir === "rtl" || resolveDir(this) === "rtl";
  }

  dirChanged(): void {
    this.#syncDir();
  }

  #syncDir = (): void => {
    const nextDir = resolveDir(this);
    if (this.getAttribute("dir") !== nextDir) {
      this.setAttribute("dir", nextDir);
    }
  };

  override connectedCallback(): void {
    super.connectedCallback();
    this.#syncDir();
    this.#unobserveLang = observeLang(() => this.#syncDir());
  }

  override disconnectedCallback(): void {
    this.#unobserveLang?.();
    this.#cleanupPointerDrag();
    super.disconnectedCallback();
  }

  show(): void {
    this.visible = true;
  }

  hide(): void {
    this.visible = false;
    this.hasSelection = false;
  }

  setPositions(
    start: { x: number; y: number; height?: number },
    end: { x: number; y: number; height?: number },
  ): void {
    this.startX = Math.round(start.x);
    this.startY = Math.round(start.y);
    this.startHeight = Math.round(start.height ?? 20);

    this.endX = Math.round(end.x);
    this.endY = Math.round(end.y);
    this.endHeight = Math.round(end.height ?? 20);

    this.hasSelection = true;
  }

  onHandlePointerDown(handle: TouchHandleKind, event: PointerEvent): void {
    event.preventDefault();
    event.stopPropagation();

    this.isDragging = true;
    this.activeHandle = handle;
    this.#activePointerId = event.pointerId;

    const target = event.currentTarget as HTMLElement;
    target.setPointerCapture(event.pointerId);

    target.addEventListener("pointermove", this.#onPointerMove);
    target.addEventListener("pointerup", this.#onPointerUp);
    target.addEventListener("pointercancel", this.#onPointerCancel);

    this.dispatchEvent(
      new CustomEvent<TouchHandleDragDetail>("handle-drag-start", {
        bubbles: true,
        composed: true,
        detail: {
          handle,
          clientX: event.clientX,
          clientY: event.clientY,
        },
      }),
    );
  }

  #onPointerMove = (event: PointerEvent): void => {
    if (!this.isDragging || !this.activeHandle || event.pointerId !== this.#activePointerId) return;
    event.preventDefault();
    event.stopPropagation();

    this.dispatchEvent(
      new CustomEvent<TouchHandleDragDetail>("handle-drag", {
        bubbles: true,
        composed: true,
        detail: {
          handle: this.activeHandle,
          clientX: event.clientX,
          clientY: event.clientY,
        },
      }),
    );
  };

  #onPointerUp = (event: PointerEvent): void => {
    if (event.pointerId !== this.#activePointerId) return;
    event.preventDefault();
    event.stopPropagation();

    const handle = this.activeHandle;
    this.#cleanupPointerDrag();

    if (handle) {
      this.dispatchEvent(
        new CustomEvent<TouchHandleDragDetail>("handle-drag-end", {
          bubbles: true,
          composed: true,
          detail: {
            handle,
            clientX: event.clientX,
            clientY: event.clientY,
          },
        }),
      );
    }
  };

  #onPointerCancel = (event: PointerEvent): void => {
    if (event.pointerId !== this.#activePointerId) return;
    this.#cleanupPointerDrag();
  };

  #cleanupPointerDrag(): void {
    if (this.startHandle && this.#activePointerId !== null) {
      try {
        this.startHandle.releasePointerCapture(this.#activePointerId);
      } catch {}
      this.startHandle.removeEventListener("pointermove", this.#onPointerMove);
      this.startHandle.removeEventListener("pointerup", this.#onPointerUp);
      this.startHandle.removeEventListener("pointercancel", this.#onPointerCancel);
    }
    if (this.endHandle && this.#activePointerId !== null) {
      try {
        this.endHandle.releasePointerCapture(this.#activePointerId);
      } catch {}
      this.endHandle.removeEventListener("pointermove", this.#onPointerMove);
      this.endHandle.removeEventListener("pointerup", this.#onPointerUp);
      this.endHandle.removeEventListener("pointercancel", this.#onPointerCancel);
    }
    this.isDragging = false;
    this.activeHandle = null;
    this.#activePointerId = null;
  }

  /**
   * Wire touch handles directly to an EditBridge instance.
   * Keeps handles synchronized with document selection and moves selection on drag.
   */
  attachToBridge(bridge: EditBridge, container: HTMLElement): () => void {
    const editor = bridge.activeEditor();

    const update = (): void => {
      if (this.isDragging) return;
      const { from, to } = editor.state.selection;
      if (from === to) {
        this.hide();
        return;
      }

      const rect = bridge.selectionClientRect(from, to);
      if (!rect) {
        this.hide();
        return;
      }

      const containerRect = container.getBoundingClientRect();
      this.setPositions(
        {
          x: rect.left - containerRect.left,
          y: rect.top - containerRect.top,
          height: rect.height,
        },
        {
          x: rect.right - containerRect.left,
          y: rect.top - containerRect.top,
          height: rect.height,
        },
      );
      this.show();
    };

    const onDrag = (e: Event): void => {
      const customEvent = e as CustomEvent<TouchHandleDragDetail>;
      const { handle, clientX, clientY } = customEvent.detail;
      const pos = bridge.posAtClient(clientX, clientY);
      if (pos == null) return;

      const { from, to } = editor.state.selection;
      if (handle === "start") {
        const nextAnchor = Math.min(pos, to);
        editor.commands.setTextSelection({ from: nextAnchor, to });
      } else {
        const nextHead = Math.max(pos, from);
        editor.commands.setTextSelection({ from, to: nextHead });
      }
    };

    editor.on("selectionUpdate", update);
    this.addEventListener("handle-drag", onDrag as EventListener);

    return () => {
      editor.off("selectionUpdate", update);
      this.removeEventListener("handle-drag", onDrag as EventListener);
      this.#cleanupPointerDrag();
    };
  }
}

export default DocenTouchHandles;
