import type { Editor } from "@docen/docx/core";

import type { DrawingHit } from "../../drawing";

/**
 * Home → Editing → Select → Select Objects (and the Draw tab's Select button).
 *
 * Word's object-selection mode: clicks select a floating object instead of
 * placing a caret, a drag on empty canvas marquees a set of objects, Ctrl+click
 * toggles membership, Delete removes the selected objects, and Esc/empty click
 * deselects (and exits the mode once nothing is selected). This class owns the
 * mode state and its overlay (marquee + per-object frames) and drives the
 * document through the host's editor/node resolvers — the bridge keeps only the
 * pointer/key routing.
 */

/** The hit-test surface the mode needs from the bridge/stage. */
export interface ObjectSelectionHost {
  /** The topmost painted drawing at a page-local point (null = empty canvas). */
  drawingAt(page: number, lx: number, ly: number): DrawingHit | null;
  /** Every painted drawing box across pages (the marquee's candidate set). */
  allDrawingBoxes(): DrawingHit[];
  /** Re-resolve a hit against fresh geometry after a re-render (null = gone). */
  resolveBox(hit: DrawingHit): DrawingHit | null;
  /** The PM node position for a hit (null when the hit cannot pair). */
  nodePosOf(hit: DrawingHit): number | null;
  /** The page frame element (overlay geometry + marquee coordinate space). */
  pageHost(page: number): HTMLElement | null;
  /** The overlay host element the frames/marquee mount in. */
  overlayHost(): HTMLElement | null;
  /** Mode-state notifier: the host mirrors it (ribbon lit state, cursor). */
  onChange?(active: boolean): void;
  scale(): number;
  editor(): Editor;
}

/** True when two hits address the same painted drawing. */
function sameHit(a: DrawingHit, b: DrawingHit): boolean {
  return (
    a.para === b.para &&
    a.index === b.index &&
    a.kind === b.kind &&
    samePath(a.childPath, b.childPath)
  );
}

function samePath(a: readonly number[] | undefined, b: readonly number[] | undefined): boolean {
  if (a === b) return true;
  if (!a || !b || a.length !== b.length) return false;
  return a.every((v, i) => v === b[i]);
}

/** A press handled by the mode — client coords plus the page-local hit point. */
export interface ObjectSelectionPress {
  page: number;
  lx: number;
  ly: number;
  clientX: number;
  clientY: number;
  ctrlKey: boolean;
  metaKey: boolean;
}

export class ObjectSelectionMode {
  readonly #host: ObjectSelectionHost;
  /** The absolutely-positioned overlay (marquee + selection frames). */
  readonly el: HTMLDivElement;
  #active = false;
  /** The selected drawings (deduped; the last press's target first is not
   *  significant — every member renders the same frame). */
  #items: DrawingHit[] = [];
  #frames: HTMLDivElement[] = [];
  #marquee: {
    sx: number;
    sy: number;
    cx: number;
    cy: number;
    moved: boolean;
    additive: boolean;
  } | null = null;

  constructor(host: ObjectSelectionHost) {
    this.#host = host;
    this.el = document.createElement("div");
    this.el.className = "docen-object-selection";
    Object.assign(this.el.style, {
      position: "absolute",
      inset: "0",
      pointerEvents: "none",
      zIndex: "9",
      display: "none",
    } satisfies Partial<CSSStyleDeclaration>);
  }

  /** Mount the overlay into the bridge's input layer. */
  mount(hostEl: HTMLElement): void {
    hostEl.append(this.el);
  }

  get active(): boolean {
    return this.#active;
  }

  get count(): number {
    return this.#items.length;
  }

  /** Painted overlay elements (selection frames + a live marquee). */
  get frameCount(): number {
    return this.#frames.length;
  }

  get dragging(): boolean {
    return this.#marquee?.moved === true;
  }

  /** A marquee press is armed (the pointer is down on empty canvas). */
  get pressed(): boolean {
    return this.#marquee != null;
  }

  /** Toggle the mode. Turning it off clears the selection (Word's button). */
  setActive(on: boolean): void {
    if (on === this.#active) return;
    this.#active = on;
    if (!on) {
      this.#items = [];
      this.#cancelMarquee();
    }
    this.place();
    this.#host.onChange?.(on);
  }

  clearSelection(): void {
    if (!this.#items.length && !this.#marquee) return;
    this.#items = [];
    this.#cancelMarquee();
    this.place();
  }

  /** A left press while the mode is active. Always consumed: it selects a hit,
   *  toggles with Ctrl, or arms the marquee. Returns false only when the mode
   *  is off (the caller then keeps its normal chains). */
  press(press: ObjectSelectionPress): boolean {
    if (!this.#active) return false;
    const hit = this.#host.drawingAt(press.page, press.lx, press.ly);
    if (hit && this.#host.nodePosOf(hit) != null) {
      if (press.ctrlKey || press.metaKey) {
        const at = this.#items.findIndex((m) => sameHit(m, hit));
        if (at >= 0) this.#items.splice(at, 1);
        else this.#items.push(hit);
      } else {
        this.#items = [hit];
      }
      this.place();
      return true;
    }
    const hostRect = this.#host.overlayHost()?.getBoundingClientRect();
    const left = hostRect?.left ?? 0;
    const top = hostRect?.top ?? 0;
    this.#marquee = {
      sx: press.clientX - left,
      sy: press.clientY - top,
      cx: press.clientX - left,
      cy: press.clientY - top,
      moved: false,
      additive: press.ctrlKey || press.metaKey,
    };
    return true;
  }

  /** Pointer move while a marquee press is armed. */
  move(clientX: number, clientY: number): void {
    const marquee = this.#marquee;
    if (!marquee) return;
    const hostRect = this.#host.overlayHost()?.getBoundingClientRect();
    marquee.cx = clientX - (hostRect?.left ?? 0);
    marquee.cy = clientY - (hostRect?.top ?? 0);
    if (!marquee.moved && Math.hypot(marquee.cx - marquee.sx, marquee.cy - marquee.sy) >= 3) {
      marquee.moved = true;
    }
    this.place();
  }

  /** Pointer release: commit the marquee selection, or treat a non-moved press
   *  as an empty click — clear the selection, or leave the mode when nothing
   *  was selected (Word's Esc/empty-click semantics). */
  release(): boolean {
    const marquee = this.#marquee;
    this.#marquee = null;
    if (!marquee) return false;
    if (!marquee.moved) {
      if (this.#items.length) this.clearSelection();
      else this.setActive(false);
      return true;
    }
    const picked = this.#marqueeHits(marquee);
    this.#items = marquee.additive ? dedupe([...this.#items, ...picked]) : picked;
    this.place();
    return true;
  }

  /** Delete/Backspace on the selection — remove every selected object in one
   *  transaction (descending position order keeps earlier positions valid).
   *  False when there is nothing to delete. */
  deleteSelection(): boolean {
    if (!this.#active || !this.#items.length) return false;
    const editor = this.#host.editor();
    const positions = this.#items
      .map((hit) => this.#host.nodePosOf(hit))
      .filter((pos): pos is number => pos != null)
      .sort((a, b) => b - a);
    if (!positions.length) return false;
    const tr = editor.state.tr;
    let deleted = false;
    for (const pos of positions) {
      const node = editor.state.doc.nodeAt(pos);
      if (!node) continue;
      tr.delete(pos, pos + node.nodeSize);
      deleted = true;
    }
    if (!deleted) return false;
    editor.view.dispatch(tr);
    this.clearSelection();
    return true;
  }

  /** Esc's stepdown: cancel a live marquee, then clear the selection, then
   *  leave the mode. False when there was nothing to cancel. */
  escape(): boolean {
    if (this.#marquee) {
      this.#cancelMarquee();
      this.place();
      return true;
    }
    if (this.#items.length) {
      this.clearSelection();
      return true;
    }
    if (this.#active) {
      this.setActive(false);
      return true;
    }
    return false;
  }

  /** Drop selections that no longer paint (a re-render re-objects the laid
   *  paragraphs, so every hit re-resolves by identity). */
  refresh(): void {
    if (!this.#items.length) return;
    this.#items = this.#items
      .map((hit) => this.#host.resolveBox(hit))
      .filter((hit): hit is DrawingHit => hit != null);
    this.place();
  }

  destroy(): void {
    this.el.remove();
  }

  /** Redraw the overlay: one frame per selected object (mounted inside its
   *  page frame, so it scrolls with the page like the drawing overlay), plus
   *  the viewport-space marquee. */
  place(): void {
    for (const frame of this.#frames) frame.remove();
    this.#frames = [];
    if (!this.#active) {
      this.el.style.display = "none";
      return;
    }
    const scale = this.#host.scale();
    for (const hit of this.#items) {
      const frameHost = this.#host.pageHost(hit.page);
      if (!frameHost) continue;
      const el = document.createElement("div");
      el.className = "docen-object-selection-frame";
      Object.assign(el.style, {
        position: "absolute",
        left: `${hit.x * scale}px`,
        top: `${hit.y * scale}px`,
        width: `${Math.max(1, hit.width * scale)}px`,
        height: `${Math.max(1, hit.height * scale)}px`,
        border: "1.5px solid var(--docen-theme-accent1, #2b579a)",
        background: "rgba(43, 87, 154, 0.08)",
        pointerEvents: "none",
        boxSizing: "border-box",
        zIndex: "9",
      } satisfies Partial<CSSStyleDeclaration>);
      frameHost.append(el);
      this.#frames.push(el);
    }
    const marquee = this.#marquee;
    if (marquee?.moved) {
      const left = Math.min(marquee.sx, marquee.cx);
      const top = Math.min(marquee.sy, marquee.cy);
      const box = document.createElement("div");
      box.className = "docen-object-selection-marquee";
      Object.assign(box.style, {
        position: "absolute",
        left: `${left}px`,
        top: `${top}px`,
        width: `${Math.abs(marquee.cx - marquee.sx)}px`,
        height: `${Math.abs(marquee.cy - marquee.sy)}px`,
        border: "1px dashed var(--docen-theme-accent1, #2b579a)",
        background: "rgba(43, 87, 154, 0.10)",
        pointerEvents: "none",
        boxSizing: "border-box",
      } satisfies Partial<CSSStyleDeclaration>);
      this.el.append(box);
      this.#frames.push(box);
    }
    // The shared overlay hosts only the viewport-space marquee; the selection
    // frames ride their page hosts.
    this.el.style.display = marquee?.moved ? "block" : "none";
  }

  /** The floating drawings whose painted boxes intersect the marquee. The
   *  marquee lives in host coordinates; each page frame converts a slice of it
   *  into the page-local px the boxes carry. Inline art rides the text stream
   *  and is never marquee-selected (Word selects it only by direct click). */
  #marqueeHits(marquee: { sx: number; sy: number; cx: number; cy: number }): DrawingHit[] {
    const host = this.#host.overlayHost();
    const hostRect = host?.getBoundingClientRect();
    if (!hostRect) return [];
    const scale = this.#host.scale() || 1;
    const left = hostRect.left + Math.min(marquee.sx, marquee.cx);
    const top = hostRect.top + Math.min(marquee.sy, marquee.cy);
    const right = hostRect.left + Math.max(marquee.sx, marquee.cx);
    const bottom = hostRect.top + Math.max(marquee.sy, marquee.cy);
    const byPage = new Map<number, DrawingHit[]>();
    for (const box of this.#host.allDrawingBoxes()) {
      if (box.kind !== "drawing" || box.chartPart) continue;
      const list = byPage.get(box.page);
      if (list) list.push(box);
      else byPage.set(box.page, [box]);
    }
    const picked: DrawingHit[] = [];
    for (const [page, boxes] of byPage) {
      const frame = this.#host.pageHost(page);
      const rect = frame?.getBoundingClientRect();
      if (!rect) continue;
      const x0 = (left - rect.left) / scale;
      const y0 = (top - rect.top) / scale;
      const x1 = (right - rect.left) / scale;
      const y1 = (bottom - rect.top) / scale;
      for (const box of boxes) {
        if (box.x < x1 && box.x + box.width > x0 && box.y < y1 && box.y + box.height > y0) {
          picked.push(box);
        }
      }
    }
    return picked;
  }

  #cancelMarquee(): void {
    this.#marquee = null;
  }
}

/** Dedupe hits by identity (a marquee may cross a grouped box's members). */
function dedupe(hits: DrawingHit[]): DrawingHit[] {
  const out: DrawingHit[] = [];
  for (const hit of hits) {
    if (!out.some((other) => sameHit(other, hit))) out.push(hit);
  }
  return out;
}
