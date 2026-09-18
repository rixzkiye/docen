import type { Box } from "./geometry";

export interface VertexPoint {
  pathIndex: number;
  commandIndex: number;
  pointIndex?: number;
  x: number; // In path coordinates
  y: number;
}

export interface NodeEditOverlayCallbacks {
  scale(): number;
  applyCustomGeometry(customGeometry: Record<string, unknown>): void;
  onExit?(): void;
}

/**
 * NodeEditOverlay — "Edit Points" vertex handle overlay for shapes with customGeometry.
 * Renders vertex handles (small black squares with white borders) along the path outline
 * and supports live dragging of vertices to edit the shape geometry.
 */
export class NodeEditOverlay {
  readonly el: HTMLDivElement;
  #callbacks: NodeEditOverlayCallbacks;
  #box: Box | null = null;
  #rotation = 0;
  #cg: Record<string, unknown> | null = null;
  #svgEl: SVGSVGElement;
  #pathEl: SVGPathElement;
  #handlesContainer: HTMLDivElement;
  #activeDrag: {
    vertex: VertexPoint;
    startPointerX: number;
    startPointerY: number;
    initialVertexX: number;
    initialVertexY: number;
  } | null = null;

  #onDocDown = (e: PointerEvent): void => this.#onDocumentPointerDown(e);
  #onDocKey = (e: KeyboardEvent): void => this.#onDocumentKeyDown(e);

  constructor(callbacks: NodeEditOverlayCallbacks) {
    this.#callbacks = callbacks;
    this.el = document.createElement("div");
    this.el.setAttribute("data-docen-overlay", "");
    this.el.setAttribute("data-node-edit", "");
    Object.assign(this.el.style, {
      position: "absolute",
      pointerEvents: "none",
      zIndex: "7",
      display: "none",
      overflow: "visible",
    } satisfies Partial<CSSStyleDeclaration>);

    // SVG contour outline
    this.#svgEl = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    this.#svgEl.setAttribute("overflow", "visible");
    Object.assign(this.#svgEl.style, {
      position: "absolute",
      left: "0",
      top: "0",
      width: "100%",
      height: "100%",
      pointerEvents: "none",
    } satisfies Partial<CSSStyleDeclaration>);

    this.#pathEl = document.createElementNS("http://www.w3.org/2000/svg", "path");
    this.#pathEl.setAttribute("fill", "none");
    this.#pathEl.setAttribute("stroke", "#D83B01"); // Word's red contour outline
    this.#pathEl.setAttribute("stroke-width", "1.5");
    this.#svgEl.appendChild(this.#pathEl);
    this.el.appendChild(this.#svgEl);

    // Handles container
    this.#handlesContainer = document.createElement("div");
    Object.assign(this.#handlesContainer.style, {
      position: "absolute",
      inset: "0",
      pointerEvents: "none",
    } satisfies Partial<CSSStyleDeclaration>);
    this.el.appendChild(this.#handlesContainer);
  }

  get active(): boolean {
    return this.#box != null;
  }

  show(box: Box, rotation: number, customGeometry: Record<string, unknown>): void {
    this.#box = box;
    this.#rotation = rotation;
    this.#cg = JSON.parse(JSON.stringify(customGeometry));
    this.el.style.display = "block";
    this.#render();
    document.addEventListener("pointerdown", this.#onDocDown, true);
    document.addEventListener("keydown", this.#onDocKey, true);
  }

  hide(): void {
    if (this.#box == null) return;
    this.#box = null;
    this.#cg = null;
    this.#activeDrag = null;
    this.el.style.display = "none";
    this.#handlesContainer.innerHTML = "";
    document.removeEventListener("pointerdown", this.#onDocDown, true);
    document.removeEventListener("keydown", this.#onDocKey, true);
    this.#callbacks.onExit?.();
  }

  commit(): void {
    if (this.#cg) {
      this.#callbacks.applyCustomGeometry(this.#cg);
    }
    this.hide();
  }

  cancel(): void {
    this.hide();
  }

  #collectVertices(): VertexPoint[] {
    const list: VertexPoint[] = [];
    if (!this.#cg || !Array.isArray(this.#cg.pathList)) return list;

    this.#cg.pathList.forEach((path: any, pathIndex: number) => {
      if (!Array.isArray(path.commands)) return;
      path.commands.forEach((cmd: any, commandIndex: number) => {
        const type = String(cmd.command).toLowerCase();
        if (type === "moveto" || type === "lineto" || type === "lnto") {
          if (cmd.point) {
            list.push({
              pathIndex,
              commandIndex,
              x: Number(cmd.point.x) || 0,
              y: Number(cmd.point.y) || 0,
            });
          }
        } else if (type === "quadbezto" || type === "cubicbezto") {
          if (Array.isArray(cmd.points)) {
            cmd.points.forEach((pt: any, pointIndex: number) => {
              list.push({
                pathIndex,
                commandIndex,
                pointIndex,
                x: Number(pt.x) || 0,
                y: Number(pt.y) || 0,
              });
            });
          }
        }
      });
    });

    return list;
  }

  #render(): void {
    if (!this.#box || !this.#cg) return;
    const scale = this.#callbacks.scale();
    const box = this.#box;

    Object.assign(this.el.style, {
      left: `${box.x * scale}px`,
      top: `${box.y * scale}px`,
      width: `${box.width * scale}px`,
      height: `${box.height * scale}px`,
      transform: this.#rotation ? `rotate(${this.#rotation}deg)` : "none",
      transformOrigin: "center center",
    } satisfies Partial<CSSStyleDeclaration>);

    const pathList = Array.isArray(this.#cg.pathList) ? this.#cg.pathList : [];
    const firstPath = pathList[0] || {};
    const pw = Number(firstPath.w) || box.width;
    const ph = Number(firstPath.h) || box.height;
    const sx = (box.width * scale) / pw;
    const sy = (box.height * scale) / ph;

    // Build SVG path string for preview
    let d = "";
    if (firstPath && Array.isArray(firstPath.commands)) {
      for (const cmd of firstPath.commands) {
        const type = String(cmd.command).toLowerCase();
        if (type === "moveto" && cmd.point) {
          d += `M ${Number(cmd.point.x) * sx} ${Number(cmd.point.y) * sy} `;
        } else if ((type === "lineto" || type === "lnto") && cmd.point) {
          d += `L ${Number(cmd.point.x) * sx} ${Number(cmd.point.y) * sy} `;
        } else if (type === "quadbezto" && Array.isArray(cmd.points) && cmd.points.length >= 2) {
          d += `Q ${Number(cmd.points[0].x) * sx} ${Number(cmd.points[0].y) * sy} ${Number(cmd.points[1].x) * sx} ${Number(cmd.points[1].y) * sy} `;
        } else if (type === "cubicbezto" && Array.isArray(cmd.points) && cmd.points.length >= 3) {
          d += `C ${Number(cmd.points[0].x) * sx} ${Number(cmd.points[0].y) * sy} ${Number(cmd.points[1].x) * sx} ${Number(cmd.points[1].y) * sy} ${Number(cmd.points[2].x) * sx} ${Number(cmd.points[2].y) * sy} `;
        } else if (type === "close") {
          d += "Z ";
        }
      }
    }
    this.#pathEl.setAttribute("d", d.trim());

    // Render vertex handles
    this.#handlesContainer.innerHTML = "";
    const vertices = this.#collectVertices();

    vertices.forEach((v, idx) => {
      const handle = document.createElement("div");
      handle.setAttribute("data-vertex-handle", String(idx));
      const px = v.x * sx;
      const py = v.y * sy;

      Object.assign(handle.style, {
        position: "absolute",
        left: `${px - 4}px`,
        top: `${py - 4}px`,
        width: "8px",
        height: "8px",
        background: "#000",
        border: "1px solid #fff",
        boxSizing: "border-box",
        cursor: "crosshair",
        pointerEvents: "auto",
        zIndex: "10",
      } satisfies Partial<CSSStyleDeclaration>);

      handle.addEventListener("pointerdown", (e) => this.#onVertexPointerDown(v, e));
      this.#handlesContainer.appendChild(handle);
    });
  }

  #onVertexPointerDown(v: VertexPoint, e: PointerEvent): void {
    e.stopPropagation();
    e.preventDefault();
    this.#activeDrag = {
      vertex: v,
      startPointerX: e.clientX,
      startPointerY: e.clientY,
      initialVertexX: v.x,
      initialVertexY: v.y,
    };

    const onMove = (moveEv: PointerEvent): void => {
      if (!this.#activeDrag || !this.#box || !this.#cg) return;
      const scale = this.#callbacks.scale();
      const pathList = (this.#cg.pathList as any[]) || [];
      const path = pathList[this.#activeDrag.vertex.pathIndex] || {};
      const pw = Number(path.w) || this.#box.width;
      const ph = Number(path.h) || this.#box.height;
      const sx = (this.#box.width * scale) / pw;
      const sy = (this.#box.height * scale) / ph;

      const dx = (moveEv.clientX - this.#activeDrag.startPointerX) / sx;
      const dy = (moveEv.clientY - this.#activeDrag.startPointerY) / sy;

      const newX = Math.round(this.#activeDrag.initialVertexX + dx);
      const newY = Math.round(this.#activeDrag.initialVertexY + dy);

      // Mutate vertex in memory
      const cmd = path.commands[this.#activeDrag.vertex.commandIndex];
      if (this.#activeDrag.vertex.pointIndex != null) {
        cmd.points[this.#activeDrag.vertex.pointIndex].x = String(newX);
        cmd.points[this.#activeDrag.vertex.pointIndex].y = String(newY);
      } else {
        cmd.point.x = String(newX);
        cmd.point.y = String(newY);
      }

      this.#render();
    };

    const onUp = (_upEv: PointerEvent): void => {
      window.removeEventListener("pointermove", onMove, true);
      window.removeEventListener("pointerup", onUp, true);
      if (this.#activeDrag) {
        this.#activeDrag = null;
        if (this.#cg) {
          this.#callbacks.applyCustomGeometry(this.#cg);
        }
      }
    };

    window.addEventListener("pointermove", onMove, true);
    window.addEventListener("pointerup", onUp, true);
  }

  #onDocumentPointerDown(e: PointerEvent): void {
    if (this.el.contains(e.target as Node)) return;
    this.commit();
  }

  #onDocumentKeyDown(e: KeyboardEvent): void {
    if (e.key === "Escape") {
      e.preventDefault();
      this.cancel();
    } else if (e.key === "Enter") {
      e.preventDefault();
      this.commit();
    }
  }
}
