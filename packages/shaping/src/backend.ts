import { CanvasBackend } from "./canvas-backend.js";
import { HarfbuzzBackend } from "./harfbuzz-backend.js";
import { RustybuzzBackend } from "./rustybuzz-backend.js";
import type { ShapingBackend } from "./types.js";

const backends = new Map<string, ShapingBackend>();
let defaultBackendId = "rustybuzz";

export function registerShapingBackend(backend: ShapingBackend): void {
  backends.set(backend.id, backend);
}

export function setDefaultShapingBackend(id: string): void {
  defaultBackendId = id;
}

export function getDefaultShapingBackendId(): string {
  return defaultBackendId;
}

export function hasShapingBackend(id: string): boolean {
  return backends.has(id) || id === "rustybuzz" || id === "harfbuzz" || id === "canvas";
}

export function getShapingBackend(id?: string): ShapingBackend {
  const targetId = id ?? defaultBackendId;
  const existing = backends.get(targetId);
  if (existing) {
    return existing;
  }

  if (targetId === "rustybuzz") {
    const b = new RustybuzzBackend();
    backends.set("rustybuzz", b);
    return b;
  }
  if (targetId === "harfbuzz") {
    const b = new HarfbuzzBackend();
    backends.set("harfbuzz", b);
    return b;
  }
  if (targetId === "canvas") {
    const b = new CanvasBackend();
    backends.set("canvas", b);
    return b;
  }

  throw new Error(`Shaping backend "${targetId}" is not registered.`);
}
