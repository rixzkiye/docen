export interface ShapingWasmExports {
  readonly memory: WebAssembly.Memory;
  alloc(size: number): number;
  dealloc(ptr: number, size: number): void;
  register_font(ptr: number, len: number): number;
  register_font_index(ptr: number, len: number, index: number): number;
  drop_font(fontId: number): number;
  get_shape_buffer_ptr(): number;
  shape_text(
    fontId: number,
    textPtr: number,
    textLen: number,
    direction: number,
    scriptTag: number,
    langPtr: number,
    langLen: number,
    featuresPtr: number,
    featuresCount: number,
    variationsPtr: number,
    variationsCount: number,
  ): number;
  get_metrics_buffer_ptr(): number;
  get_font_metrics(fontId: number): number;
  get_font_metrics_var(fontId: number, variationsPtr: number, variationsCount: number): number;
  get_outline_buffer_ptr(): number;
  get_glyph_outline(fontId: number, glyphId: number): number;
  get_glyph_outline_var(
    fontId: number,
    glyphId: number,
    variationsPtr: number,
    variationsCount: number,
  ): number;
  get_font_axes_ptr(): number;
  get_font_axes(fontId: number): number;
  get_string_buffer_ptr(): number;
  get_font_name(fontId: number, nameId: number): number;
  get_glyph_count(fontId: number): number;
  get_font_fs_type(fontId: number): number;
}

let wasmInstance: WebAssembly.Instance | null = null;
let wasmExports: ShapingWasmExports | null = null;

/**
 * Initialize the docen-shaping WebAssembly module.
 * Can be provided an ArrayBuffer/Uint8Array, or defaults to reading/fetching the vendored wasm file.
 */
export async function initShapingWasm(
  wasmInput?: ArrayBuffer | Uint8Array,
): Promise<ShapingWasmExports> {
  if (wasmExports) {
    return wasmExports;
  }

  let bytes: ArrayBuffer | Uint8Array;
  if (wasmInput) {
    bytes = wasmInput;
  } else if (
    typeof window !== "undefined" ||
    typeof document !== "undefined" ||
    !(globalThis as unknown as { process?: { versions?: { node?: string } } }).process?.versions
      ?.node
  ) {
    const res = await fetch(new URL("../wasm/docen_shaping.wasm", import.meta.url));
    bytes = await res.arrayBuffer();
  } else {
    // Dynamic import to prevent browser bundlers (Vite/webpack) from failing on Node builtins
    const { readFileSync } = await import("node:fs");
    const { fileURLToPath } = await import("node:url");
    const wasmPath = fileURLToPath(new URL("../wasm/docen_shaping.wasm", import.meta.url));
    bytes = readFileSync(wasmPath);
  }

  const module = await WebAssembly.compile(bytes as BufferSource);
  wasmInstance = await WebAssembly.instantiate(module, {});
  wasmExports = wasmInstance.exports as unknown as ShapingWasmExports;
  return wasmExports;
}

export function getShapingWasm(): ShapingWasmExports {
  if (!wasmExports) {
    throw new Error("[@docen/shaping] WASM module not initialized. Call initShapingWasm() first.");
  }
  return wasmExports;
}

export function isShapingWasmInitialized(): boolean {
  return wasmExports !== null;
}
