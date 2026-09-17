import { RustybuzzBackend } from "./rustybuzz-backend.js";
import type { ShapingOptions, ShapingResult } from "./types.js";
import { initShapingWasm } from "./wasm-loader.js";

export type WorkerRequest =
  | { readonly type: "init"; readonly wasmBytes?: ArrayBuffer }
  | { readonly type: "registerFont"; readonly fontId: number; readonly fontData: Uint8Array }
  | { readonly type: "dropFont"; readonly fontId: number }
  | {
      readonly type: "shapeBatch";
      readonly id: number;
      readonly items: readonly {
        readonly fontId: number;
        readonly text: string;
        readonly options?: ShapingOptions;
      }[];
    };

export type WorkerResponse =
  | { readonly type: "initDone"; readonly success: boolean; readonly error?: string }
  | { readonly type: "fontRegistered"; readonly fontId: number }
  | { readonly type: "fontDropped"; readonly fontId: number }
  | {
      readonly type: "batchResult";
      readonly id: number;
      readonly results: readonly ShapingResult[];
    }
  | { readonly type: "error"; readonly id?: number; readonly error: string };

/**
 * Worker handler implementation for execution inside Web Worker or Node worker_thread.
 */
export class ShapingWorkerHandler {
  private backend?: RustybuzzBackend;
  private fontMap = new Map<number, number>();
  private isInitialized = false;

  get initialized(): boolean {
    return this.isInitialized;
  }

  async handleMessage(
    message: WorkerRequest,
    postResponse: (response: WorkerResponse, transfer?: Transferable[]) => void,
  ): Promise<void> {
    try {
      switch (message.type) {
        case "init": {
          await initShapingWasm(message.wasmBytes);
          this.backend = new RustybuzzBackend();
          this.isInitialized = true;
          postResponse({ type: "initDone", success: true });
          break;
        }
        case "registerFont": {
          if (!this.backend) throw new Error("Worker not initialized");
          const wasmId = this.backend.registerFont(message.fontData);
          this.fontMap.set(message.fontId, wasmId);
          postResponse({ type: "fontRegistered", fontId: message.fontId });
          break;
        }
        case "dropFont": {
          if (this.backend) {
            const wasmId = this.fontMap.get(message.fontId);
            if (wasmId !== undefined) {
              this.backend.dropFont(wasmId);
              this.fontMap.delete(message.fontId);
            }
          }
          postResponse({ type: "fontDropped", fontId: message.fontId });
          break;
        }
        case "shapeBatch": {
          if (!this.backend) throw new Error("Worker not initialized");
          const results: ShapingResult[] = [];
          for (const item of message.items) {
            const wasmId = this.fontMap.get(item.fontId) ?? item.fontId;
            results.push(this.backend.shape(wasmId, item.text, item.options));
          }
          postResponse({ type: "batchResult", id: message.id, results });
          break;
        }
      }
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      postResponse({
        type: "error",
        id: "id" in message ? message.id : undefined,
        error: errorMsg,
      });
    }
  }
}

/**
 * Client interface for offloading shaping tasks to a worker pool or fallback in-thread engine.
 */
export class ShapingWorkerClient {
  private backend?: RustybuzzBackend;

  async init(wasmInput?: ArrayBuffer | Uint8Array): Promise<void> {
    await initShapingWasm(wasmInput);
    this.backend = new RustybuzzBackend();
  }

  async registerFont(fontData: Uint8Array): Promise<number> {
    if (!this.backend) await this.init();
    return this.backend!.registerFont(fontData);
  }

  async dropFont(fontId: number): Promise<void> {
    if (this.backend) {
      this.backend.dropFont(fontId);
    }
  }

  async shape(fontId: number, text: string, options?: ShapingOptions): Promise<ShapingResult> {
    if (!this.backend) await this.init();
    return this.backend!.shape(fontId, text, options);
  }

  async shapeBatch(
    items: readonly {
      readonly fontId: number;
      readonly text: string;
      readonly options?: ShapingOptions;
    }[],
  ): Promise<readonly ShapingResult[]> {
    if (!this.backend) await this.init();
    return items.map((item) => this.backend!.shape(item.fontId, item.text, item.options));
  }

  dispose(): void {
    // Cleanup resources
  }
}
