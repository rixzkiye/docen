import { RustybuzzBackend } from "./rustybuzz-backend.js";
import type { ShapingOptions, ShapingResult } from "./types.js";
import { initShapingWasm } from "./wasm-loader.js";

export type WorkerRequest =
  | { readonly type: "init"; readonly id?: number; readonly wasmBytes?: ArrayBuffer }
  | {
      readonly type: "registerFont";
      readonly id?: number;
      readonly fontId: number;
      readonly fontData: Uint8Array;
    }
  | { readonly type: "dropFont"; readonly id?: number; readonly fontId: number }
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
  | {
      readonly type: "initDone";
      readonly id?: number;
      readonly success: boolean;
      readonly error?: string;
    }
  | { readonly type: "fontRegistered"; readonly id?: number; readonly fontId: number }
  | { readonly type: "fontDropped"; readonly id?: number; readonly fontId: number }
  | {
      readonly type: "batchResult";
      readonly id: number;
      readonly results: readonly ShapingResult[];
    }
  | { readonly type: "error"; readonly id?: number; readonly error: string };

/**
 * Worker handler implementation for execution inside a Web Worker or Node worker_thread.
 *
 * A worker entry is five lines:
 *
 * ```ts
 * import { ShapingWorkerHandler } from "@docen/shaping";
 * const handler = new ShapingWorkerHandler();
 * self.onmessage = (event) =>
 *   handler.handleMessage(event.data, (response, transfer) => self.postMessage(response, transfer));
 * ```
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
    const id = "id" in message ? message.id : undefined;
    try {
      switch (message.type) {
        case "init": {
          await initShapingWasm(message.wasmBytes);
          this.backend = new RustybuzzBackend();
          this.isInitialized = true;
          postResponse({ type: "initDone", ...(id !== undefined ? { id } : {}), success: true });
          break;
        }
        case "registerFont": {
          if (!this.backend) throw new Error("Worker not initialized");
          const wasmId = this.backend.registerFont(message.fontData);
          this.fontMap.set(message.fontId, wasmId);
          postResponse({
            type: "fontRegistered",
            ...(id !== undefined ? { id } : {}),
            fontId: message.fontId,
          });
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
          postResponse({
            type: "fontDropped",
            ...(id !== undefined ? { id } : {}),
            fontId: message.fontId,
          });
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
        ...(id !== undefined ? { id } : {}),
        error: errorMsg,
      });
    }
  }
}

/** The slice of the Worker API the client uses — a DOM Worker, a Node
 *  worker_threads Worker adapter, or a test double. */
export interface ShapingWorkerTransport {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  addEventListener(
    type: "message" | "error",
    listener: (event: { readonly data?: unknown; readonly message?: string }) => void,
  ): void;
  terminate?(): void;
}

/**
 * Client for offloading shaping to a worker, with an honest in-thread
 * fallback. Pass a Worker built from the {@link ShapingWorkerHandler} entry
 * (see the README) to run shaping off the main thread; without one — or when
 * the worker fails to initialize — every call runs the same WASM engine on the
 * calling thread. `offloaded` reports which path is live.
 */
export class ShapingWorkerClient {
  private worker?: ShapingWorkerTransport;
  private backend?: RustybuzzBackend;
  private readonly fontMap = new Map<number, number>();
  private nextFontId = 1;
  private nextRequestId = 1;
  private readonly pending = new Map<
    number,
    {
      resolve: (response: WorkerResponse) => void;
      reject: (error: Error) => void;
      timer?: ReturnType<typeof setTimeout>;
    }
  >();
  private pendingInit?: {
    resolve: () => void;
    reject: (error: Error) => void;
    timer?: ReturnType<typeof setTimeout>;
  };

  constructor(worker?: ShapingWorkerTransport) {
    this.worker = worker;
    this.worker?.addEventListener("message", (event) => {
      this.#handleResponse(event.data as WorkerResponse);
    });
    this.worker?.addEventListener("error", (event) => {
      const error = new Error(event.message ?? "Shaping worker error");
      this.#rejectAll(error);
    });
  }

  /** Whether shaping actually runs in the worker (false = in-thread fallback). */
  get offloaded(): boolean {
    return this.worker !== undefined && this.backend === undefined;
  }

  async init(wasmInput?: ArrayBuffer | Uint8Array): Promise<void> {
    if (this.worker) {
      try {
        const bytes = wasmInput
          ? wasmInput instanceof Uint8Array
            ? wasmInput.slice().buffer
            : wasmInput.slice(0)
          : undefined;
        await new Promise<void>((resolve, reject) => {
          const timer = setTimeout(() => {
            this.pendingInit = undefined;
            reject(new Error("Shaping worker init timed out"));
          }, 30_000);
          this.pendingInit = { resolve, reject, timer };
          this.worker!.postMessage({
            type: "init",
            id: this.nextRequestId++,
            ...(bytes ? { wasmBytes: bytes } : {}),
          });
        });
        return;
      } catch {
        // The worker could not initialize (no module support, CSP, wasm fetch
        // failed) — tear it down and shape on this thread instead.
        this.#detachWorker();
      }
    }
    await initShapingWasm(wasmInput);
    this.backend = new RustybuzzBackend();
  }

  async registerFont(fontData: Uint8Array): Promise<number> {
    const fontId = this.nextFontId++;
    if (this.worker) {
      // Transfer a copy: the caller may keep and reuse its bytes.
      const copy = fontData.slice();
      const response = await this.#request(
        {
          type: "registerFont",
          id: this.nextRequestId++,
          fontId,
          fontData: copy,
        },
        [copy.buffer],
      );
      if (response.type !== "fontRegistered") throw new Error("Unexpected worker response");
      return response.fontId;
    }
    if (!this.backend) await this.init();
    this.fontMap.set(fontId, this.backend!.registerFont(fontData));
    return fontId;
  }

  async dropFont(fontId: number): Promise<void> {
    if (this.worker) {
      await this.#request({ type: "dropFont", id: this.nextRequestId++, fontId });
      return;
    }
    const wasmId = this.fontMap.get(fontId);
    if (wasmId !== undefined) {
      this.backend?.dropFont(wasmId);
      this.fontMap.delete(fontId);
    }
  }

  async shape(fontId: number, text: string, options?: ShapingOptions): Promise<ShapingResult> {
    const [result] = await this.shapeBatch([{ fontId, text, options }]);
    if (!result) throw new Error("Shaping worker returned no result");
    return result;
  }

  async shapeBatch(
    items: readonly {
      readonly fontId: number;
      readonly text: string;
      readonly options?: ShapingOptions;
    }[],
  ): Promise<readonly ShapingResult[]> {
    if (this.worker) {
      const response = await this.#request({
        type: "shapeBatch",
        id: this.nextRequestId++,
        items,
      });
      if (response.type !== "batchResult") throw new Error("Unexpected worker response");
      return response.results;
    }
    if (!this.backend) await this.init();
    return items.map((item) =>
      this.backend!.shape(this.fontMap.get(item.fontId) ?? item.fontId, item.text, item.options),
    );
  }

  dispose(): void {
    this.#detachWorker();
    this.#rejectAll(new Error("ShapingWorkerClient disposed"));
  }

  #request(message: WorkerRequest, transfer?: Transferable[]): Promise<WorkerResponse> {
    return new Promise((resolve, reject) => {
      const id = message.id;
      if (id === undefined) {
        reject(new Error("Worker requests must carry an id"));
        return;
      }
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`Shaping worker request ${id} timed out`));
      }, 30_000);
      this.pending.set(id, { resolve, reject, timer });
      this.worker!.postMessage(message, transfer);
    });
  }

  #handleResponse(response: WorkerResponse): void {
    if (response.type === "initDone") {
      const pending = this.pendingInit;
      this.pendingInit = undefined;
      if (pending?.timer) clearTimeout(pending.timer);
      if (response.success) pending?.resolve();
      else pending?.reject(new Error(response.error ?? "Shaping worker init failed"));
      return;
    }
    if (response.type === "error") {
      if (response.id === undefined) {
        this.pendingInit?.reject(new Error(response.error));
        this.pendingInit = undefined;
        return;
      }
      const pending = this.pending.get(response.id);
      if (pending) {
        this.pending.delete(response.id);
        if (pending.timer) clearTimeout(pending.timer);
        pending.reject(new Error(response.error));
      }
      return;
    }
    const pending = this.pending.get(response.id ?? -1);
    if (pending) {
      this.pending.delete(response.id!);
      if (pending.timer) clearTimeout(pending.timer);
      pending.resolve(response);
    }
  }

  #detachWorker(): void {
    this.worker?.terminate?.();
    this.worker = undefined;
  }

  #rejectAll(error: Error): void {
    if (this.pendingInit?.timer) clearTimeout(this.pendingInit.timer);
    this.pendingInit?.reject(error);
    this.pendingInit = undefined;
    for (const pending of this.pending.values()) {
      if (pending.timer) clearTimeout(pending.timer);
      pending.reject(error);
    }
    this.pending.clear();
  }
}
