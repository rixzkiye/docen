export interface FontCacheOptions {
  readonly maxMemoryEntries?: number;
  readonly enableOpfs?: boolean;
}

export class FontCache {
  private readonly memoryCache = new Map<string, Uint8Array>();
  private readonly maxEntries: number;
  private readonly enableOpfs: boolean;
  private opfsDirectory: any = null;

  constructor(options?: FontCacheOptions) {
    this.maxEntries = options?.maxMemoryEntries ?? 50;
    this.enableOpfs = options?.enableOpfs ?? true;
  }

  private async getOpfsRoot(): Promise<any> {
    if (!this.enableOpfs) return null;
    if (this.opfsDirectory) return this.opfsDirectory;

    if (
      typeof navigator !== "undefined" &&
      navigator.storage &&
      typeof navigator.storage.getDirectory === "function"
    ) {
      try {
        const root = await navigator.storage.getDirectory();
        this.opfsDirectory = await root.getDirectoryHandle("docen-fonts", {
          create: true,
        });
        return this.opfsDirectory;
      } catch {
        return null;
      }
    }
    return null;
  }

  async get(key: string): Promise<Uint8Array | undefined> {
    // 1. Check memory LRU cache
    if (this.memoryCache.has(key)) {
      const data = this.memoryCache.get(key)!;
      // Refresh LRU order
      this.memoryCache.delete(key);
      this.memoryCache.set(key, data);
      return data;
    }

    // 2. Check OPFS persistent storage
    const opfs = await this.getOpfsRoot();
    if (opfs) {
      try {
        const fileHandle = await opfs.getFileHandle(key);
        const file = await fileHandle.getFile();
        const buffer = await file.arrayBuffer();
        const data = new Uint8Array(buffer);
        this.putMemory(key, data);
        return data;
      } catch {
        // Not found in OPFS
      }
    }

    return undefined;
  }

  async set(key: string, data: Uint8Array): Promise<void> {
    this.putMemory(key, data);

    const opfs = await this.getOpfsRoot();
    if (opfs) {
      try {
        const fileHandle = await opfs.getFileHandle(key, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(data);
        await writable.close();
      } catch {
        // Ignore OPFS write errors
      }
    }
  }

  async delete(key: string): Promise<boolean> {
    const inMem = this.memoryCache.delete(key);

    const opfs = await this.getOpfsRoot();
    let inOpfs = false;
    if (opfs) {
      try {
        await opfs.removeEntry(key);
        inOpfs = true;
      } catch {
        // Not found
      }
    }

    return inMem || inOpfs;
  }

  async clear(): Promise<void> {
    this.memoryCache.clear();
    const opfs = await this.getOpfsRoot();
    if (opfs) {
      try {
        // Remove and recreate
        const root = await navigator.storage.getDirectory();
        await root.removeEntry("docen-fonts", { recursive: true });
        this.opfsDirectory = null;
      } catch {
        // Ignore
      }
    }
  }

  size(): number {
    return this.memoryCache.size;
  }

  private putMemory(key: string, data: Uint8Array): void {
    if (this.memoryCache.has(key)) {
      this.memoryCache.delete(key);
    } else if (this.memoryCache.size >= this.maxEntries) {
      // Evict oldest entry (first item in Map)
      const oldestKey = this.memoryCache.keys().next().value;
      if (oldestKey !== undefined) {
        this.memoryCache.delete(oldestKey);
      }
    }
    this.memoryCache.set(key, data);
  }
}
