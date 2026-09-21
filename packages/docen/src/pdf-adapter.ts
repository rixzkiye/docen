import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import type { DocumentOptions, JSONContent } from "@docen/docx";
import { renderPdf, type RenderPdfOptions } from "@docen/pdf/node";

export type PdfBackendType = "node" | "chromium" | (string & {});

export interface PdfBackendAdapter {
  readonly name: string;
  render(input: JSONContent | DocumentOptions, options?: GeneratePdfOptions): Promise<Uint8Array>;
}

export interface ChromiumBackendOptions {
  /** Explicit path to chromium binary (or CHROME_PATH env). */
  chromePath?: string;
  /** Timeout in ms per render (default: 60000). */
  timeoutMs?: number;
  /** Path to scripts/pdf-render.mjs or custom worker script */
  scriptPath?: string;
  /** Fixed date string for reproducible PDF generation */
  date?: string;
}

export interface GeneratePdfOptions extends RenderPdfOptions {
  /**
   * PDF rendering backend to use:
   * - "node" (default): Pure Node.js headless backend using @docen/pdf + PaintKit.
   *   Fast (~50-200ms), low memory (~5-20 MB transient), zero browser / DOM dependencies.
   * - "chromium": Headless Chromium browser rendering via Playwright / Chrome CLI.
   *   Exact browser DOM canvas match, supports remote browser clusters.
   * - Custom adapter: Object implementing PdfBackendAdapter.
   */
  backend?: PdfBackendType | PdfBackendAdapter;

  /**
   * Chromium-specific options when backend is "chromium".
   */
  chromium?: ChromiumBackendOptions;
}

/**
 * Pure Node.js headless PDF backend adapter.
 * Uses @docen/pdf + PaintKit without browser or DOM.
 */
export const nodeBackendAdapter: PdfBackendAdapter = {
  name: "node",
  async render(input, options) {
    return renderPdf(input, options);
  },
};

/**
 * Headless Chromium browser PDF backend adapter.
 * Delegates to scripts/pdf-render.mjs via Playwright/Chromium.
 */
export const chromiumBackendAdapter: PdfBackendAdapter = {
  name: "chromium",
  async render(input, options) {
    const scriptPath =
      options?.chromium?.scriptPath ?? path.resolve(process.cwd(), "scripts/pdf-render.mjs");

    if (!fs.existsSync(scriptPath)) {
      throw new Error(
        `Chromium PDF backend script not found at "${scriptPath}". Ensure scripts/pdf-render.mjs is present or provide options.chromium.scriptPath.`,
      );
    }

    const tmpJson = path.join(
      os.tmpdir(),
      `docen-pdf-${Date.now()}-${Math.random().toString(36).slice(2)}.json`,
    );
    fs.writeFileSync(tmpJson, JSON.stringify(input));

    const args = ["--json", tmpJson, "--out", "-"];
    if (options?.chromium?.date) {
      args.push("--date", options.chromium.date);
    }
    if (options?.pdfa) {
      args.push("--pdfa", typeof options.pdfa === "string" ? options.pdfa : "2b");
    }
    if (options?.pdfUa) {
      args.push("--pdf-ua");
    }
    if (options?.textMode) {
      args.push("--mode", options.textMode);
    }

    try {
      // Lazy so bundlers never pull node:child_process statically; this
      // adapter is server-only (see the node-only "./pdf" export).
      const { execFileSync } = await import("node:child_process");
      const env = {
        ...process.env,
        ...(options?.chromium?.chromePath ? { CHROME_PATH: options.chromium.chromePath } : {}),
      };

      const stdout = execFileSync("node", [scriptPath, ...args], {
        env,
        timeout: options?.chromium?.timeoutMs ?? 60_000,
        maxBuffer: 64 * 1024 * 1024,
      });

      return new Uint8Array(stdout.buffer, stdout.byteOffset, stdout.byteLength);
    } finally {
      if (fs.existsSync(tmpJson)) {
        try {
          fs.unlinkSync(tmpJson);
        } catch {
          // ignore cleanup error
        }
      }
    }
  },
};

const backendRegistry = new Map<string, PdfBackendAdapter>([
  ["node", nodeBackendAdapter],
  ["chromium", chromiumBackendAdapter],
]);

export function registerPdfBackend(adapter: PdfBackendAdapter): void {
  backendRegistry.set(adapter.name, adapter);
}

export function getPdfBackend(name: string): PdfBackendAdapter | undefined {
  return backendRegistry.get(name);
}

/**
 * Generate a PDF from Tiptap JSON or DocumentOptions.
 *
 * High-level unified API mirroring `generateDOCX`:
 * - Uses the pure Node headless backend by default (`@docen/pdf` + `PaintKit`).
 * - Supports custom or Chromium-based backends via `options.backend`.
 * - Emits standard PDF Uint8Array bytes.
 */
export async function generatePDF(
  input: JSONContent | DocumentOptions,
  options?: GeneratePdfOptions,
): Promise<Uint8Array> {
  const backendSpec = options?.backend ?? "node";
  if (typeof backendSpec === "object" && backendSpec !== null && "render" in backendSpec) {
    return backendSpec.render(input, options);
  }
  const adapter = backendRegistry.get(backendSpec);
  if (!adapter) {
    throw new Error(
      `Unknown PDF backend: "${backendSpec}". Available backends: ${Array.from(backendRegistry.keys()).join(", ")}`,
    );
  }
  return adapter.render(input, options);
}
