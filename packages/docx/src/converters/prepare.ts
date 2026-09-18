import { encodeBase64 } from "@office-open/core";
import { imageMeta } from "image-meta";

import type { JSONContent } from "../core";

// ── Types ──

/**
 * A prepare step that transforms Tiptap JSON. Steps run on a caller-owned copy
 * of the document and may mutate it in place (e.g. embed fetched resources);
 * {@link prepareDocument} never mutates the document it is given.
 */
export type PrepareStep = (json: JSONContent) => Promise<void>;

/**
 * Fetch handler for external image URLs.
 *
 * Receives the URL, returns the image binary data.
 * Override to customize fetching (proxy, auth, caching, etc.).
 */
export type ImageFetchHandler = (url: string) => Promise<Uint8Array>;

/**
 * Explicit opt-in policy for fetching `http(s)` images. There is no implicit
 * network access: {@link prepareImages} without `allow` entries is a no-op, and
 * only the listed hosts are contacted.
 */
export interface PrepareImagesPolicy {
  /**
   * Hostnames allowed to be fetched (required — an empty list disables the
   * step). Case-insensitive match on `URL.hostname`: an exact entry matches
   * only that host, a `*.` prefix matches the bare domain and its subdomains
   * (`"*.example.com"` → `example.com`, `a.example.com`). Ports are ignored; a
   * custom `fetch` transport is trusted with the allowlisted URL.
   */
  allow: readonly string[];
  /** Maximum accepted image size in bytes (default {@link DEFAULT_IMAGE_MAX_BYTES}). */
  maxBytes?: number;
  /** Maximum redirects followed by the default transport (default {@link DEFAULT_IMAGE_MAX_REDIRECTS}). */
  maxRedirects?: number;
  /** Per-request timeout in ms for the default transport (default {@link DEFAULT_IMAGE_TIMEOUT_MS}). */
  timeoutMs?: number;
  /**
   * Custom transport (proxy/auth/caching). Passing one is an explicit opt-in;
   * the scheme/host allowlist and the `maxBytes` cap still apply to it, while
   * redirect/timeout policy is the transport's own.
   */
  fetch?: ImageFetchHandler;
}

// ── Built-in steps ──

/** Default per-image response cap (8 MiB, mirroring the Akademi importer's owned-media cap). */
export const DEFAULT_IMAGE_MAX_BYTES = 8 * 1024 * 1024;
/** Default redirect cap for the built-in transport. */
export const DEFAULT_IMAGE_MAX_REDIRECTS = 3;
/** Default per-request timeout for the built-in transport. */
export const DEFAULT_IMAGE_TIMEOUT_MS = 10_000;

const REDIRECT_STATUSES = new Set([301, 302, 303, 307, 308]);

/**
 * Create a prepare step that fetches external image URLs and converts them to
 * data URLs. Network access only happens when a policy with a non-empty
 * `allow` list is passed — the default is a no-op that leaves external URLs
 * untouched (the renderer drops images without embedded data).
 *
 * Enforced on every fetch: `http(s)` scheme only, host allowlist, response
 * size cap, redirect cap and a per-request timeout (defaults to the built-in
 * `fetch`; a custom `fetch` transport keeps the scheme/host/size caps).
 *
 * @example
 * ```ts
 * // Server opt-in: fetch only from the CDN
 * await generateDOCX(json, {
 *   prepare: [prepareImages({ allow: ["cdn.example.com"] }), prepareImageSizes()],
 * });
 *
 * // Default: no network — external http images stay untouched
 * await prepareDocument(json);
 * ```
 */
export function prepareImages(policy?: PrepareImagesPolicy): PrepareStep {
  const allow = normalizeAllowList(policy?.allow);
  if (!policy || allow.length === 0) return async () => undefined;

  const maxBytes = policy.maxBytes ?? DEFAULT_IMAGE_MAX_BYTES;
  const maxRedirects = policy.maxRedirects ?? DEFAULT_IMAGE_MAX_REDIRECTS;
  const timeoutMs = policy.timeoutMs ?? DEFAULT_IMAGE_TIMEOUT_MS;
  const handler = async (url: string): Promise<Uint8Array> => {
    assertAllowedUrl(url, allow);
    const bytes = policy.fetch
      ? await policy.fetch(url)
      : await fetchWithinPolicy(url, { allow, maxBytes, maxRedirects, timeoutMs });
    if (bytes.byteLength > maxBytes) {
      throw new Error(`Image exceeds the ${maxBytes}-byte cap: ${url}`);
    }
    return bytes;
  };

  return async (json: JSONContent) => {
    await walkImages(json, handler);
  };
}

// ── Direct fetch (explicit UI actions) ──

/**
 * Fetch one image for an explicit, user-initiated action (Insert → Online
 * Pictures). Unlike {@link prepareImages} — a batch policy for untrusted
 * documents that requires a host allowlist — this runs on an explicit user
 * action, so `http(s)` is allowed without a host list while the same scheme,
 * credentials, size, redirect and timeout guards apply. `data:image/...`
 * URLs are decoded locally and never hit the network.
 */
export interface FetchImageOptions {
  /** Optional host allowlist; when non-empty only those hosts are contacted. */
  allow?: readonly string[];
  /** Response cap in bytes (default {@link DEFAULT_IMAGE_MAX_BYTES}). */
  maxBytes?: number;
  /** Redirect cap for the default transport (default {@link DEFAULT_IMAGE_MAX_REDIRECTS}). */
  maxRedirects?: number;
  /** Per-request timeout in ms for the default transport (default {@link DEFAULT_IMAGE_TIMEOUT_MS}). */
  timeoutMs?: number;
  /** Custom transport (proxy/auth/caching); scheme/credentials and the size cap still apply. */
  fetch?: ImageFetchHandler;
}

/**
 * Fetch one image for an explicit UI action, enforcing the shared guards.
 *
 * @param raw - `data:image/...` URL or public `http(s)` URL (no credentials)
 * @param options - Optional allowlist/caps/transport overrides
 */
export async function fetchImageHandler(
  raw: string,
  options: FetchImageOptions = {},
): Promise<Uint8Array> {
  const maxBytes = options.maxBytes ?? DEFAULT_IMAGE_MAX_BYTES;
  const embedded = decodeDataUrl(raw);
  if (embedded) {
    if (embedded.byteLength > maxBytes) {
      throw new Error(`Image exceeds the ${maxBytes}-byte cap`);
    }
    return embedded;
  }
  const allow = normalizeAllowList(options.allow);
  assertAllowedUrl(raw, allow);
  const bytes = options.fetch
    ? await options.fetch(raw)
    : await fetchWithinPolicy(raw, {
        allow,
        maxBytes,
        maxRedirects: options.maxRedirects ?? DEFAULT_IMAGE_MAX_REDIRECTS,
        timeoutMs: options.timeoutMs ?? DEFAULT_IMAGE_TIMEOUT_MS,
      });
  if (bytes.byteLength > maxBytes) {
    throw new Error(`Image exceeds the ${maxBytes}-byte cap: ${raw}`);
  }
  return bytes;
}

/**
 * Create a prepare step that fills in missing image dimensions by reading image
 * metadata. Intended to run after {@link prepareImages} so HTTP sources are
 * already embedded as data URLs.
 *
 * Only images lacking `width` or `height` are probed — the DOCX round-trip path
 * already carries `transformation.width/height`, so this mainly serves images
 * entering via HTML/Markdown (which carry no intrinsic size). Images whose
 * metadata can't be read are left untouched; `renderDocx` falls back to 400×300
 * (see extensions/image.ts).
 */
export function prepareImageSizes(): PrepareStep {
  return async (json: JSONContent) => {
    await walkImageSizes(json);
  };
}

// ── Pipeline ──

/** Built-in prepare steps, run when no custom steps are provided. */
const DEFAULT_STEPS: readonly PrepareStep[] = [prepareImageSizes()];

/**
 * Run prepare steps on a copy of a Tiptap JSON document before compilation.
 *
 * The input is never mutated: the document is `structuredClone`d and the steps
 * run on the clone, which is returned. Steps run sequentially in order.
 *
 * Defaults to `[prepareImageSizes()]` when no steps are provided — the default
 * pipeline never touches the network. Fetching `http(s)` images requires an
 * explicit {@link prepareImages} policy with an `allow` list.
 *
 * @returns The prepared copy (safe to compile; the input keeps its own value).
 *
 * @example
 * ```ts
 * const json = parseHTML(html);
 * const prepared = await prepareDocument(json); // local-only; json unchanged
 * const docOpts = compileDocument(prepared);
 * ```
 */
export async function prepareDocument(
  json: JSONContent,
  steps: readonly PrepareStep[] = DEFAULT_STEPS,
): Promise<JSONContent> {
  const prepared = structuredClone(json);
  for (const step of steps) {
    await step(prepared);
  }
  return prepared;
}

// ── Internals ──

function normalizeAllowList(allow: readonly string[] | undefined): readonly string[] {
  return (allow ?? []).map((host) => host.trim().toLowerCase()).filter((host) => host.length > 0);
}

function hostAllowed(hostname: string, allow: readonly string[]): boolean {
  const host = hostname.toLowerCase();
  return allow.some((entry) =>
    entry.startsWith("*.")
      ? host === entry.slice(2) || host.endsWith(entry.slice(1))
      : host === entry,
  );
}

/** Reject anything that is not an allowlisted `http(s)` URL without credentials. */
function assertAllowedUrl(raw: string, allow: readonly string[]): void {
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    throw new Error(`Unsupported image URL: ${raw}`);
  }
  if (url.protocol !== "http:" && url.protocol !== "https:") {
    throw new Error(`Unsupported image URL scheme: ${url.protocol}`);
  }
  if (url.username || url.password) {
    throw new Error("Image URL credentials are not allowed");
  }
  // An empty allowlist means "any public host" (used by the direct,
  // user-initiated `fetchImageHandler`); batch `prepareImages` always passes a
  // non-empty list.
  if (allow.length > 0 && !hostAllowed(url.hostname, allow)) {
    throw new Error(`Image host is not allowlisted: ${url.hostname}`);
  }
}

interface FetchPolicy {
  readonly allow: readonly string[];
  readonly maxBytes: number;
  readonly maxRedirects: number;
  readonly timeoutMs: number;
}

/**
 * Default transport: manual redirects (each hop re-validated against the
 * allowlist and counted), streamed body with a byte cap, per-hop timeout.
 */
async function fetchWithinPolicy(raw: string, policy: FetchPolicy): Promise<Uint8Array> {
  let current = raw;
  for (let redirects = 0; ; redirects++) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), policy.timeoutMs);
    try {
      const response = await fetch(current, { redirect: "manual", signal: controller.signal });
      if (REDIRECT_STATUSES.has(response.status)) {
        if (redirects >= policy.maxRedirects) {
          throw new Error(`Image exceeded the ${policy.maxRedirects}-redirect cap: ${raw}`);
        }
        const location = response.headers.get("location");
        if (!location) throw new Error(`Redirect without a Location header: ${current}`);
        current = new URL(location, current).toString();
        assertAllowedUrl(current, policy.allow);
        continue;
      }
      if (!response.ok) {
        throw new Error(
          `Failed to fetch image from ${current}: ${response.status} ${response.statusText}`,
        );
      }
      const declared = Number(response.headers.get("content-length"));
      if (Number.isFinite(declared) && declared > policy.maxBytes) {
        throw new Error(`Image exceeds the ${policy.maxBytes}-byte cap: ${current}`);
      }
      return await readCapped(response, policy.maxBytes, current);
    } finally {
      clearTimeout(timer);
    }
  }
}

/** Read a response body, aborting as soon as it crosses `maxBytes`. */
async function readCapped(response: Response, maxBytes: number, url: string): Promise<Uint8Array> {
  if (!response.body) {
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (bytes.byteLength > maxBytes) {
      throw new Error(`Image exceeds the ${maxBytes}-byte cap: ${url}`);
    }
    return bytes;
  }
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > maxBytes) {
      await reader.cancel();
      throw new Error(`Image exceeds the ${maxBytes}-byte cap: ${url}`);
    }
    chunks.push(value);
  }
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}

async function toDataUrl(src: string, handler: ImageFetchHandler): Promise<string> {
  const data = await handler(src);
  // Infer MIME from the image header bytes (imageMeta reads magic bytes), not
  // the URL extension — extensions are unreliable for CDN/authenticated/redirect
  // URLs (`img.png?token=`, `/avatar`, no extension) and would mislabel a real
  // JPEG as PNG, corrupting the embedded blip and breaking re-import.
  let mime = "image/png";
  try {
    const type = imageMeta(data).type;
    if (type) mime = type === "jpg" ? "image/jpeg" : `image/${type}`;
  } catch {
    // Unreadable/unknown header — leave the PNG fallback.
  }
  return `data:${mime};base64,${encodeBase64(data)}`;
}

async function walkImages(node: JSONContent, handler: ImageFetchHandler): Promise<void> {
  if (node.type === "image" && node.attrs) {
    const src = node.attrs.src as string | undefined;
    if (src && (src.startsWith("http://") || src.startsWith("https://"))) {
      try {
        node.attrs.src = await toDataUrl(src, handler);
      } catch (error) {
        console.warn(
          `Failed to fetch image: ${src}`,
          error instanceof Error ? error.message : error,
        );
      }
    }
  }

  const tasks: Promise<void>[] = [];
  for (const child of node.content ?? []) {
    tasks.push(walkImages(child, handler));
  }
  await Promise.all(tasks);
}

async function walkImageSizes(node: JSONContent): Promise<void> {
  if (node.type === "image" && node.attrs) {
    const attrs = node.attrs;
    const needsWidth = attrs.width == null;
    const needsHeight = attrs.height == null;
    if (needsWidth || needsHeight) {
      const bytes = decodeDataUrl(attrs.src as string | undefined);
      if (bytes) {
        try {
          const meta = imageMeta(bytes);
          if (needsWidth && typeof meta.width === "number") attrs.width = meta.width;
          if (needsHeight && typeof meta.height === "number") attrs.height = meta.height;
        } catch {
          // Unreadable or unsupported image — leave attrs; renderDocx falls back.
        }
      }
    }
  }

  const tasks: Promise<void>[] = [];
  for (const child of node.content ?? []) {
    tasks.push(walkImageSizes(child));
  }
  await Promise.all(tasks);
}

/** Decode a `data:image/...;base64,...` URL to bytes; null if not a data URL. */
function decodeDataUrl(src: string | undefined): Uint8Array | null {
  if (!src) return null;
  const match = src.match(/^data:image\/[\w.+-]+;base64,(.+)$/);
  if (!match) return null;
  const binary = atob(match[1]);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}
