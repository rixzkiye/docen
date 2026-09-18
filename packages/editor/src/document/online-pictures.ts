import { fetchImageHandler } from "@docen/docx";
import { imageMeta } from "image-meta";

/**
 * Insert → Pictures → Online Pictures.
 *
 * Word's dialog searches Bing; the browser element has no picture service, so
 * the equivalent is a URL entry that fetches the image at insert time. The
 * source may be an http(s) URL or a `data:image/…` URL (the same payloads the
 * DOCX prepare step embeds), and the fetched bytes are guarded before they
 * reach the document: a byte cap, a magic-byte image sniff (the URL extension
 * is never trusted), and a natural-size probe so the node lands at Word's
 * natural-size-else-content-width geometry.
 *
 * The fetch goes through the engine's public `fetchImageHandler` (the same
 * handler `prepareDocument` uses) — if the R8-D1 prepare policy later adds an
 * allowlist or size cap there, this call site inherits it.
 */

/** The online-picture download cap (10 MiB) — a larger response is refused
 *  before it becomes a document attachment. */
export const ONLINE_PICTURE_MAX_BYTES = 10 * 1024 * 1024;

/** Why an online picture could not be embedded — the dialog maps each to a
 *  localized, actionable message. */
export type OnlinePictureFailure =
  | "empty-url"
  | "invalid-url"
  | "unsupported-url"
  | "fetch-failed"
  | "too-large"
  | "unsupported-type";

/** The embedded picture the dialog staged. */
export interface OnlinePicture {
  /** The embedded `data:image/…` source. */
  src: string;
  /** Natural width in px (absent when the probe failed). */
  width?: number;
  /** Natural height in px (absent when the probe failed). */
  height?: number;
  /** Alt text derived from the source's file name (empty for data URLs). */
  alt: string;
}

export type OnlinePictureResult =
  | { ok: true; picture: OnlinePicture }
  | { ok: false; reason: OnlinePictureFailure };

export interface OnlinePictureOptions {
  /** Fetch handler (default: the engine's public `fetchImageHandler`). An
   *  injected handler keeps tests hermetic and lets a host proxy requests. */
  fetch?: (url: string) => Promise<Uint8Array>;
  /** Byte cap (default {@link ONLINE_PICTURE_MAX_BYTES}). */
  maxBytes?: number;
  /** Natural-size probe; absent leaves width/height unset (the painter then
   *  falls back to the image's own default box). */
  probeSize?: (src: string) => Promise<{ w: number; h: number }>;
}

/** image-meta's type tokens → the MIME the embedded data URL carries. */
const MIME_BY_TYPE: Record<string, string> = {
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  png: "image/png",
  gif: "image/gif",
  webp: "image/webp",
  bmp: "image/bmp",
  tiff: "image/tiff",
  ico: "image/x-icon",
  avif: "image/avif",
  svg: "image/svg+xml",
};

/** Base64-encode bytes (chunked — a 10 MiB spread would blow the call stack). */
export function base64FromBytes(bytes: Uint8Array): string {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

/** The file-name-derived alt text (Word names an inserted web picture after
 *  its source file); empty when the source carries no usable name. */
export function onlinePictureAlt(url: string): string {
  if (/^data:/i.test(url)) return "";
  try {
    const name = decodeURIComponent(new URL(url).pathname.split("/").pop() ?? "");
    return name.replace(/\.[^.]+$/, "");
  } catch {
    return "";
  }
}

/**
 * Validate + download + embed one online picture source. Every failure is a
 * typed reason — the dialog renders the matching localized message without
 * ever throwing.
 */
export async function resolveOnlinePicture(
  url: string,
  options: OnlinePictureOptions = {},
): Promise<OnlinePictureResult> {
  const trimmed = (url ?? "").trim();
  if (!trimmed) return { ok: false, reason: "empty-url" };
  const isData = /^data:image\//i.test(trimmed);
  if (!isData) {
    if (!/^https?:\/\//i.test(trimmed)) return { ok: false, reason: "unsupported-url" };
    try {
      if (!/^https?:$/.test(new URL(trimmed).protocol)) {
        return { ok: false, reason: "unsupported-url" };
      }
    } catch {
      return { ok: false, reason: "invalid-url" };
    }
  }

  let bytes: Uint8Array;
  try {
    bytes = await (options.fetch ?? fetchImageHandler)(trimmed);
  } catch {
    return { ok: false, reason: "fetch-failed" };
  }
  if (!(bytes instanceof Uint8Array) || bytes.byteLength === 0) {
    return { ok: false, reason: "fetch-failed" };
  }
  if (bytes.byteLength > (options.maxBytes ?? ONLINE_PICTURE_MAX_BYTES)) {
    return { ok: false, reason: "too-large" };
  }

  // Magic-byte sniff — a URL ending in .png that returns HTML must not embed.
  let type: string | undefined;
  try {
    type = imageMeta(bytes).type;
  } catch {
    return { ok: false, reason: "unsupported-type" };
  }
  const mime = type ? MIME_BY_TYPE[type] : undefined;
  if (!mime) return { ok: false, reason: "unsupported-type" };

  const src = `data:${mime};base64,${base64FromBytes(bytes)}`;
  let width: number | undefined;
  let height: number | undefined;
  if (options.probeSize) {
    try {
      const size = await options.probeSize(src);
      if (size.w > 0 && size.h > 0) {
        width = size.w;
        height = size.h;
      }
    } catch {
      // Undecodable by the host image element — keep the source, no size.
    }
  }
  return {
    ok: true,
    picture: {
      src,
      ...(width != null ? { width } : {}),
      ...(height != null ? { height } : {}),
      alt: onlinePictureAlt(trimmed),
    },
  };
}
