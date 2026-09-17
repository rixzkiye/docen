import type { FontIdentity } from "./types.js";

export interface FontDescriptor {
  readonly name: string;
  readonly postscriptName?: string;
  readonly familyName?: string;
  readonly weight?: number;
  readonly style?: "normal" | "italic";
  readonly script?: string;
  readonly hash?: string;
}

/**
 * Computes a deterministic SHA-256 hex digest for font byte content.
 */
export async function computeFontHash(fontBytes: Uint8Array): Promise<string> {
  const hashBuffer = await crypto.subtle.digest("SHA-256", fontBytes as BufferSource);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
}

export function createFontDescriptor(identity: FontIdentity, hash?: string): FontDescriptor {
  return {
    name: identity.familyName ?? identity.postscriptName ?? "Unknown Font",
    familyName: identity.familyName,
    postscriptName: identity.postscriptName,
    hash,
  };
}
