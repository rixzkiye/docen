/**
 * Encrypted (password-protected) OOXML containers.
 *
 * A protected .docx is not a ZIP at all — it is a Compound File Binary (CFB)
 * container whose streams hold an agile/standard encryption envelope. office-
 * open can pass such a container through at the `DocumentOptions` level
 * (`{ sections: [], encrypted: { data } }`), but docen's document pipeline has
 * no password UI and cannot decrypt it: resolving it would silently produce an
 * empty page. The public parse entries therefore reject the container with an
 * explicit, typed error the editor surfaces as an open refusal.
 */

/** CFB signature (`D0 CF 11 E0 A1 B1 1A E1`) — an OLE compound container. */
const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

export const ENCRYPTED_DOCUMENT_CODE = "ENCRYPTED_DOCUMENT";

/** Thrown by `parseDOCX`/`parseDOCXSync` for a password-protected package. */
export class EncryptedDocumentError extends Error {
  readonly code = ENCRYPTED_DOCUMENT_CODE;

  constructor(message = "This document is encrypted (password-protected) and cannot be opened.") {
    super(message);
    this.name = "EncryptedDocumentError";
  }
}

/** Whether the bytes start with the CFB compound-file signature. */
export function isEncryptedContainerBytes(data: Uint8Array): boolean {
  if (data.byteLength < CFB_MAGIC.length) return false;
  for (let i = 0; i < CFB_MAGIC.length; i++) {
    if (data[i] !== CFB_MAGIC[i]) return false;
  }
  return true;
}

/** Reject an encrypted container before any ZIP/archive work runs — a CFB is
 *  not a ZIP, so the archive guard would otherwise report a misleading zip
 *  error for a merely locked file. */
export function assertNotEncryptedContainer(data: Uint8Array): void {
  if (isEncryptedContainerBytes(data)) throw new EncryptedDocumentError();
}
