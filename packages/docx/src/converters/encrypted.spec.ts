import { describe, expect, it } from "vitest";

import {
  ENCRYPTED_DOCUMENT_CODE,
  EncryptedDocumentError,
  generateDOCXSync,
  isEncryptedContainerBytes,
  parseDOCX,
  parseDOCXSync,
} from "../index";

/**
 * Item-14 encrypted-container oracle: a CFB (compound file) package is not a
 * ZIP, so the pipeline must refuse it with the typed encrypted error — never
 * with a misleading archive/zip rejection and never by silently resolving an
 * empty document.
 */

/** Minimal CFB header (the 8-byte magic is all the detector needs). */
const CFB_BYTES = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x00]);

describe("encrypted document containers", () => {
  it("recognizes the CFB signature", () => {
    expect(isEncryptedContainerBytes(CFB_BYTES)).toBe(true);
    expect(isEncryptedContainerBytes(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).toBe(false);
    expect(isEncryptedContainerBytes(new Uint8Array(4))).toBe(false);
  });

  it("parseDOCXSync throws EncryptedDocumentError, not an archive rejection", () => {
    let thrown: unknown;
    try {
      parseDOCXSync(CFB_BYTES);
    } catch (err) {
      thrown = err;
    }
    expect(thrown).toBeInstanceOf(EncryptedDocumentError);
    expect((thrown as EncryptedDocumentError).code).toBe(ENCRYPTED_DOCUMENT_CODE);
    expect((thrown as Error).message).toMatch(/encrypted|password/i);
  });

  it("parseDOCX (async) throws the same typed error", async () => {
    await expect(parseDOCX(CFB_BYTES)).rejects.toBeInstanceOf(EncryptedDocumentError);
  });

  it("does not affect ordinary packages", () => {
    const bytes = generateDOCXSync({ type: "doc", content: [{ type: "paragraph" }] });
    const json = parseDOCXSync(new Uint8Array(bytes as Buffer));
    expect(json.type).toBe("doc");
  });
});
