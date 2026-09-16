import { describe, expect, it } from "vitest";

import {
  detectOpenFormat,
  FLAT_OPC_UNSUPPORTED,
  OpenFormatError,
  SAVE_FORMATS,
  suggestedFileName,
} from "./file-formats";

/** A picked-file stand-in (Node's File carries no bytes until read). */
function file(name: string, type = ""): File {
  return new File([new Uint8Array([1, 2, 3])], name, { type });
}

describe("detectOpenFormat", () => {
  it("detects the docx family by extension", () => {
    expect(detectOpenFormat(file("Report.docx"))).toBe("docx");
    expect(detectOpenFormat(file("Macros.docm"))).toBe("docm");
    expect(detectOpenFormat(file("Letter.dotx"))).toBe("dotx");
    expect(detectOpenFormat(file("Letter-macros.dotm"))).toBe("dotm");
  });

  it("detects by MIME when the name carries no extension", () => {
    expect(
      detectOpenFormat(file("download", "application/vnd.ms-word.document.macroEnabled.12")),
    ).toBe("docm");
    expect(
      detectOpenFormat(
        file("download", "application/vnd.openxmlformats-officedocument.wordprocessingml.template"),
      ),
    ).toBe("dotx");
    expect(
      detectOpenFormat(file("download", "application/vnd.ms-word.template.macroEnabled.12")),
    ).toBe("dotm");
    expect(
      detectOpenFormat(
        file("download", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
      ),
    ).toBe("docx");
    expect(detectOpenFormat(file("download", "text/markdown"))).toBe("markdown");
    expect(detectOpenFormat(file("download", "application/rtf"))).toBe("rtf");
    expect(detectOpenFormat(file("download", "text/rtf"))).toBe("rtf");
    expect(detectOpenFormat(file("download", "text/plain"))).toBe("text");
  });

  it("keeps the existing docx/markdown/rtf/text detection", () => {
    expect(detectOpenFormat(file("Notes.md"))).toBe("markdown");
    expect(detectOpenFormat(file("Notes.markdown"))).toBe("markdown");
    expect(detectOpenFormat(file("Legacy.rtf"))).toBe("rtf");
    expect(detectOpenFormat(file("Notes.txt"))).toBe("text");
  });

  it("refuses Flat OPC XML with a clear message", () => {
    expect(() => detectOpenFormat(file("Flat.xml"))).toThrow(FLAT_OPC_UNSUPPORTED);
    expect(() => detectOpenFormat(file("Flat", "application/xml"))).toThrow(FLAT_OPC_UNSUPPORTED);
    expect(() => detectOpenFormat(file("Flat", "text/xml"))).toThrow(FLAT_OPC_UNSUPPORTED);
  });

  it("parses the MIME base type (browsers append parameters)", () => {
    expect(() => detectOpenFormat(file("Flat", "application/xml;charset=utf-8"))).toThrow(
      FLAT_OPC_UNSUPPORTED,
    );
    expect(detectOpenFormat(file("download", "text/markdown;charset=utf-8"))).toBe("markdown");
  });

  it("carries a machine-readable code so the host can localize the refusal", () => {
    const refusal = (picked: File): OpenFormatError | undefined => {
      try {
        detectOpenFormat(picked);
        return undefined;
      } catch (err) {
        return err as OpenFormatError;
      }
    };
    const flat = refusal(file("Flat.xml"));
    expect(flat).toBeInstanceOf(OpenFormatError);
    expect(flat?.code).toBe("flat-opc");
    const unknown = refusal(file("Legacy.pages"));
    expect(unknown).toBeInstanceOf(OpenFormatError);
    expect(unknown?.code).toBe("unsupported");
    expect(unknown?.file).toBe("Legacy.pages");
  });

  it("refuses unknown formats", () => {
    expect(() => detectOpenFormat(file("Legacy.pages"))).toThrow(/Unsupported file type/);
    expect(() => detectOpenFormat(file("book", "application/epub+zip"))).toThrow(
      /Unsupported file type/,
    );
  });
});

describe("SAVE_FORMATS", () => {
  it("declares the full docx family, markdown, pdf, rtf, html, txt, and odt", () => {
    expect(Object.keys(SAVE_FORMATS).sort()).toEqual([
      "docm",
      "docx",
      "dotm",
      "dotx",
      "html",
      "markdown",
      "odt",
      "pdf",
      "rtf",
      "txt",
    ]);
  });

  it("carries the bare MIME + extension the save picker needs", () => {
    expect(SAVE_FORMATS.docx).toEqual({
      description: "Word Document",
      mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      ext: ".docx",
    });
    expect(SAVE_FORMATS.docm).toEqual({
      description: "Word Macro-Enabled Document",
      mime: "application/vnd.ms-word.document.macroEnabled.12",
      ext: ".docm",
    });
    expect(SAVE_FORMATS.dotx).toEqual({
      description: "Word Template",
      mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
      ext: ".dotx",
    });
    expect(SAVE_FORMATS.dotm).toEqual({
      description: "Word Macro-Enabled Template",
      mime: "application/vnd.ms-word.template.macroEnabled.12",
      ext: ".dotm",
    });
    expect(SAVE_FORMATS.rtf).toEqual({
      description: "Rich Text Format",
      mime: "application/rtf",
      ext: ".rtf",
    });
    expect(SAVE_FORMATS.html).toEqual({
      description: "Web Page",
      mime: "text/html",
      ext: ".html",
    });
    expect(SAVE_FORMATS.txt).toEqual({
      description: "Plain Text",
      mime: "text/plain",
      ext: ".txt",
    });
    expect(SAVE_FORMATS.odt).toEqual({
      description: "OpenDocument Text",
      mime: "application/vnd.oasis.opendocument.text",
      ext: ".odt",
    });
    // showSaveFilePicker rejects accept MIME keys carrying parameters.
    for (const cfg of Object.values(SAVE_FORMATS)) {
      expect(cfg.mime).not.toContain(";");
      expect(cfg.ext.startsWith(".")).toBe(true);
    }
  });
});

describe("suggestedFileName", () => {
  it("swaps any known document extension for the target format's", () => {
    expect(suggestedFileName("Report.docx", SAVE_FORMATS.dotx)).toBe("Report.dotx");
    expect(suggestedFileName("Macros.docm", SAVE_FORMATS.docx)).toBe("Macros.docx");
    expect(suggestedFileName("Letter.dotx", SAVE_FORMATS.docm)).toBe("Letter.docm");
    expect(suggestedFileName("Notes.md", SAVE_FORMATS.docx)).toBe("Notes.docx");
    expect(suggestedFileName("Document.docx", SAVE_FORMATS.dotm)).toBe("Document.dotm");
    expect(suggestedFileName("Draft.rtf", SAVE_FORMATS.odt)).toBe("Draft.odt");
    expect(suggestedFileName("Page.html", SAVE_FORMATS.txt)).toBe("Page.txt");
  });

  it("keeps names without a known extension (and the localized default)", () => {
    expect(suggestedFileName("无标题", SAVE_FORMATS.dotx)).toBe("无标题.dotx");
    expect(suggestedFileName("Report", SAVE_FORMATS.docx)).toBe("Report.docx");
  });
});
