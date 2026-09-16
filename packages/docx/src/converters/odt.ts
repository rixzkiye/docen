import type { JSONContent } from "../core";

// CRC32 Table
const CRC_TABLE = new Uint32Array(256);
for (let i = 0; i < 256; i++) {
  let c = i;
  for (let k = 0; k < 8; k++) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  CRC_TABLE[i] = c;
}

function crc32(buf: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = CRC_TABLE[(c ^ buf[i]!) & 0xff]! ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

interface ZipEntry {
  name: string;
  data: Uint8Array;
}

/**
 * Creates an uncompressed ZIP archive (Method 0: STORE), which is fully valid
 * and compatible with all ZIP / ODF readers.
 */
function createZip(entries: ZipEntry[]): Uint8Array {
  const enc = new TextEncoder();
  let totalSize = 0;
  const localHeaders: { entry: ZipEntry; nameBytes: Uint8Array; offset: number; crc: number }[] =
    [];

  for (const entry of entries) {
    const nameBytes = enc.encode(entry.name);
    const offset = totalSize;
    const crc = crc32(entry.data);
    localHeaders.push({ entry, nameBytes, offset, crc });
    totalSize += 30 + nameBytes.length + entry.data.length;
  }

  const centralDirOffset = totalSize;
  let centralDirSize = 0;

  for (const h of localHeaders) {
    centralDirSize += 46 + h.nameBytes.length;
  }

  totalSize += centralDirSize + 22; // + End of Central Directory

  const out = new Uint8Array(totalSize);
  const view = new DataView(out.buffer);
  let pos = 0;

  // 1. Write Local File Headers + Data
  for (const h of localHeaders) {
    view.setUint32(pos, 0x04034b50, true); // Local file header signature
    view.setUint16(pos + 4, 20, true); // Version needed to extract (2.0)
    view.setUint16(pos + 6, 0, true); // General purpose bit flag
    view.setUint16(pos + 8, 0, true); // Compression method (0 = STORE)
    view.setUint16(pos + 10, 0, true); // Last mod file time
    view.setUint16(pos + 12, 0, true); // Last mod file date
    view.setUint32(pos + 14, h.crc, true); // CRC-32
    view.setUint32(pos + 18, h.entry.data.length, true); // Compressed size
    view.setUint32(pos + 22, h.entry.data.length, true); // Uncompressed size
    view.setUint16(pos + 26, h.nameBytes.length, true); // File name length
    view.setUint16(pos + 28, 0, true); // Extra field length
    pos += 30;

    out.set(h.nameBytes, pos);
    pos += h.nameBytes.length;

    out.set(h.entry.data, pos);
    pos += h.entry.data.length;
  }

  // 2. Write Central Directory Entries
  for (const h of localHeaders) {
    view.setUint32(pos, 0x02014b50, true); // Central directory file header signature
    view.setUint16(pos + 4, 20, true); // Version made by
    view.setUint16(pos + 6, 20, true); // Version needed to extract
    view.setUint16(pos + 8, 0, true); // General purpose bit flag
    view.setUint16(pos + 10, 0, true); // Compression method
    view.setUint16(pos + 12, 0, true); // Last mod file time
    view.setUint16(pos + 14, 0, true); // Last mod file date
    view.setUint32(pos + 16, h.crc, true); // CRC-32
    view.setUint32(pos + 20, h.entry.data.length, true); // Compressed size
    view.setUint32(pos + 24, h.entry.data.length, true); // Uncompressed size
    view.setUint16(pos + 28, h.nameBytes.length, true); // File name length
    view.setUint16(pos + 30, 0, true); // Extra field length
    view.setUint16(pos + 32, 0, true); // File comment length
    view.setUint16(pos + 34, 0, true); // Disk number start
    view.setUint16(pos + 36, 0, true); // Internal file attributes
    view.setUint32(pos + 38, 0, true); // External file attributes
    view.setUint32(pos + 42, h.offset, true); // Relative offset of local header
    pos += 46;

    out.set(h.nameBytes, pos);
    pos += h.nameBytes.length;
  }

  // 3. Write End of Central Directory Record
  view.setUint32(pos, 0x06054b50, true); // End of central dir signature
  view.setUint16(pos + 4, 0, true); // Number of this disk
  view.setUint16(pos + 6, 0, true); // Number of the disk with start of central directory
  view.setUint16(pos + 8, localHeaders.length, true); // Total entries in central directory on this disk
  view.setUint16(pos + 10, localHeaders.length, true); // Total entries in central directory
  view.setUint32(pos + 12, centralDirSize, true); // Size of central directory
  view.setUint32(pos + 16, centralDirOffset, true); // Offset of start of central directory
  view.setUint16(pos + 20, 0, true); // ZIP file comment length

  return out;
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Generate an OpenDocument Text (.odt) package from Tiptap JSONContent.
 */
export async function generateODT(doc: JSONContent | JSONContent[]): Promise<Uint8Array> {
  const enc = new TextEncoder();
  const nodes = Array.isArray(doc) ? doc : (doc.content ?? []);

  const renderInline = (node: JSONContent): string => {
    if (node.type === "hardBreak") return "<text:line-break/>";
    if (node.type !== "text" || !node.text) return "";

    const text = escapeXml(node.text);
    if (!node.marks || node.marks.length === 0) return text;

    let style = "";
    for (const m of node.marks) {
      if (m.type === "bold") style = "Bold";
      else if (m.type === "italic") style = "Italic";
      else if (m.type === "underline") style = "Underline";
    }

    if (style) {
      return `<text:span text:style-name="${style}">${text}</text:span>`;
    }
    return text;
  };

  const renderBlock = (node: JSONContent): string => {
    if (node.type === "paragraph") {
      let styleName = "";
      if (node.attrs?.textAlign === "center") styleName = ' text:style-name="Center"';
      else if (node.attrs?.textAlign === "right") styleName = ' text:style-name="Right"';
      const content = (node.content ?? []).map(renderInline).join("");
      return `<text:p${styleName}>${content}</text:p>`;
    }

    if (node.type === "heading") {
      const level = Math.min(6, Math.max(1, Number(node.attrs?.level) || 1));
      const content = (node.content ?? []).map(renderInline).join("");
      return `<text:h text:outline-level="${level}">${content}</text:h>`;
    }

    if (node.type === "bulletList" || node.type === "orderedList") {
      const items = (node.content ?? [])
        .map((li) => {
          const pNodes = (li.content ?? []).filter((c) => c.type === "paragraph");
          const body =
            pNodes.length > 0
              ? pNodes.map((p) => renderBlock(p)).join("")
              : `<text:p>${(li.content ?? []).map(renderInline).join("")}</text:p>`;
          return `<text:list-item>${body}</text:list-item>`;
        })
        .join("");
      return `<text:list>${items}</text:list>`;
    }

    if (node.type === "table") {
      const rows = (node.content ?? [])
        .map((row) => {
          const cells = (row.content ?? [])
            .map((cell) => {
              const content = (cell.content ?? []).map(renderBlock).join("");
              return `<table:table-cell>${content}</table:table-cell>`;
            })
            .join("");
          return `<table:table-row>${cells}</table:table-row>`;
        })
        .join("");
      return `<table:table>${rows}</table:table>`;
    }

    if (node.content) {
      return node.content.map(renderBlock).join("");
    }

    return "";
  };

  const bodyXml = nodes.map(renderBlock).join("\n");

  const contentXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  office:version="1.3">
  <office:automatic-styles>
    <style:style style:name="Bold" style:family="text">
      <style:text-properties fo:font-weight="bold"/>
    </style:style>
    <style:style style:name="Italic" style:family="text">
      <style:text-properties fo:font-style="italic"/>
    </style:style>
    <style:style style:name="Underline" style:family="text">
      <style:text-properties style:text-underline-style="solid"/>
    </style:style>
    <style:style style:name="Center" style:family="paragraph">
      <style:paragraph-properties fo:text-align="center"/>
    </style:style>
    <style:style style:name="Right" style:family="paragraph">
      <style:paragraph-properties fo:text-align="end"/>
    </style:style>
  </office:automatic-styles>
  <office:body>
    <office:text>
      ${bodyXml}
    </office:text>
  </office:body>
</office:document-content>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  office:version="1.3">
  <office:styles>
    <style:default-style style:family="paragraph">
      <style:text-properties fo:font-size="11pt" style:font-name="Calibri"/>
    </style:default-style>
  </office:styles>
</office:document-styles>`;

  const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">
  <manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`;

  const entries: ZipEntry[] = [
    { name: "mimetype", data: enc.encode("application/vnd.oasis.opendocument.text") },
    { name: "META-INF/manifest.xml", data: enc.encode(manifestXml) },
    { name: "styles.xml", data: enc.encode(stylesXml) },
    { name: "content.xml", data: enc.encode(contentXml) },
  ];

  return createZip(entries);
}
