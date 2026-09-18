import { toUint8Array } from "@office-open/core";
import { Inflate } from "fflate";

/**
 * Untrusted DOCX admission guard.
 *
 * `parseDOCX`/`parseDOCXSync` receive arbitrary user files, so the ZIP package
 * is validated before office-open ever inflates it: entry count/paths/sizes,
 * compression method + ratio, declared AND actual inflated sizes (a zip bomb
 * can lie about its declared size, so every DEFLATE payload is streamed
 * through a capped inflater), XML part budgets (depth/nodes/attributes/text)
 * and per-media limits.
 *
 * Limits mirror the Akademi DOCX importer's admission contract
 * (`src/services/docx-importer/ooxml-admission.ts`: 32 MiB archive, 512
 * entries, 16 MiB/entry, 64 MiB aggregate, ratio 100, 64 XML depth, 100k
 * nodes, 200k attribute, 8 MiB attribute/text budgets, 1024 relationships)
 * plus owned-media caps (8 MiB per image, 32 MiB total) and an entry cap for
 * `word/media/*`. Docen adds `mediaEntries`, since the importer rejected
 * active content before media counting.
 */

/** Admission limits for untrusted DOCX archives. */
export interface ArchiveLimits {
  readonly archiveBytes: number;
  readonly entries: number;
  readonly filenameBytes: number;
  readonly entryUncompressedBytes: number;
  readonly aggregateUncompressedBytes: number;
  readonly compressionRatio: number;
  readonly xmlBytes: number;
  readonly xmlAggregateBytes: number;
  readonly xmlNodes: number;
  readonly xmlAttributes: number;
  readonly xmlAttributeBytes: number;
  readonly xmlDepth: number;
  readonly xmlTextBytes: number;
  readonly relationships: number;
  readonly externalTargetBytes: number;
  readonly mediaEntries: number;
  readonly mediaEntryBytes: number;
  readonly mediaAggregateBytes: number;
}

/** The admission limits applied to every `parseDOCX` input. */
export const ARCHIVE_LIMITS: Readonly<ArchiveLimits> = Object.freeze({
  archiveBytes: 32 * 1024 * 1024,
  entries: 512,
  filenameBytes: 255,
  entryUncompressedBytes: 16 * 1024 * 1024,
  aggregateUncompressedBytes: 64 * 1024 * 1024,
  compressionRatio: 100,
  xmlBytes: 4 * 1024 * 1024,
  xmlAggregateBytes: 16 * 1024 * 1024,
  xmlNodes: 100_000,
  xmlAttributes: 200_000,
  xmlAttributeBytes: 8 * 1024 * 1024,
  xmlDepth: 64,
  xmlTextBytes: 8 * 1024 * 1024,
  relationships: 1_024,
  externalTargetBytes: 2_048,
  mediaEntries: 256,
  mediaEntryBytes: 8 * 1024 * 1024,
  mediaAggregateBytes: 32 * 1024 * 1024,
});

export type ArchiveRejectionCode =
  | "ZIP_MAGIC_INVALID"
  | "ZIP_CENTRAL_DIRECTORY_INVALID"
  | "ZIP_MULTIDISK_UNSUPPORTED"
  | "ZIP64_UNSUPPORTED"
  | "ZIP_ENTRY_LIMIT_EXCEEDED"
  | "ZIP_FILENAME_INVALID"
  | "ZIP_DUPLICATE_PATH"
  | "ZIP_ENCRYPTED_ENTRY"
  | "ZIP_COMPRESSION_UNSUPPORTED"
  | "ZIP_ENTRY_SIZE_EXCEEDED"
  | "ZIP_AGGREGATE_SIZE_EXCEEDED"
  | "ZIP_COMPRESSION_RATIO_EXCEEDED"
  | "ZIP_LOCAL_HEADER_INVALID"
  | "ZIP_PAYLOAD_INVALID"
  | "ZIP_CRC_MISMATCH"
  | "ZIP_MEDIA_LIMIT_EXCEEDED"
  | "OOXML_XML_SIZE_EXCEEDED"
  | "OOXML_XML_INVALID"
  | "OOXML_XML_NODE_LIMIT_EXCEEDED"
  | "OOXML_XML_ATTRIBUTE_LIMIT_EXCEEDED"
  | "OOXML_XML_DEPTH_EXCEEDED"
  | "OOXML_XML_TEXT_LIMIT_EXCEEDED"
  | "OOXML_DTD_FORBIDDEN"
  | "OOXML_RELATIONSHIP_LIMIT_EXCEEDED";

/** Rejection thrown for any archive that violates the admission limits. */
export class ArchiveRejection extends Error {
  constructor(
    readonly code: ArchiveRejectionCode,
    message: string,
    readonly part?: string,
  ) {
    super(message);
    this.name = "ArchiveRejection";
  }
}

interface ZipEntry {
  readonly path: string;
  readonly normalizedPath: string;
  readonly isDirectory: boolean;
  readonly method: number;
  readonly crc32: number;
  readonly compressedBytes: number;
  readonly uncompressedBytes: number;
  readonly dataStart: number;
  readonly dataEnd: number;
}

interface XmlBudget {
  bytes: number;
  nodes: number;
  attributes: number;
  attributeBytes: number;
  textBytes: number;
  relationships: number;
}

function reject(code: ArchiveRejectionCode, message: string, part?: string): never {
  throw new ArchiveRejection(code, message, part);
}

const u16 = (b: Uint8Array, o: number, code: ArchiveRejectionCode, msg: string): number => {
  if (o < 0 || o + 2 > b.length) reject(code, msg);
  return b[o]! | (b[o + 1]! << 8);
};
const u32 = (b: Uint8Array, o: number, code: ArchiveRejectionCode, msg: string): number => {
  if (o < 0 || o + 4 > b.length) reject(code, msg);
  return (b[o]! | (b[o + 1]! << 8) | (b[o + 2]! << 16) | (b[o + 3]! << 24)) >>> 0;
};

function normalizeArchivePath(value: string, limits: ArchiveLimits): string {
  if (
    !value ||
    utf8Length(value) > limits.filenameBytes ||
    value.startsWith("/") ||
    value.includes("\\") ||
    // eslint-disable-next-line no-control-regex -- control characters are invalid ZIP path bytes
    /[\0-\x1f\x7f]/.test(value)
  ) {
    reject("ZIP_FILENAME_INVALID", "ZIP entry path is invalid");
  }
  const path = value.endsWith("/") ? value.slice(0, -1) : value;
  if (!path || path.split("/").some((segment) => !segment || segment === "." || segment === "..")) {
    reject("ZIP_FILENAME_INVALID", "ZIP entry path contains an absolute or traversal segment");
  }
  return path.normalize("NFC").toLocaleLowerCase("en-US");
}

function utf8Length(value: string): number {
  return new TextEncoder().encode(value).length;
}

function findEndOfCentralDirectory(input: Uint8Array): number {
  const min = Math.max(0, input.length - 65_557);
  for (let offset = input.length - 22; offset >= min; offset--) {
    if (u32(input, offset, "ZIP_CENTRAL_DIRECTORY_INVALID", "EOCD truncated") !== 0x06054b50)
      continue;
    const commentLength = u16(
      input,
      offset + 20,
      "ZIP_CENTRAL_DIRECTORY_INVALID",
      "EOCD truncated",
    );
    if (offset + 22 + commentLength === input.length) return offset;
  }
  reject("ZIP_CENTRAL_DIRECTORY_INVALID", "ZIP end of central directory is missing");
}

/** Walk and validate the central directory; throws on any structural violation. */
function centralDirectory(input: Uint8Array, limits: ArchiveLimits): ZipEntry[] {
  if (input.length < 4 || input[0] !== 0x50 || input[1] !== 0x4b) {
    reject("ZIP_MAGIC_INVALID", "DOCX is not a PK ZIP archive");
  }
  if (input.length > limits.archiveBytes) {
    reject(
      "ZIP_ENTRY_SIZE_EXCEEDED",
      `DOCX archive exceeds the ${limits.archiveBytes}-byte compressed input limit`,
    );
  }
  const eocd = findEndOfCentralDirectory(input);
  const disk = u16(input, eocd + 4, "ZIP_CENTRAL_DIRECTORY_INVALID", "EOCD malformed");
  const directoryDisk = u16(input, eocd + 6, "ZIP_CENTRAL_DIRECTORY_INVALID", "EOCD malformed");
  const entriesOnDisk = u16(input, eocd + 8, "ZIP_CENTRAL_DIRECTORY_INVALID", "EOCD malformed");
  const count = u16(input, eocd + 10, "ZIP_CENTRAL_DIRECTORY_INVALID", "EOCD malformed");
  const directorySize = u32(input, eocd + 12, "ZIP_CENTRAL_DIRECTORY_INVALID", "EOCD malformed");
  const directoryOffset = u32(input, eocd + 16, "ZIP_CENTRAL_DIRECTORY_INVALID", "EOCD malformed");
  const commentLength = u16(input, eocd + 20, "ZIP_CENTRAL_DIRECTORY_INVALID", "EOCD malformed");
  if (disk !== 0 || directoryDisk !== 0 || entriesOnDisk !== count) {
    reject("ZIP_MULTIDISK_UNSUPPORTED", "Multi-disk ZIP archives are unsupported");
  }
  if (count === 0xffff || directorySize === 0xffffffff || directoryOffset === 0xffffffff) {
    reject("ZIP64_UNSUPPORTED", "ZIP64 DOCX archives are unsupported");
  }
  if (count > limits.entries) {
    reject("ZIP_ENTRY_LIMIT_EXCEEDED", "DOCX archive entry count exceeds the limit");
  }
  if (eocd + 22 + commentLength !== input.length || directoryOffset + directorySize > eocd) {
    reject("ZIP_CENTRAL_DIRECTORY_INVALID", "ZIP central directory bounds are invalid");
  }

  const entries: ZipEntry[] = [];
  const paths = new Set<string>();
  const ranges: Array<{ offset: number; end: number; path: string }> = [];
  let cursor = directoryOffset;
  let declaredAggregate = 0;
  for (let index = 0; index < count; index++) {
    if (
      u32(input, cursor, "ZIP_CENTRAL_DIRECTORY_INVALID", "central entry malformed") !== 0x02014b50
    ) {
      reject("ZIP_CENTRAL_DIRECTORY_INVALID", "ZIP central directory entry signature is invalid");
    }
    const flags = u16(
      input,
      cursor + 8,
      "ZIP_CENTRAL_DIRECTORY_INVALID",
      "central entry truncated",
    );
    const method = u16(
      input,
      cursor + 10,
      "ZIP_CENTRAL_DIRECTORY_INVALID",
      "central entry truncated",
    );
    const crc32 = u32(
      input,
      cursor + 16,
      "ZIP_CENTRAL_DIRECTORY_INVALID",
      "central entry truncated",
    );
    const compressedBytes = u32(
      input,
      cursor + 20,
      "ZIP_CENTRAL_DIRECTORY_INVALID",
      "central entry truncated",
    );
    const uncompressedBytes = u32(
      input,
      cursor + 24,
      "ZIP_CENTRAL_DIRECTORY_INVALID",
      "central entry truncated",
    );
    const nameLength = u16(
      input,
      cursor + 28,
      "ZIP_CENTRAL_DIRECTORY_INVALID",
      "central entry truncated",
    );
    const extraLength = u16(
      input,
      cursor + 30,
      "ZIP_CENTRAL_DIRECTORY_INVALID",
      "central entry truncated",
    );
    const entryCommentLength = u16(
      input,
      cursor + 32,
      "ZIP_CENTRAL_DIRECTORY_INVALID",
      "central entry truncated",
    );
    const localHeaderOffset = u32(
      input,
      cursor + 42,
      "ZIP_CENTRAL_DIRECTORY_INVALID",
      "central entry truncated",
    );
    const end = cursor + 46 + nameLength + extraLength + entryCommentLength;
    if (end > directoryOffset + directorySize) {
      reject("ZIP_CENTRAL_DIRECTORY_INVALID", "ZIP central entry exceeds central directory bounds");
    }
    const path = utf8DecodeStrict(
      input.subarray(cursor + 46, cursor + 46 + nameLength),
      "ZIP_FILENAME_INVALID",
      "ZIP entry filename is not valid UTF-8",
    );
    const normalizedPath = normalizeArchivePath(path, limits);
    if (paths.has(normalizedPath)) {
      reject("ZIP_DUPLICATE_PATH", `ZIP archive has duplicate path: ${path}`);
    }
    paths.add(normalizedPath);
    if ((flags & 0x41) !== 0) {
      reject("ZIP_ENCRYPTED_ENTRY", `Encrypted ZIP entry is unsupported: ${path}`, path);
    }
    if (method !== 0 && method !== 8) {
      reject("ZIP_COMPRESSION_UNSUPPORTED", `Unsupported ZIP compression method in ${path}`, path);
    }
    if (uncompressedBytes > limits.entryUncompressedBytes) {
      reject("ZIP_ENTRY_SIZE_EXCEEDED", `ZIP entry exceeds uncompressed size limit: ${path}`, path);
    }
    if (
      uncompressedBytes > 0 &&
      (compressedBytes === 0 || uncompressedBytes / compressedBytes > limits.compressionRatio)
    ) {
      reject(
        "ZIP_COMPRESSION_RATIO_EXCEEDED",
        `ZIP entry compression ratio exceeds limit: ${path}`,
        path,
      );
    }
    const isDirectory = path.endsWith("/");
    if (isDirectory && (compressedBytes !== 0 || uncompressedBytes !== 0)) {
      reject("ZIP_LOCAL_HEADER_INVALID", `ZIP directory has payload bytes: ${path}`, path);
    }
    declaredAggregate += uncompressedBytes;
    if (declaredAggregate > limits.aggregateUncompressedBytes) {
      reject(
        "ZIP_AGGREGATE_SIZE_EXCEEDED",
        "DOCX declared uncompressed size exceeds aggregate limit",
      );
    }
    const range = validateLocalHeader(
      input,
      localHeaderOffset,
      flags,
      method,
      crc32,
      compressedBytes,
      uncompressedBytes,
      path,
      directoryOffset,
    );
    ranges.push({ offset: localHeaderOffset, end: range.dataEnd, path });
    entries.push({
      path,
      normalizedPath,
      isDirectory,
      method,
      crc32,
      compressedBytes,
      uncompressedBytes,
      dataStart: range.dataStart,
      dataEnd: range.dataEnd,
    });
    cursor = end;
  }
  if (cursor !== directoryOffset + directorySize) {
    reject(
      "ZIP_CENTRAL_DIRECTORY_INVALID",
      "ZIP central directory size does not match its entries",
    );
  }
  let previous: { offset: number; end: number; path: string } | undefined;
  for (const range of ranges.sort((a, b) => a.offset - b.offset)) {
    if (previous && range.offset < previous.end) {
      reject(
        "ZIP_LOCAL_HEADER_INVALID",
        `ZIP local headers or payloads overlap: ${range.path}`,
        range.path,
      );
    }
    previous = range;
  }
  return entries;
}

function validateLocalHeader(
  input: Uint8Array,
  offset: number,
  flags: number,
  method: number,
  crc32: number,
  compressedBytes: number,
  uncompressedBytes: number,
  path: string,
  centralDirectoryOffset: number,
): { dataStart: number; dataEnd: number } {
  if (
    offset >= centralDirectoryOffset ||
    u32(input, offset, "ZIP_LOCAL_HEADER_INVALID", "local header truncated") !== 0x04034b50
  ) {
    reject("ZIP_LOCAL_HEADER_INVALID", `ZIP local header signature is invalid for ${path}`, path);
  }
  const localFlags = u16(input, offset + 6, "ZIP_LOCAL_HEADER_INVALID", "local header truncated");
  const localMethod = u16(input, offset + 8, "ZIP_LOCAL_HEADER_INVALID", "local header truncated");
  const localCrc = u32(input, offset + 14, "ZIP_LOCAL_HEADER_INVALID", "local header truncated");
  const localCompressed = u32(
    input,
    offset + 18,
    "ZIP_LOCAL_HEADER_INVALID",
    "local header truncated",
  );
  const localUncompressed = u32(
    input,
    offset + 22,
    "ZIP_LOCAL_HEADER_INVALID",
    "local header truncated",
  );
  const nameLength = u16(input, offset + 26, "ZIP_LOCAL_HEADER_INVALID", "local header truncated");
  const extraLength = u16(input, offset + 28, "ZIP_LOCAL_HEADER_INVALID", "local header truncated");
  const nameEnd = offset + 30 + nameLength;
  if (nameEnd > centralDirectoryOffset) {
    reject("ZIP_LOCAL_HEADER_INVALID", `ZIP local header name exceeds payload area: ${path}`, path);
  }
  const localName = utf8DecodeStrict(
    input.subarray(offset + 30, nameEnd),
    "ZIP_LOCAL_HEADER_INVALID",
    "ZIP local entry filename is invalid",
  );
  if (localName !== path || localFlags !== flags || localMethod !== method) {
    reject(
      "ZIP_LOCAL_HEADER_INVALID",
      `ZIP local header does not match central directory for ${path}`,
      path,
    );
  }
  // Bit 3 marks a streaming write: the local sizes/CRC are placeholders (zero)
  // and the real values live in a data descriptor after the payload; the
  // central directory stays authoritative for bounds, inflation and CRC. The
  // descriptor must still agree with it (see readDataDescriptor), and a
  // writer that fills both must fill them exactly.
  const streamed = (flags & 0x08) !== 0;
  if (
    !streamed &&
    (localCrc !== crc32 ||
      localCompressed !== compressedBytes ||
      localUncompressed !== uncompressedBytes)
  ) {
    reject(
      "ZIP_LOCAL_HEADER_INVALID",
      `ZIP local header does not match central directory for ${path}`,
      path,
    );
  }
  const dataStart = nameEnd + extraLength;
  if (dataStart > centralDirectoryOffset || compressedBytes > centralDirectoryOffset - dataStart) {
    reject("ZIP_LOCAL_HEADER_INVALID", `ZIP payload bounds are invalid for ${path}`, path);
  }
  const dataEnd = dataStart + compressedBytes;
  if (streamed) {
    const descriptor = readDataDescriptor(input, dataEnd, centralDirectoryOffset, path);
    if (
      descriptor.crc32 !== crc32 ||
      descriptor.compressedBytes !== compressedBytes ||
      descriptor.uncompressedBytes !== uncompressedBytes
    ) {
      reject(
        "ZIP_LOCAL_HEADER_INVALID",
        `ZIP data descriptor does not match central directory for ${path}`,
        path,
      );
    }
  }
  return { dataStart, dataEnd };
}

/**
 * Read and return the data descriptor that follows a streaming entry's
 * payload. The 0x08074b50 signature is optional (APPNOTE 4.3.9.3); writers
 * that emit it (LibreOffice, Info-ZIP) and writers that omit it are both
 * accepted, with the descriptor required to fit before the central directory.
 */
function readDataDescriptor(
  input: Uint8Array,
  offset: number,
  centralDirectoryOffset: number,
  path: string,
): { crc32: number; compressedBytes: number; uncompressedBytes: number } {
  const signed =
    offset + 4 <= centralDirectoryOffset &&
    u32(input, offset, "ZIP_LOCAL_HEADER_INVALID", "ZIP data descriptor is truncated") ===
      0x08074b50;
  const cursor = signed ? offset + 4 : offset;
  if (cursor + 12 > centralDirectoryOffset) {
    reject("ZIP_LOCAL_HEADER_INVALID", `ZIP data descriptor is truncated for ${path}`, path);
  }
  return {
    crc32: u32(input, cursor, "ZIP_LOCAL_HEADER_INVALID", "ZIP data descriptor is truncated"),
    compressedBytes: u32(
      input,
      cursor + 4,
      "ZIP_LOCAL_HEADER_INVALID",
      "ZIP data descriptor is truncated",
    ),
    uncompressedBytes: u32(
      input,
      cursor + 8,
      "ZIP_LOCAL_HEADER_INVALID",
      "ZIP data descriptor is truncated",
    ),
  };
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < table.length; index++) {
    let value = index;
    for (let bit = 0; bit < 8; bit++) value = value & 1 ? (value >>> 1) ^ 0xedb88320 : value >>> 1;
    table[index] = value >>> 0;
  }
  return table;
})();

function crc32Update(value: number, input: Uint8Array): number {
  let next = value;
  for (let i = 0; i < input.length; i++)
    next = CRC_TABLE[(next ^ input[i]!) & 0xff]! ^ (next >>> 8);
  return next;
}

function utf8DecodeStrict(
  bytes: Uint8Array,
  code: ArchiveRejectionCode,
  message: string,
  part?: string,
): string {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    reject(code, message, part);
  }
}

/** True for entries office-open parses as XML (its readers only accept these). */
function isXmlPart(entry: ZipEntry): boolean {
  const path = entry.normalizedPath;
  return path === "[content_types].xml" || path.endsWith(".xml") || path.endsWith(".rels");
}

function isMediaPart(entry: ZipEntry): boolean {
  return entry.normalizedPath.startsWith("word/media/");
}

/**
 * Stream-inflate one entry with hard caps on output, CRC and, for XML parts,
 * decode + scan the retained bytes against the XML budgets.
 */
function validatePayload(
  input: Uint8Array,
  entry: ZipEntry,
  limits: ArchiveLimits,
  payloadBudget: { bytes: number },
  xmlBudget: XmlBudget,
  media: { count: number; entryBytes: number; totalBytes: number },
): void {
  const xml = isXmlPart(entry);
  const mediaPart = isMediaPart(entry);
  // The output cap is the tightest of the entry's own declaration, the global
  // entry cap, and the XML/media-specific budget; `capCode` reports which one
  // tripped so callers can tell a generic bomb from an XML/media budget.
  let entryCap =
    entry.uncompressedBytes > 0 ? entry.uncompressedBytes : limits.entryUncompressedBytes;
  let capCode: ArchiveRejectionCode = "ZIP_ENTRY_SIZE_EXCEEDED";
  if (xml && limits.xmlBytes < entryCap) {
    entryCap = limits.xmlBytes;
    capCode = "OOXML_XML_SIZE_EXCEEDED";
  }
  if (mediaPart && limits.mediaEntryBytes < entryCap) {
    entryCap = limits.mediaEntryBytes;
    capCode = "ZIP_MEDIA_LIMIT_EXCEEDED";
  }
  const chunks: Uint8Array[] = [];
  let outputBytes = 0;
  let checksum = 0xffffffff;
  const consume = (chunk: Uint8Array): void => {
    const next = outputBytes + chunk.length;
    // Abort at the first over-budget chunk: a DEFLATE bomb must stop the
    // inflater, not merely stop retaining output.
    if (next > entryCap) {
      reject(capCode, `Inflated entry exceeds its admission limit: ${entry.path}`, entry.path);
    }
    if (payloadBudget.bytes + chunk.length > limits.aggregateUncompressedBytes) {
      reject(
        "ZIP_AGGREGATE_SIZE_EXCEEDED",
        "DOCX actual uncompressed size exceeds aggregate limit",
      );
    }
    payloadBudget.bytes += chunk.length;
    outputBytes = next;
    checksum = crc32Update(checksum, chunk);
    if (xml) chunks.push(chunk);
  };

  const raw = input.subarray(entry.dataStart, entry.dataEnd);
  if (entry.method === 0) {
    if (entry.compressedBytes !== entry.uncompressedBytes) {
      reject(
        "ZIP_PAYLOAD_INVALID",
        `Stored ZIP entry declares mismatched sizes: ${entry.path}`,
        entry.path,
      );
    }
    consume(raw);
  } else {
    inflateDeflateBounded(raw, consume, entry.path);
  }
  if (outputBytes !== entry.uncompressedBytes) {
    reject(
      "ZIP_PAYLOAD_INVALID",
      `Inflated size does not match central directory: ${entry.path}`,
      entry.path,
    );
  }
  if ((checksum ^ 0xffffffff) >>> 0 !== entry.crc32) {
    reject("ZIP_CRC_MISMATCH", `ZIP CRC mismatch: ${entry.path}`, entry.path);
  }
  if (mediaPart) {
    media.count += 1;
    media.entryBytes = Math.max(media.entryBytes, outputBytes);
    media.totalBytes += outputBytes;
    if (media.count > limits.mediaEntries) {
      reject(
        "ZIP_MEDIA_LIMIT_EXCEEDED",
        `DOCX media entry count exceeds the limit: ${entry.path}`,
        entry.path,
      );
    }
    if (media.entryBytes > limits.mediaEntryBytes) {
      reject(
        "ZIP_MEDIA_LIMIT_EXCEEDED",
        `DOCX media entry exceeds the size limit: ${entry.path}`,
        entry.path,
      );
    }
    if (media.totalBytes > limits.mediaAggregateBytes) {
      reject("ZIP_MEDIA_LIMIT_EXCEEDED", "DOCX media exceeds the aggregate size limit", entry.path);
    }
  }
  if (!xml) return;
  scanXmlBytes(concatChunks(chunks, outputBytes), entry.path, limits, xmlBudget);
}

function concatChunks(chunks: readonly Uint8Array[], total: number): Uint8Array {
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.length;
  }
  return bytes;
}

/** Inflate a raw DEFLATE stream, forwarding bounded output chunks to `consume`. */
function inflateDeflateBounded(
  raw: Uint8Array,
  consume: (chunk: Uint8Array) => void,
  part: string,
): void {
  let failure: unknown;
  const inflate = new Inflate((chunk) => {
    if (failure !== undefined) return;
    try {
      if (chunk.length > 0) consume(chunk);
    } catch (error) {
      failure = error;
    }
  });
  // Feed in small chunks: fflate inflates exactly the pushed bytes per call,
  // so the transient output of one push is bounded (DEFLATE's ~1032:1 ceiling
  // × the chunk size) and `consume` can abort the bomb at the cap instead of
  // materializing the whole member. A single `push(raw, true)` would inflate
  // the entire stream into memory before the cap could run.
  const CHUNK = 8192;
  try {
    for (let offset = 0; offset < raw.length && failure === undefined; offset += CHUNK) {
      const end = Math.min(offset + CHUNK, raw.length);
      inflate.push(raw.subarray(offset, end), end === raw.length);
    }
  } catch (error) {
    if (failure === undefined) failure = error;
  }
  if (failure !== undefined) {
    if (failure instanceof ArchiveRejection) throw failure;
    reject("ZIP_PAYLOAD_INVALID", `Invalid DEFLATE payload: ${part}`, part);
  }
}

// ── XML budget scan ──

/**
 * Bounded XML scan (no DOM, no entity expansion): tracks element depth, node
 * and attribute counts, attribute bytes and text bytes, and rejects DTD /
 * entity declarations. It deliberately does not build an AST — office-open
 * parses the retained bytes afterwards; this pass only proves the part fits
 * the admission budget.
 */
function scanXmlBytes(
  bytes: Uint8Array,
  part: string,
  limits: ArchiveLimits,
  budget: XmlBudget,
): void {
  if (bytes.length > limits.xmlBytes) {
    reject("OOXML_XML_SIZE_EXCEEDED", `XML part exceeds size limit: ${part}`, part);
  }
  budget.bytes += bytes.length;
  if (budget.bytes > limits.xmlAggregateBytes) {
    reject("OOXML_XML_SIZE_EXCEEDED", "Aggregate XML bytes exceed admission limit", part);
  }
  const source = utf8DecodeStrict(
    bytes,
    "OOXML_XML_INVALID",
    `XML part is not valid UTF-8: ${part}`,
    part,
  );
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) {
    reject("OOXML_DTD_FORBIDDEN", `DTD or entity declaration is forbidden in ${part}`, part);
  }
  const stack: string[] = [];
  let cursor = 0;
  let rootSeen = false;
  while (cursor < source.length) {
    const start = source.indexOf("<", cursor);
    const text = source.slice(cursor, start < 0 ? source.length : start);
    if (text.length > 0) {
      budget.textBytes += utf8Length(text);
      if (budget.textBytes > limits.xmlTextBytes) {
        reject("OOXML_XML_TEXT_LIMIT_EXCEEDED", "Aggregate XML text exceeds admission limit", part);
      }
      if (stack.length === 0 && text.trim().length > 0) {
        reject("OOXML_XML_INVALID", `Unexpected XML text outside root: ${part}`, part);
      }
    }
    if (start < 0) break;
    if (source.startsWith("<!--", start)) {
      const end = source.indexOf("-->", start + 4);
      if (end < 0) reject("OOXML_XML_INVALID", `Unterminated XML comment: ${part}`, part);
      cursor = end + 3;
      continue;
    }
    if (source.startsWith("<![CDATA[", start)) {
      const end = source.indexOf("]]>", start + 9);
      if (end < 0) reject("OOXML_XML_INVALID", `Unterminated XML CDATA: ${part}`, part);
      budget.textBytes += utf8Length(source.slice(start + 9, end));
      if (budget.textBytes > limits.xmlTextBytes) {
        reject("OOXML_XML_TEXT_LIMIT_EXCEEDED", "Aggregate XML text exceeds admission limit", part);
      }
      cursor = end + 3;
      continue;
    }
    if (source.startsWith("<?", start)) {
      const end = source.indexOf("?>", start + 2);
      if (end < 0)
        reject("OOXML_XML_INVALID", `Unterminated XML processing instruction: ${part}`, part);
      cursor = end + 2;
      continue;
    }
    const end = xmlTagEnd(source, start, part);
    const tag = source.slice(start, end);
    cursor = end;
    if (tag.startsWith("</")) {
      const name = tag.slice(2, -1).trim();
      const open = stack.pop();
      if (open === undefined || open !== name) {
        reject("OOXML_XML_INVALID", `Malformed XML nesting: ${part}`, part);
      }
      continue;
    }
    if (tag.startsWith("<!")) {
      reject("OOXML_XML_INVALID", `Unsupported XML declaration: ${part}`, part);
    }
    const selfClosing = /\/\s*>$/.test(tag);
    const body = tag.slice(1, selfClosing ? tag.lastIndexOf("/") : -1).trim();
    const name = /^([^\s/>]+)/.exec(body)?.[1];
    if (!name) reject("OOXML_XML_INVALID", `Malformed XML element: ${part}`, part);
    if (stack.length + 1 > limits.xmlDepth) {
      reject("OOXML_XML_DEPTH_EXCEEDED", `XML depth exceeds limit: ${part}`, part);
    }
    budget.nodes += 1;
    if (budget.nodes > limits.xmlNodes) {
      reject(
        "OOXML_XML_NODE_LIMIT_EXCEEDED",
        "Aggregate XML node count exceeds admission limit",
        part,
      );
    }
    const attributes = parseTagAttributes(body.slice(name.length), part, limits, budget);
    if (part.endsWith(".rels") && name === "Relationship") {
      budget.relationships += 1;
      if (budget.relationships > limits.relationships) {
        reject("OOXML_RELATIONSHIP_LIMIT_EXCEEDED", "OOXML relationship count exceeds limit", part);
      }
      const target = attributes.get("Target");
      if (
        attributes.get("TargetMode") === "External" &&
        target !== undefined &&
        utf8Length(target) > limits.externalTargetBytes
      ) {
        reject("OOXML_XML_TEXT_LIMIT_EXCEEDED", "External relationship target exceeds limit", part);
      }
    }
    if (stack.length === 0) {
      if (rootSeen) reject("OOXML_XML_INVALID", `XML has multiple roots: ${part}`, part);
      rootSeen = true;
    }
    if (!selfClosing) stack.push(name);
  }
  if (!rootSeen) reject("OOXML_XML_INVALID", `XML root is missing: ${part}`, part);
  if (stack.length > 0) reject("OOXML_XML_INVALID", `XML root is unclosed: ${part}`, part);
}

function xmlTagEnd(source: string, start: number, part: string): number {
  let quote: string | undefined;
  for (let index = start + 1; index < source.length; index++) {
    const character = source[index]!;
    if (quote) {
      if (character === quote) quote = undefined;
    } else if (character === '"' || character === "'") {
      quote = character;
    } else if (character === ">") {
      return index + 1;
    }
  }
  reject("OOXML_XML_INVALID", `Unterminated XML tag: ${part}`, part);
}

function parseTagAttributes(
  source: string,
  part: string,
  limits: ArchiveLimits,
  budget: XmlBudget,
): Map<string, string> {
  const attributes = new Map<string, string>();
  const pattern = /([^\s=]+)\s*=\s*("[^"]*"|'[^']*')/g;
  let consumed = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(source))) {
    if (source.slice(consumed, match.index).trim()) {
      reject("OOXML_XML_INVALID", `Malformed XML attributes: ${part}`, part);
    }
    budget.attributes += 1;
    budget.attributeBytes += utf8Length(match[0]!);
    if (budget.attributes > limits.xmlAttributes) {
      reject("OOXML_XML_ATTRIBUTE_LIMIT_EXCEEDED", "Aggregate XML attributes exceed limit", part);
    }
    if (budget.attributeBytes > limits.xmlAttributeBytes) {
      reject(
        "OOXML_XML_ATTRIBUTE_LIMIT_EXCEEDED",
        "Aggregate XML attribute bytes exceed limit",
        part,
      );
    }
    attributes.set(match[1]!, match[2]!.slice(1, -1));
    consumed = pattern.lastIndex;
  }
  if (source.slice(consumed).trim()) {
    reject("OOXML_XML_INVALID", `Malformed XML attributes: ${part}`, part);
  }
  return attributes;
}

/**
 * Validate an untrusted DOCX archive against {@link ARCHIVE_LIMITS}.
 *
 * @throws {@link ArchiveRejection} for any limit violation (message + `code`
 * identify the failed rule; `part` names the offending ZIP entry when known).
 */
export function assertArchiveWithinLimits(
  bytes: Uint8Array,
  limits: ArchiveLimits = ARCHIVE_LIMITS,
): void {
  const entries = centralDirectory(bytes, limits);
  const payloadBudget = { bytes: 0 };
  const xmlBudget: XmlBudget = {
    bytes: 0,
    nodes: 0,
    attributes: 0,
    attributeBytes: 0,
    textBytes: 0,
    relationships: 0,
  };
  const media = { count: 0, entryBytes: 0, totalBytes: 0 };
  for (const entry of entries) {
    if (entry.isDirectory) continue;
    validatePayload(bytes, entry, limits, payloadBudget, xmlBudget, media);
  }
}

/**
 * Normalize any `parseDOCX` input to bytes, enforcing the compressed-input cap
 * while reading: a Blob/ReadableStream larger than the cap is rejected before
 * it is materialized.
 */
export async function normalizeArchiveInput(data: unknown): Promise<Uint8Array> {
  if (typeof Blob !== "undefined" && data instanceof Blob) {
    if (data.size > ARCHIVE_LIMITS.archiveBytes) {
      reject(
        "ZIP_ENTRY_SIZE_EXCEEDED",
        `DOCX archive exceeds the ${ARCHIVE_LIMITS.archiveBytes}-byte compressed input limit`,
      );
    }
    return new Uint8Array(await data.arrayBuffer());
  }
  if (typeof ReadableStream !== "undefined" && data instanceof ReadableStream) {
    const reader = data.getReader();
    const chunks: Uint8Array[] = [];
    let total = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > ARCHIVE_LIMITS.archiveBytes) {
        await reader.cancel();
        reject(
          "ZIP_ENTRY_SIZE_EXCEEDED",
          `DOCX archive exceeds the ${ARCHIVE_LIMITS.archiveBytes}-byte compressed input limit`,
        );
      }
      chunks.push(value);
    }
    return concatChunks(chunks, total);
  }
  const bytes = toUint8Array(data as never);
  if (bytes.byteLength > ARCHIVE_LIMITS.archiveBytes) {
    reject(
      "ZIP_ENTRY_SIZE_EXCEEDED",
      `DOCX archive exceeds the ${ARCHIVE_LIMITS.archiveBytes}-byte compressed input limit`,
    );
  }
  return bytes;
}

/** Sync counterpart of {@link normalizeArchiveInput} (Blob/stream inputs throw). */
export function normalizeArchiveInputSync(data: unknown): Uint8Array {
  const bytes = toUint8Array(data as never);
  if (bytes.byteLength > ARCHIVE_LIMITS.archiveBytes) {
    reject(
      "ZIP_ENTRY_SIZE_EXCEEDED",
      `DOCX archive exceeds the ${ARCHIVE_LIMITS.archiveBytes}-byte compressed input limit`,
    );
  }
  return bytes;
}
