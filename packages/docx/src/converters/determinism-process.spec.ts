// Fresh-process determinism: the same generation runs in brand-new Node
// processes (see determinism-child.mjs) and must produce identical bytes, both
// on the native ZIP writer and on the pure-JS (fflate) fallback whose fixed
// mtime normalization the native writer would otherwise shadow.
//
// The tests run sequentially and spawn their children one at a time so the
// suite's concurrent shaping benchmark is not starved of CPU.

import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const child = fileURLToPath(new URL("../../tests/determinism-child.mjs", import.meta.url));

function runChild(forceJsDeflate = false, dump = false): Buffer {
  return execFileSync(process.execPath, dump ? [child, "--dump"] : [child], {
    encoding: "buffer",
    maxBuffer: 64 * 1024 * 1024,
    env: {
      ...process.env,
      ...(forceJsDeflate ? { DOCEN_FORCE_JS_DEFLATE: "1" } : {}),
    },
  });
}

function entryHeaderDates(bytes: Uint8Array): Array<{ name: string; time: number; date: number }> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.length - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
  expect(eocd, "ZIP EOCD").toBeGreaterThanOrEqual(0);
  const count = view.getUint16(eocd + 10, true);
  let cursor = view.getUint32(eocd + 16, true);
  const entries: Array<{ name: string; time: number; date: number }> = [];
  for (let index = 0; index < count; index++) {
    const time = view.getUint16(cursor + 12, true);
    const date = view.getUint16(cursor + 14, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const name = new TextDecoder().decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength));
    entries.push({ name, time, date });
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

describe.sequential("deterministic generation across fresh processes", () => {
  it("two fresh child processes hash the same bytes on both ZIP paths", { timeout: 60_000 }, () => {
    const native1 = runChild().toString("utf8").trim();
    const native2 = runChild().toString("utf8").trim();
    expect(native1).toMatch(/^[a-f0-9]{64}$/);
    expect(native2).toBe(native1);

    const js1 = runChild(true).toString("utf8").trim();
    const js2 = runChild(true).toString("utf8").trim();
    expect(js2).toBe(js1);
  });

  it(
    "pins JS-path ZIP timestamps to the epoch and keeps entry order stable",
    { timeout: 60_000 },
    () => {
      const first = entryHeaderDates(runChild(true, true));
      expect(first.length).toBeGreaterThan(0);
      for (const entry of first) {
        // DOS time 0, DOS date 1980-01-01 (1980<<9 | 1<<5 | 1 === 0x21).
        expect(entry.time, entry.name).toBe(0);
        expect(entry.date, entry.name).toBe(0x21);
      }
      const second = entryHeaderDates(runChild(true, true)).map((entry) => entry.name);
      expect(second).toEqual(first.map((entry) => entry.name));
    },
  );
});
