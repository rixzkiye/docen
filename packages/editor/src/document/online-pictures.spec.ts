// Online Pictures: the URL → embedded-picture resolver. The data-URL case
// runs the real public fetch handler; network failures/downgrades are driven
// through the injectable handler so the failure taxonomy is asserted, not the
// network stack.
import { describe, expect, it, vi } from "vitest";

import { ONLINE_PICTURE_MAX_BYTES, resolveOnlinePicture } from "./online-pictures";

/** A real 1×1 PNG. */
const PNG_1PX =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const pngBytes = (): Uint8Array => {
  const binary = atob(PNG_1PX.split(",")[1]!);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
};

describe("resolveOnlinePicture", () => {
  it("accepts a real data: URL through the public fetch handler", async () => {
    const result = await resolveOnlinePicture(PNG_1PX);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.picture.src.startsWith("data:image/png;base64,")).toBe(true);
    expect(result.picture.alt).toBe("");
    // The embedded bytes survive the re-encode unchanged.
    expect(btoa(atob(result.picture.src.split(",")[1]!))).toBe(PNG_1PX.split(",")[1]!);
  });

  it("downloads an http URL, sniffs the type, and names the alt from the path", async () => {
    const fetcher = vi.fn(async (_url: string) => pngBytes());
    const result = await resolveOnlinePicture("https://example.com/photos/team%20photo.png?x=1", {
      fetch: fetcher,
    });
    expect(fetcher).toHaveBeenCalledWith("https://example.com/photos/team%20photo.png?x=1");
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    // Extension-independent: the MIME comes from the magic bytes.
    expect(result.picture.src.startsWith("data:image/png;base64,")).toBe(true);
    expect(result.picture.alt).toBe("team photo");
  });

  it("probes the natural size when a probe is provided", async () => {
    const result = await resolveOnlinePicture(PNG_1PX, {
      probeSize: async () => ({ w: 640, h: 480 }),
    });
    expect(result).toEqual({
      ok: true,
      picture: {
        src: expect.stringContaining("data:image/png;base64,"),
        width: 640,
        height: 480,
        alt: "",
      },
    });
  });

  it("maps a denied/failed network request to fetch-failed", async () => {
    const result = await resolveOnlinePicture("https://blocked.example/img.png", {
      fetch: async () => {
        throw new TypeError("Failed to fetch");
      },
    });
    expect(result).toEqual({ ok: false, reason: "fetch-failed" });
  });

  it("refuses unsupported schemes and malformed URLs without fetching", async () => {
    const fetcher = vi.fn(async () => pngBytes());
    expect(await resolveOnlinePicture("", { fetch: fetcher })).toEqual({
      ok: false,
      reason: "empty-url",
    });
    expect(await resolveOnlinePicture("   ", { fetch: fetcher })).toEqual({
      ok: false,
      reason: "empty-url",
    });
    expect(await resolveOnlinePicture("file:///tmp/a.png", { fetch: fetcher })).toEqual({
      ok: false,
      reason: "unsupported-url",
    });
    expect(await resolveOnlinePicture("javascript:alert(1)", { fetch: fetcher })).toEqual({
      ok: false,
      reason: "unsupported-url",
    });
    expect(await resolveOnlinePicture("not a url", { fetch: fetcher })).toEqual({
      ok: false,
      reason: "unsupported-url",
    });
    expect(fetcher).not.toHaveBeenCalled();
  });

  it("refuses downloads over the byte cap", async () => {
    const huge = new Uint8Array(64);
    huge.set(pngBytes().subarray(0, 8));
    const result = await resolveOnlinePicture("https://example.com/huge.png", {
      fetch: async () => huge,
      maxBytes: 32,
    });
    expect(result).toEqual({ ok: false, reason: "too-large" });
  });

  it("refuses non-image payloads (an HTML error page behind .png)", async () => {
    const html = new TextEncoder().encode("<html><body>404</body></html>");
    expect(
      await resolveOnlinePicture("https://example.com/missing.png", { fetch: async () => html }),
    ).toEqual({ ok: false, reason: "unsupported-type" });
    expect(
      await resolveOnlinePicture("https://example.com/empty.png", {
        fetch: async () => new Uint8Array(),
      }),
    ).toEqual({ ok: false, reason: "fetch-failed" });
  });

  it("exposes the documented 10 MiB cap", () => {
    expect(ONLINE_PICTURE_MAX_BYTES).toBe(10 * 1024 * 1024);
  });
});
