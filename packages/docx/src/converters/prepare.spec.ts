import { createServer, type Server } from "node:http";
import type { AddressInfo } from "node:net";

import type { JSONContent } from "@tiptap/core";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { fetchImageHandler, prepareDocument, prepareImageSizes, prepareImages } from "../index";

const PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const PNG_BYTES = Buffer.from(PNG_BASE64, "base64");
const PNG_DATA_URL = `data:image/png;base64,${PNG_BASE64}`;

function imageDoc(src: string, attrs: Record<string, unknown> = {}): JSONContent {
  return {
    type: "doc",
    content: [{ type: "paragraph", content: [{ type: "image", attrs: { src, ...attrs } }] }],
  };
}

const src = (json: JSONContent): string =>
  (json.content?.[0]?.content?.[0]?.attrs?.src as string | undefined) ?? "";

describe("prepareDocument safety contract", () => {
  let server: Server;
  let origin = "";
  let hits = 0;

  beforeAll(async () => {
    server = createServer((request, response) => {
      hits += 1;
      const url = new URL(request.url ?? "/", "http://127.0.0.1");
      switch (url.pathname) {
        case "/ok":
          response.writeHead(200, { "content-type": "image/png" });
          response.end(PNG_BYTES);
          return;
        case "/redirect":
          response.writeHead(302, { location: "/ok" });
          response.end();
          return;
        case "/loop":
          response.writeHead(302, { location: "/loop" });
          response.end();
          return;
        case "/external":
          response.writeHead(302, { location: "http://evil.example/ok" });
          response.end();
          return;
        case "/big":
          response.writeHead(200, { "content-type": "image/png" });
          response.end(Buffer.alloc(4096, 7));
          return;
        case "/slow":
          setTimeout(() => {
            response.writeHead(200, { "content-type": "image/png" });
            response.end(PNG_BYTES);
          }, 250);
          return;
        default:
          response.writeHead(404);
          response.end();
      }
    });
    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
    origin = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  });

  afterAll(async () => {
    await new Promise<void>((resolve) => server.close(() => resolve()));
  });

  it("does not touch the network by default and does not mutate its input", async () => {
    const json = imageDoc(`${origin}/ok`, { alt: "probe" });
    const snapshot = structuredClone(json);
    const before = hits;

    const prepared = await prepareDocument(json);

    expect(hits).toBe(before);
    expect(json).toEqual(snapshot);
    expect(prepared).not.toBe(json);
    expect(src(prepared)).toBe(`${origin}/ok`);
  });

  it("prepareImages() without an allowlist is a no-op", async () => {
    const json = imageDoc(`${origin}/ok`);
    const before = hits;
    await prepareDocument(json, [prepareImages()]);
    await prepareDocument(json, [prepareImages({ allow: [] })]);
    expect(hits).toBe(before);
    expect(src(json)).toBe(`${origin}/ok`);
  });

  it("fetches and embeds an allowlisted host", async () => {
    const json = imageDoc(`${origin}/ok`);
    const prepared = await prepareDocument(json, [
      prepareImages({ allow: ["127.0.0.1"], maxBytes: 1024 }),
      prepareImageSizes(),
    ]);
    expect(src(prepared)).toBe(PNG_DATA_URL);
    expect(src(json)).toBe(`${origin}/ok`);
    expect(prepared.content?.[0]?.content?.[0]?.attrs?.width).toBe(1);
  });

  it("leaves images from a non-allowlisted host untouched without fetching", async () => {
    const json = imageDoc(`${origin}/ok`);
    const before = hits;
    const prepared = await prepareDocument(json, [prepareImages({ allow: ["cdn.example.com"] })]);
    expect(src(prepared)).toBe(`${origin}/ok`);
    expect(hits).toBe(before);
  });

  it("enforces the response size cap", async () => {
    const prepared = await prepareDocument(imageDoc(`${origin}/big`), [
      prepareImages({ allow: ["127.0.0.1"], maxBytes: 64 }),
    ]);
    expect(src(prepared)).toBe(`${origin}/big`);
  });

  it("enforces the redirect cap", async () => {
    const prepared = await prepareDocument(imageDoc(`${origin}/loop`), [
      prepareImages({ allow: ["127.0.0.1"], maxRedirects: 1 }),
    ]);
    expect(src(prepared)).toBe(`${origin}/loop`);
  });

  it("re-validates every redirect hop against the allowlist", async () => {
    const prepared = await prepareDocument(imageDoc(`${origin}/external`), [
      prepareImages({ allow: ["127.0.0.1"] }),
    ]);
    expect(src(prepared)).toBe(`${origin}/external`);
  });

  it("enforces the request timeout", async () => {
    const prepared = await prepareDocument(imageDoc(`${origin}/slow`), [
      prepareImages({ allow: ["127.0.0.1"], timeoutMs: 50 }),
    ]);
    expect(src(prepared)).toBe(`${origin}/slow`);
  });

  it("never fetches non-http(s) schemes", async () => {
    const before = hits;
    const prepared = await prepareDocument(imageDoc("file:///etc/passwd"), [
      prepareImages({ allow: ["127.0.0.1"] }),
    ]);
    expect(hits).toBe(before);
    expect(src(prepared)).toBe("file:///etc/passwd");
  });

  it("runs a custom transport only inside the policy caps", async () => {
    const calls: string[] = [];
    const results = new Map<string, Uint8Array>([
      ["https://cdn.example.com/ok.png", PNG_BYTES],
      ["https://cdn.example.com/big.png", new Uint8Array(4096)],
    ]);
    const policy = {
      allow: ["cdn.example.com"],
      maxBytes: 1024,
      fetch: async (url: string) => {
        calls.push(url);
        return results.get(url)!;
      },
    };

    const embedded = await prepareDocument(imageDoc("https://cdn.example.com/ok.png"), [
      prepareImages(policy),
    ]);
    expect(src(embedded)).toBe(PNG_DATA_URL);

    const capped = await prepareDocument(imageDoc("https://cdn.example.com/big.png"), [
      prepareImages(policy),
    ]);
    expect(src(capped)).toBe("https://cdn.example.com/big.png");

    expect(calls).toEqual(["https://cdn.example.com/ok.png", "https://cdn.example.com/big.png"]);
  });

  it("hands custom steps an isolated copy", async () => {
    const json = imageDoc(PNG_DATA_URL);
    const snapshot = structuredClone(json);
    const prepared = await prepareDocument(json, [
      async (copy) => {
        copy.content![0]!.content![0]!.attrs!.alt = "touched";
      },
    ]);
    expect(json).toEqual(snapshot);
    expect(prepared.content?.[0]?.content?.[0]?.attrs?.alt).toBe("touched");
  });
});

describe("fetchImageHandler (explicit UI fetch)", () => {
  it("decodes a data:image URL locally under the byte cap", async () => {
    const bytes = await fetchImageHandler(PNG_DATA_URL);
    expect(Buffer.from(bytes).toString("base64")).toBe(PNG_BASE64);
    await expect(fetchImageHandler(PNG_DATA_URL, { maxBytes: 4 })).rejects.toThrow(/cap/);
  });

  it("rejects non-http(s) schemes and credentialed URLs", async () => {
    await expect(fetchImageHandler("file:///etc/passwd")).rejects.toThrow(/scheme/);
    await expect(fetchImageHandler("https://user:pass@example.com/a.png")).rejects.toThrow(
      /credentials/,
    );
  });

  it("applies the size cap to a custom transport and enforces an optional allowlist", async () => {
    const transport = async (): Promise<Uint8Array> => new Uint8Array(4096);
    await expect(
      fetchImageHandler("https://cdn.example.com/big.png", { maxBytes: 1024, fetch: transport }),
    ).rejects.toThrow(/cap/);
    await expect(
      fetchImageHandler("https://other.example.com/a.png", {
        allow: ["cdn.example.com"],
        fetch: transport,
      }),
    ).rejects.toThrow(/allowlisted/);
  });
});
