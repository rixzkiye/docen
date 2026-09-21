// @vitest-environment node
import { deflateSync } from "node:zlib";

import { describe, expect, it } from "vitest";

import { NodeImageCache, decodePngToRgba } from "./node-image";

/**
 * Minimal PNG encoder for fixtures: filter-0 rows, one IDAT, CRC left zero
 * (the decoder under test reads lengths/types, not CRCs).
 */
function makePng(
  width: number,
  height: number,
  colorType: 2 | 6,
  pixels: readonly (readonly number[])[],
): Uint8Array {
  const bpp = colorType === 2 ? 3 : 4;
  const raw = new Uint8Array(height * (1 + width * bpp));
  let o = 0;
  for (let y = 0; y < height; y++) {
    raw[o++] = 0; // filter: none
    for (let x = 0; x < width; x++) {
      const px = pixels[y * width + x]!;
      for (let c = 0; c < bpp; c++) raw[o++] = px[c]!;
    }
  }
  const chunk = (type: string, data: Uint8Array): Uint8Array => {
    const out = new Uint8Array(12 + data.length);
    const view = new DataView(out.buffer);
    view.setUint32(0, data.length);
    for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
    out.set(data, 8);
    return out;
  };
  const ihdr = new Uint8Array(13);
  const ihdrView = new DataView(ihdr.buffer);
  ihdrView.setUint32(0, width);
  ihdrView.setUint32(4, height);
  ihdr[8] = 8; // bit depth
  ihdr[9] = colorType;
  return new Uint8Array([
    0x89,
    0x50,
    0x4e,
    0x47,
    0x0d,
    0x0a,
    0x1a,
    0x0a,
    ...chunk("IHDR", ihdr),
    ...chunk("IDAT", new Uint8Array(deflateSync(raw))),
    ...chunk("IEND", new Uint8Array(0)),
  ]);
}

describe("NodeImageCache channel contract", () => {
  it("decodes an RGB PNG to the exact pixels", () => {
    const png = makePng(2, 2, 2, [
      [255, 0, 0],
      [0, 255, 0],
      [0, 0, 255],
      [255, 255, 255],
    ]);
    const decoded = decodePngToRgba(png);
    expect(decoded).toBeDefined();
    expect(decoded!.width).toBe(2);
    expect(decoded!.height).toBe(2);
    expect(decoded!.hasAlpha).toBe(false);
    expect([...decoded!.rgba]).toEqual([
      255, 0, 0, 255, 0, 255, 0, 255, 0, 0, 255, 255, 255, 255, 255, 255,
    ]);
  });

  it("returns 3-channel RGB for opaque PNGs (the DeviceRGB embedding contract)", async () => {
    const cache = new NodeImageCache();
    const png = makePng(2, 2, 2, [
      [255, 0, 0],
      [0, 255, 0],
      [0, 0, 255],
      [255, 255, 255],
    ]);
    const image = await cache.get(`data:image/png;base64,${Buffer.from(png).toString("base64")}`);
    expect(image).toBeDefined();
    expect(image!.hasAlpha).toBe(false);
    // 3 channels per pixel: the exporter embeds an opaque buffer as raw
    // DeviceRGB, so a 4-channel buffer would shear every row.
    expect(image!.rgba).toHaveLength(2 * 2 * 3);
    expect(image!.rgba!.slice(0, 3)).toEqual(new Uint8Array([255, 0, 0]));
    expect(image!.rgba!.slice(9, 12)).toEqual(new Uint8Array([255, 255, 255]));
  });

  it("keeps 4-channel RGBA (with hasAlpha) so the exporter can build an /SMask", async () => {
    const cache = new NodeImageCache();
    const png = makePng(1, 1, 6, [[10, 20, 30, 128]]);
    const image = await cache.get(`data:image/png;base64,${Buffer.from(png).toString("base64")}`);
    expect(image).toBeDefined();
    expect(image!.hasAlpha).toBe(true);
    expect(image!.rgba).toHaveLength(4);
    expect(image!.rgba).toEqual(new Uint8Array([10, 20, 30, 128]));
  });
});
