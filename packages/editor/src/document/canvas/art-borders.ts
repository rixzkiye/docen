/**
 * Art border SVG presets and CSS border-image generator.
 *
 * Custom, bespoke vector graphics (non-MS assets) for Word w:art parity.
 * Clean, lightweight 9-slice SVG patterns that render crisply at any zoom.
 */

export interface ArtBorderPreset {
  id: string;
  name: string;
  nameZh: string;
}

export const ART_BORDER_PRESETS: readonly ArtBorderPreset[] = [
  { id: "stars", name: "Stars", nameZh: "星星" },
  { id: "hearts", name: "Hearts", nameZh: "心形" },
  { id: "apples", name: "Apples", nameZh: "苹果" },
  { id: "diamonds", name: "Diamonds", nameZh: "菱形" },
  { id: "decoArch", name: "Art Deco", nameZh: "装饰艺术" },
  { id: "classic", name: "Classic Flourish", nameZh: "古典花纹" },
  { id: "zigzag", name: "Zigzag Chevron", nameZh: "锯齿折线" },
  { id: "vines", name: "Floral Vines", nameZh: "藤蔓叶片" },
  { id: "dots", name: "Pearls & Dots", nameZh: "珍珠圆点" },
  { id: "doubleWave", name: "Double Wave", nameZh: "双重波浪" },
] as const;

type SvgTileBuilder = (c: string) => string;

const TILE_BUILDERS: Record<string, SvgTileBuilder> = {
  stars: (c) => `
    <g fill="${c}">
      <polygon points="10,2 12.5,7.5 18,8 14,12 15,18 10,15 5,18 6,12 2,8 7.5,7.5"/>
      <polygon points="30,2 32.5,7.5 38,8 34,12 35,18 30,15 25,18 26,12 22,8 27.5,7.5"/>
      <polygon points="50,2 52.5,7.5 58,8 54,12 55,18 50,15 45,18 46,12 42,8 47.5,7.5"/>
      <polygon points="10,22 12.5,27.5 18,28 14,32 15,38 10,35 5,38 6,32 2,28 7.5,27.5"/>
      <polygon points="50,22 52.5,27.5 58,28 54,32 55,38 50,35 45,38 46,32 42,28 47.5,27.5"/>
      <polygon points="10,42 12.5,47.5 18,48 14,52 15,58 10,55 5,58 6,52 2,48 7.5,47.5"/>
      <polygon points="30,42 32.5,47.5 38,48 34,52 35,58 30,55 25,58 26,52 22,48 27.5,47.5"/>
      <polygon points="50,42 52.5,47.5 58,48 54,52 55,58 50,55 45,58 46,52 42,48 47.5,47.5"/>
    </g>`,

  hearts: (c) => `
    <g fill="${c}">
      <path d="M10,6 C8,3 4,4 4,8 C4,12 8,15 10,17 C12,15 16,12 16,8 C16,4 12,3 10,6 Z"/>
      <path d="M30,6 C28,3 24,4 24,8 C24,12 28,15 30,17 C32,15 36,12 36,8 C36,4 32,3 30,6 Z"/>
      <path d="M50,6 C48,3 44,4 44,8 C44,12 48,15 50,17 C52,15 56,12 56,8 C56,4 52,3 50,6 Z"/>
      <path d="M10,26 C8,23 4,24 4,28 C4,32 8,35 10,37 C12,35 16,32 16,28 C16,24 12,23 10,26 Z"/>
      <path d="M50,26 C48,23 44,24 44,28 C44,32 48,35 50,37 C52,35 56,32 56,28 C56,24 52,23 50,26 Z"/>
      <path d="M10,46 C8,43 4,44 4,48 C4,52 8,55 10,57 C12,55 16,52 16,48 C16,44 12,43 10,46 Z"/>
      <path d="M30,46 C28,43 24,44 24,48 C24,52 28,55 30,57 C32,55 36,52 36,48 C36,44 32,43 30,46 Z"/>
      <path d="M50,46 C48,43 44,44 44,48 C44,52 48,55 50,57 C52,55 56,52 56,48 C56,44 52,43 50,46 Z"/>
    </g>`,

  apples: (c) => `
    <g fill="${c}" stroke="${c}" stroke-width="1">
      <path d="M10,4 C11,2 13,2 13,2 M10,5 C7,4 4,6 4,11 C4,16 8,18 10,18 C12,18 16,16 16,11 C16,6 13,4 10,5 Z"/>
      <path d="M30,4 C31,2 33,2 33,2 M30,5 C27,4 24,6 24,11 C24,16 28,18 30,18 C32,18 36,16 36,11 C36,6 33,4 30,5 Z"/>
      <path d="M50,4 C51,2 53,2 53,2 M50,5 C47,4 44,6 44,11 C44,16 48,18 50,18 C52,18 56,16 56,11 C56,6 53,4 50,5 Z"/>
      <path d="M10,24 C11,22 13,22 13,22 M10,25 C7,24 4,26 4,31 C4,36 8,38 10,38 C12,38 16,36 16,31 C16,26 13,24 10,25 Z"/>
      <path d="M50,24 C51,22 53,22 53,22 M50,25 C47,24 44,26 44,31 C44,36 48,38 50,38 C52,38 56,36 56,31 C56,26 53,24 50,25 Z"/>
      <path d="M10,44 C11,42 13,42 13,42 M10,45 C7,44 4,46 4,51 C4,56 8,58 10,58 C12,58 16,56 16,51 C16,46 13,44 10,45 Z"/>
      <path d="M30,44 C31,42 33,42 33,42 M30,45 C27,44 24,46 24,51 C24,56 28,58 30,58 C32,58 36,56 36,51 C36,46 33,44 30,45 Z"/>
      <path d="M50,44 C51,42 53,42 53,42 M50,45 C47,44 44,46 44,51 C44,56 48,58 50,58 C52,58 56,56 56,51 C56,46 53,44 50,45 Z"/>
    </g>`,

  diamonds: (c) => `
    <g fill="${c}">
      <polygon points="10,2 18,10 10,18 2,10"/>
      <polygon points="30,2 38,10 30,18 22,10"/>
      <polygon points="50,2 58,10 50,18 42,10"/>
      <polygon points="10,22 18,30 10,38 2,30"/>
      <polygon points="50,22 58,30 50,38 42,30"/>
      <polygon points="10,42 18,50 10,58 2,50"/>
      <polygon points="30,42 38,50 30,58 22,50"/>
      <polygon points="50,42 58,50 50,58 42,50"/>
    </g>`,

  decoArch: (c) => `
    <g fill="none" stroke="${c}" stroke-width="1.5">
      <path d="M2,18 L2,10 L10,2 L18,10 L18,18 M5,18 L5,11 L10,6 L15,11 L15,18 M8,18 L8,13 L10,11 L12,13 L12,18"/>
      <path d="M22,18 L22,10 L30,2 L38,10 L38,18 M25,18 L25,11 L30,6 L35,11 L35,18 M28,18 L28,13 L30,11 L32,13 L32,18"/>
      <path d="M42,18 L42,10 L50,2 L58,10 L58,18 M45,18 L45,11 L50,6 L55,11 L55,18 M48,18 L48,13 L50,11 L52,13 L52,18"/>
      <path d="M2,38 L2,30 L10,22 L18,30 L18,38 M5,38 L5,31 L10,26 L15,31 L15,38"/>
      <path d="M42,38 L42,30 L50,22 L58,30 L58,38 M45,38 L45,31 L50,26 L55,31 L55,38"/>
      <path d="M2,58 L2,50 L10,42 L18,50 L18,58 M5,58 L5,51 L10,46 L15,51 L15,58"/>
      <path d="M22,58 L22,50 L30,42 L38,50 L38,58 M25,58 L25,51 L30,46 L35,51 L35,58"/>
      <path d="M42,58 L42,50 L50,42 L58,50 L58,58 M45,58 L45,51 L50,46 L55,51 L55,58"/>
    </g>`,

  classic: (c) => `
    <g fill="none" stroke="${c}" stroke-width="1.5">
      <circle cx="10" cy="10" r="7"/>
      <circle cx="30" cy="10" r="7"/>
      <circle cx="50" cy="10" r="7"/>
      <circle cx="10" cy="30" r="7"/>
      <circle cx="50" cy="30" r="7"/>
      <circle cx="10" cy="50" r="7"/>
      <circle cx="30" cy="50" r="7"/>
      <circle cx="50" cy="50" r="7"/>
      <circle cx="10" cy="10" r="3" fill="${c}"/>
      <circle cx="30" cy="10" r="3" fill="${c}"/>
      <circle cx="50" cy="10" r="3" fill="${c}"/>
      <circle cx="10" cy="30" r="3" fill="${c}"/>
      <circle cx="50" cy="30" r="3" fill="${c}"/>
      <circle cx="10" cy="50" r="3" fill="${c}"/>
      <circle cx="30" cy="50" r="3" fill="${c}"/>
      <circle cx="50" cy="50" r="3" fill="${c}"/>
    </g>`,

  zigzag: (c) => `
    <g fill="none" stroke="${c}" stroke-width="2">
      <path d="M0,5 L10,15 L20,5 L30,15 L40,5 L50,15 L60,5"/>
      <path d="M0,20 L10,30 L20,20 L30,30 L40,20 L50,30 L60,20"/>
      <path d="M0,45 L10,55 L20,45 L30,55 L40,45 L50,55 L60,45"/>
    </g>`,

  vines: (c) => `
    <g fill="${c}" stroke="${c}" stroke-width="1">
      <path d="M0,10 Q10,5 20,10 T40,10 T60,10" fill="none" stroke-width="2"/>
      <path d="M10,8 Q8,3 12,2 Q14,5 10,8 Z"/>
      <path d="M30,8 Q28,3 32,2 Q34,5 30,8 Z"/>
      <path d="M50,8 Q48,3 52,2 Q54,5 50,8 Z"/>
      <path d="M0,50 Q10,45 20,50 T40,50 T60,50" fill="none" stroke-width="2"/>
      <path d="M10,48 Q8,43 12,42 Q14,45 10,48 Z"/>
      <path d="M30,48 Q28,43 32,42 Q34,45 30,48 Z"/>
      <path d="M50,48 Q48,43 52,42 Q54,45 50,48 Z"/>
    </g>`,

  dots: (c) => `
    <g fill="${c}">
      <circle cx="10" cy="10" r="5"/>
      <circle cx="30" cy="10" r="5"/>
      <circle cx="50" cy="10" r="5"/>
      <circle cx="10" cy="30" r="5"/>
      <circle cx="50" cy="30" r="5"/>
      <circle cx="10" cy="50" r="5"/>
      <circle cx="30" cy="50" r="5"/>
      <circle cx="50" cy="50" r="5"/>
    </g>`,

  doubleWave: (c) => `
    <g fill="none" stroke="${c}" stroke-width="2">
      <path d="M0,6 Q10,2 20,6 T40,6 T60,6"/>
      <path d="M0,14 Q10,10 20,14 T40,14 T60,14"/>
      <path d="M0,46 Q10,42 20,46 T40,46 T60,46"/>
      <path d="M0,54 Q10,50 20,54 T40,54 T60,54"/>
    </g>`,
};

/** Normalize any incoming art name (including Word's 160+ OOXML tokens) into a preset. */
export function resolveArtPreset(name: string): string {
  const lower = name.toLowerCase();
  if (lower.includes("star")) return "stars";
  if (lower.includes("heart")) return "hearts";
  if (lower.includes("apple") || lower.includes("fruit")) return "apples";
  if (lower.includes("diamond")) return "diamonds";
  if (lower.includes("deco") || lower.includes("arch")) return "decoArch";
  if (lower.includes("zig") || lower.includes("chev")) return "zigzag";
  if (lower.includes("vine") || lower.includes("leaf") || lower.includes("flora")) return "vines";
  if (lower.includes("dot") || lower.includes("bead") || lower.includes("pearl")) return "dots";
  if (lower.includes("wave") || lower.includes("ribbon")) return "doubleWave";
  if (TILE_BUILDERS[name]) return name;
  return "classic";
}

export function isArtPreset(name: string): boolean {
  const resolved = resolveArtPreset(name);
  return Boolean(TILE_BUILDERS[resolved]);
}

/** Build an SVG data URI string for CSS border-image-source. */
export function getArtBorderSvgDataUri(artName: string, colorHex?: string): string {
  const preset = resolveArtPreset(artName);
  const builder = TILE_BUILDERS[preset] ?? TILE_BUILDERS.classic!;
  const color =
    colorHex && colorHex !== "auto"
      ? colorHex.startsWith("#")
        ? colorHex
        : `#${colorHex}`
      : "#2F5597";
  const inner = builder(color);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="60" height="60" viewBox="0 0 60 60">${inner}</svg>`;
  return `data:image/svg+xml;utf8,${encodeURIComponent(svg)}`;
}
