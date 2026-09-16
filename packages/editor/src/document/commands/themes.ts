export interface ThemeColorScheme {
  id: string;
  name: string;
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hlink: string;
  folHlink: string;
}

export interface ThemeFontScheme {
  id: string;
  name: string;
  majorFont: string; // Headings
  minorFont: string; // Body
}

export interface ThemeDefinition {
  id: string;
  name: string;
  colors: ThemeColorScheme;
  fonts: ThemeFontScheme;
}

export const COLOR_SCHEMES: Record<string, ThemeColorScheme> = {
  office: {
    id: "office",
    name: "Office",
    dk1: "000000",
    lt1: "ffffff",
    dk2: "1f497d",
    lt2: "eeece1",
    accent1: "4f81bd",
    accent2: "c0504d",
    accent3: "9bbb59",
    accent4: "8064a2",
    accent5: "4bacc6",
    accent6: "f79646",
    hlink: "0000ff",
    folHlink: "800080",
  },
  facet: {
    id: "facet",
    name: "Facet",
    dk1: "000000",
    lt1: "ffffff",
    dk2: "222b35",
    lt2: "e8ebee",
    accent1: "91a14f",
    accent2: "59958f",
    accent3: "8064a2",
    accent4: "c87c4a",
    accent5: "487e8f",
    accent6: "bc5652",
    hlink: "3e6b7d",
    folHlink: "765082",
  },
  integral: {
    id: "integral",
    name: "Integral",
    dk1: "000000",
    lt1: "ffffff",
    dk2: "2b2e34",
    lt2: "f3f3f4",
    accent1: "1b587c",
    accent2: "b84627",
    accent3: "cf7c1c",
    accent4: "497a48",
    accent5: "714674",
    accent6: "3d647c",
    hlink: "1b587c",
    folHlink: "714674",
  },
  ion: {
    id: "ion",
    name: "Ion",
    dk1: "000000",
    lt1: "ffffff",
    dk2: "293845",
    lt2: "f1f3f5",
    accent1: "cf3d1e",
    accent2: "f17829",
    accent3: "fec829",
    accent4: "5e9e37",
    accent5: "357a85",
    accent6: "67537a",
    hlink: "cf3d1e",
    folHlink: "67537a",
  },
  organic: {
    id: "organic",
    name: "Organic",
    dk1: "000000",
    lt1: "ffffff",
    dk2: "39352e",
    lt2: "f6f4f0",
    accent1: "727a3c",
    accent2: "a88e38",
    accent3: "bf6e2e",
    accent4: "a84a32",
    accent5: "687b87",
    accent6: "595959",
    hlink: "727a3c",
    folHlink: "a84a32",
  },
  retrospect: {
    id: "retrospect",
    name: "Retrospect",
    dk1: "000000",
    lt1: "ffffff",
    dk2: "302e2b",
    lt2: "f3f2ee",
    accent1: "d1583a",
    accent2: "dca03d",
    accent3: "9fa03e",
    accent4: "527a69",
    accent5: "615b7a",
    accent6: "975369",
    hlink: "d1583a",
    folHlink: "975369",
  },
  slice: {
    id: "slice",
    name: "Slice",
    dk1: "000000",
    lt1: "ffffff",
    dk2: "2a2a2a",
    lt2: "f3f3f3",
    accent1: "365f91",
    accent2: "b84627",
    accent3: "7e9c3c",
    accent4: "604a7b",
    accent5: "31859c",
    accent6: "e46c0a",
    hlink: "365f91",
    folHlink: "604a7b",
  },
  wisp: {
    id: "wisp",
    name: "Wisp",
    dk1: "000000",
    lt1: "ffffff",
    dk2: "2d3330",
    lt2: "f4f6f5",
    accent1: "487d76",
    accent2: "9f9a46",
    accent3: "bb6640",
    accent4: "a2515b",
    accent5: "587588",
    accent6: "766a85",
    hlink: "487d76",
    folHlink: "a2515b",
  },
};

export const FONT_SCHEMES: Record<string, ThemeFontScheme> = {
  office: {
    id: "office",
    name: "Office (Aptos / Aptos Display)",
    majorFont: "Aptos Display",
    minorFont: "Aptos",
  },
  calibri: {
    id: "calibri",
    name: "Calibri / Calibri Light",
    majorFont: "Calibri Light",
    minorFont: "Calibri",
  },
  centuryGothic: {
    id: "centuryGothic",
    name: "Century Gothic",
    majorFont: "Century Gothic",
    minorFont: "Century Gothic",
  },
  georgia: {
    id: "georgia",
    name: "Georgia / Garamond",
    majorFont: "Georgia",
    minorFont: "Garamond",
  },
  arial: {
    id: "arial",
    name: "Arial",
    majorFont: "Arial Black",
    minorFont: "Arial",
  },
  times: {
    id: "times",
    name: "Times New Roman",
    majorFont: "Times New Roman",
    minorFont: "Times New Roman",
  },
};

export const THEMES: Record<string, ThemeDefinition> = {
  office: {
    id: "office",
    name: "Office",
    colors: COLOR_SCHEMES.office,
    fonts: FONT_SCHEMES.office,
  },
  facet: {
    id: "facet",
    name: "Facet",
    colors: COLOR_SCHEMES.facet,
    fonts: FONT_SCHEMES.centuryGothic,
  },
  integral: {
    id: "integral",
    name: "Integral",
    colors: COLOR_SCHEMES.integral,
    fonts: FONT_SCHEMES.georgia,
  },
  ion: {
    id: "ion",
    name: "Ion",
    colors: COLOR_SCHEMES.ion,
    fonts: FONT_SCHEMES.centuryGothic,
  },
  organic: {
    id: "organic",
    name: "Organic",
    colors: COLOR_SCHEMES.organic,
    fonts: FONT_SCHEMES.georgia,
  },
  retrospect: {
    id: "retrospect",
    name: "Retrospect",
    colors: COLOR_SCHEMES.retrospect,
    fonts: FONT_SCHEMES.times,
  },
  slice: {
    id: "slice",
    name: "Slice",
    colors: COLOR_SCHEMES.slice,
    fonts: FONT_SCHEMES.arial,
  },
  wisp: {
    id: "wisp",
    name: "Wisp",
    colors: COLOR_SCHEMES.wisp,
    fonts: FONT_SCHEMES.calibri,
  },
};

export function getTheme(id: string): ThemeDefinition {
  return THEMES[id] ?? THEMES.office;
}

export function getColorScheme(id: string): ThemeColorScheme {
  return COLOR_SCHEMES[id] ?? COLOR_SCHEMES.office;
}

export function getFontScheme(id: string): ThemeFontScheme {
  return FONT_SCHEMES[id] ?? FONT_SCHEMES.office;
}
