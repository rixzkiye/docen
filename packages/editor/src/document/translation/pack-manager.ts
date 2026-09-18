/**
 * Translation Pack Manager: manages preinstalled and on-demand translation packs
 * for language pairs (built-in en<->id, on-demand fr, de, es, zh).
 */

import {
  EN_TO_ID_PHRASES,
  EN_TO_ID_WORDS,
  ID_TO_EN_PHRASES,
  ID_TO_EN_WORDS,
} from "./id-en-dictionary";
import type { LanguageInfo, TranslationLanguagePack, TranslationPackInfo } from "./types";

const ALL_LANGUAGES: Record<string, LanguageInfo> = {
  en: { code: "en", name: "English", nativeName: "English" },
  id: { code: "id", name: "Indonesian", nativeName: "Bahasa Indonesia" },
  fr: { code: "fr", name: "French", nativeName: "Français" },
  de: { code: "de", name: "German", nativeName: "Deutsch" },
  es: { code: "es", name: "Spanish", nativeName: "Español" },
  zh: { code: "zh", name: "Chinese (Simplified)", nativeName: "中文 (简体)" },
};

/** Built-in English -> Indonesian pack */
const BUILTIN_EN_ID: TranslationLanguagePack = {
  id: "en-id",
  name: "English → Indonesian",
  from: "en",
  to: "id",
  dictionary: EN_TO_ID_WORDS,
  phrases: EN_TO_ID_PHRASES,
};

/** Built-in Indonesian -> English pack */
const BUILTIN_ID_EN: TranslationLanguagePack = {
  id: "id-en",
  name: "Indonesian → English",
  from: "id",
  to: "en",
  dictionary: ID_TO_EN_WORDS,
  phrases: ID_TO_EN_PHRASES,
};

/** Office and document vocabulary for downloadable packs */
const FR_VOCAB: Record<string, string> = {
  document: "document",
  documents: "documents",
  file: "fichier",
  files: "fichiers",
  table: "tableau",
  tables: "tableaux",
  row: "ligne",
  rows: "lignes",
  column: "colonne",
  columns: "colonnes",
  page: "page",
  pages: "pages",
  paragraph: "paragraphe",
  paragraphs: "paragraphes",
  title: "titre",
  text: "texte",
  save: "enregistrer",
  open: "ouvrir",
  close: "fermer",
  print: "imprimer",
  insert: "insérer",
  delete: "supprimer",
  edit: "modifier",
  copy: "copier",
  paste: "coller",
  cut: "couper",
  undo: "annuler",
  redo: "rétablir",
  search: "rechercher",
  find: "trouver",
  replace: "remplacer",
  select: "sélectionner",
  yes: "oui",
  no: "non",
  cancel: "annuler",
  apply: "appliquer",
  help: "aide",
  options: "options",
  settings: "paramètres",
  tools: "outils",
  good: "bon",
  bad: "mauvais",
  great: "excellent",
  important: "important",
  new: "nouveau",
  old: "ancien",
  status: "statut",
  summary: "résumé",
  report: "rapport",
};

const DE_VOCAB: Record<string, string> = {
  document: "Dokument",
  documents: "Dokumente",
  file: "Datei",
  files: "Dateien",
  table: "Tabelle",
  tables: "Tabellen",
  row: "Zeile",
  rows: "Zeilen",
  column: "Spalte",
  columns: "Spalten",
  page: "Seite",
  pages: "Seiten",
  paragraph: "Absatz",
  paragraphs: "Absätze",
  title: "Titel",
  text: "Text",
  save: "speichern",
  open: "öffnen",
  close: "schließen",
  print: "drucken",
  insert: "einfügen",
  delete: "löschen",
  edit: "bearbeiten",
  copy: "kopieren",
  paste: "einfügen",
  cut: "ausschneiden",
  undo: "rückgängig",
  redo: "wiederholen",
  search: "suchen",
  find: "finden",
  replace: "ersetzen",
  select: "auswählen",
  yes: "ja",
  no: "nein",
  cancel: "abbrechen",
  apply: "anwenden",
  help: "Hilfe",
  options: "Optionen",
  settings: "Einstellungen",
  tools: "Werkzeuge",
  good: "gut",
  bad: "schlecht",
  great: "großartig",
  important: "wichtig",
  new: "neu",
  old: "alt",
  status: "Status",
  summary: "Zusammenfassung",
  report: "Bericht",
};

const ES_VOCAB: Record<string, string> = {
  document: "documento",
  documents: "documentos",
  file: "archivo",
  files: "archivos",
  table: "tabla",
  tables: "tablas",
  row: "fila",
  rows: "filas",
  column: "columna",
  columns: "columnas",
  page: "página",
  pages: "páginas",
  paragraph: "párrafo",
  paragraphs: "párrafos",
  title: "título",
  text: "texto",
  save: "guardar",
  open: "abrir",
  close: "cerrar",
  print: "imprimir",
  insert: "insertar",
  delete: "eliminar",
  edit: "editar",
  copy: "copiar",
  paste: "pegar",
  cut: "cortar",
  undo: "deshacer",
  redo: "rehacer",
  search: "buscar",
  find: "encontrar",
  replace: "reemplazar",
  select: "seleccionar",
  yes: "sí",
  no: "no",
  cancel: "cancelar",
  apply: "aplicar",
  help: "ayuda",
  options: "opciones",
  settings: "configuración",
  tools: "herramientas",
  good: "bueno",
  bad: "malo",
  great: "genial",
  important: "importante",
  new: "nuevo",
  old: "viejo",
  status: "estado",
  summary: "resumen",
  report: "informe",
};

const ZH_VOCAB: Record<string, string> = {
  document: "文档",
  documents: "文档",
  file: "文件",
  files: "文件",
  table: "表格",
  tables: "表格",
  row: "行",
  rows: "行",
  column: "列",
  columns: "列",
  page: "页面",
  pages: "页面",
  paragraph: "段落",
  paragraphs: "段落",
  title: "标题",
  text: "文本",
  save: "保存",
  open: "打开",
  close: "关闭",
  print: "打印",
  insert: "插入",
  delete: "删除",
  edit: "编辑",
  copy: "复制",
  paste: "粘贴",
  cut: "剪切",
  undo: "撤销",
  redo: "重做",
  search: "搜索",
  find: "查找",
  replace: "替换",
  select: "选择",
  yes: "是",
  no: "否",
  cancel: "取消",
  apply: "应用",
  help: "帮助",
  options: "选项",
  settings: "设置",
  tools: "工具",
  good: "好",
  bad: "坏",
  great: "棒",
  important: "重要",
  new: "新",
  old: "旧",
  status: "状态",
  summary: "摘要",
  report: "报告",
};

export class TranslationPackManager {
  private installedPacks = new Map<string, TranslationLanguagePack>();

  constructor() {
    this.installedPacks.set("en-id", BUILTIN_EN_ID);
    this.installedPacks.set("id-en", BUILTIN_ID_EN);
  }

  /** Normalizes a language key e.g. "en-US" -> "en" */
  normalizeLang(lang?: string): string {
    if (!lang) return "en";
    const clean = lang.toLowerCase().trim();
    return clean.split("-")[0]!;
  }

  isInstalled(packId: string): boolean {
    const clean = packId.toLowerCase().trim();
    if (this.installedPacks.has(clean)) return true;
    // Check if both directions for language are installed or single lang key
    for (const pack of this.installedPacks.values()) {
      if (pack.id === clean || pack.to === clean || pack.from === clean) {
        return true;
      }
    }
    return false;
  }

  getPack(from: string, to: string): TranslationLanguagePack | null {
    const fromClean = this.normalizeLang(from);
    const toClean = this.normalizeLang(to);
    const pairId = `${fromClean}-${toClean}`;
    if (this.installedPacks.has(pairId)) {
      return this.installedPacks.get(pairId)!;
    }
    // Also check target language pack if available
    if (this.installedPacks.has(toClean)) {
      const pack = this.installedPacks.get(toClean)!;
      if (pack.from === fromClean && pack.to === toClean) return pack;
    }
    return null;
  }

  installPack(pack: TranslationLanguagePack): void {
    const key = pack.id.toLowerCase().trim();
    this.installedPacks.set(key, pack);
  }

  uninstallPack(packId: string): boolean {
    const clean = packId.toLowerCase().trim();
    // Do not allow uninstalling built-in packs
    if (clean === "en-id" || clean === "id-en" || clean === "en" || clean === "id") {
      return false;
    }
    let removed = false;
    if (this.installedPacks.has(clean)) {
      this.installedPacks.delete(clean);
      removed = true;
    }
    // Also remove any pack whose id, to, or from matches the language
    for (const [key, pack] of Array.from(this.installedPacks.entries())) {
      if (
        pack.id === clean ||
        pack.to === clean ||
        pack.from === clean ||
        pack.id === `en-${clean}` ||
        pack.id === `${clean}-en`
      ) {
        this.installedPacks.delete(key);
        removed = true;
      }
    }
    return removed;
  }

  listInstalledPacks(): Array<{ id: string; name: string }> {
    return Array.from(this.installedPacks.values()).map((p) => ({
      id: p.id,
      name: p.name,
    }));
  }

  listAvailablePacks(): TranslationPackInfo[] {
    const catalog: TranslationPackInfo[] = [
      {
        id: "en-id",
        name: "English ↔ Indonesian",
        installed: true,
        size: "180 KB",
        from: "en",
        to: "id",
      },
      {
        id: "fr",
        name: "Français (French)",
        installed: this.isInstalled("fr"),
        size: "120 KB",
        from: "en",
        to: "fr",
      },
      {
        id: "de",
        name: "Deutsch (German)",
        installed: this.isInstalled("de"),
        size: "140 KB",
        from: "en",
        to: "de",
      },
      {
        id: "es",
        name: "Español (Spanish)",
        installed: this.isInstalled("es"),
        size: "115 KB",
        from: "en",
        to: "es",
      },
      {
        id: "zh",
        name: "中文 (Chinese - Simplified)",
        installed: this.isInstalled("zh"),
        size: "130 KB",
        from: "en",
        to: "zh",
      },
    ];
    return catalog;
  }

  /**
   * Simulates downloading and installing an on-demand translation pack.
   * Extensible on-demand pack format for fr, de, es, zh.
   */
  async downloadPack(packId: string): Promise<TranslationLanguagePack> {
    const clean = this.normalizeLang(packId);
    let vocab: Record<string, string> = {};
    let langName = clean.toUpperCase();

    if (clean === "fr") {
      vocab = FR_VOCAB;
      langName = "Français (French)";
    } else if (clean === "de") {
      vocab = DE_VOCAB;
      langName = "Deutsch (German)";
    } else if (clean === "es") {
      vocab = ES_VOCAB;
      langName = "Español (Spanish)";
    } else if (clean === "zh") {
      vocab = ZH_VOCAB;
      langName = "中文 (Chinese - Simplified)";
    }

    const reverseVocab: Record<string, string> = {};
    for (const [k, v] of Object.entries(vocab)) {
      reverseVocab[v.toLowerCase()] = k;
    }

    const pack: TranslationLanguagePack = {
      id: `en-${clean}`,
      name: `English → ${langName}`,
      from: "en",
      to: clean,
      dictionary: vocab,
    };

    const reversePack: TranslationLanguagePack = {
      id: `${clean}-en`,
      name: `${langName} → English`,
      from: clean,
      to: "en",
      dictionary: reverseVocab,
    };

    // Install both forward and reverse
    this.installedPacks.set(pack.id, pack);
    this.installedPacks.set(reversePack.id, reversePack);
    // Also index under the language code directly
    this.installedPacks.set(clean, pack);

    return pack;
  }

  /** Returns list of currently supported languages based on installed packs */
  getSupportedLanguages(): LanguageInfo[] {
    const codes = new Set<string>(["en", "id"]);
    for (const pack of this.installedPacks.values()) {
      if (pack.from) codes.add(pack.from);
      if (pack.to) codes.add(pack.to);
    }
    const result: LanguageInfo[] = [];
    for (const code of codes) {
      if (ALL_LANGUAGES[code]) {
        result.push(ALL_LANGUAGES[code]);
      } else {
        result.push({ code, name: code.toUpperCase(), nativeName: code.toUpperCase() });
      }
    }
    return result;
  }
}

export const translationPackManager = new TranslationPackManager();
