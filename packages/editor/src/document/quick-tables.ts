import type { JSONContent } from "@docen/docx";
import { zipSync } from "@office-open/core";

import type { BuildingBlock } from "./building-blocks";

export type QuickTableId =
  | "calendar1"
  | "calendar2"
  | "matrix"
  | "tabularList"
  | "tabular-list"
  | "doubleTable"
  | "double-table"
  | "subheadings";

export interface QuickTablePreset {
  readonly id: string;
  readonly name: string;
  readonly nameZh: string;
  readonly description: string;
  readonly descriptionZh: string;
  build(): JSONContent;
}

const GRID_BORDER = { style: "single", px: 1, color: "D2D0CE" };

function cell(
  text: string,
  options?: {
    bold?: boolean;
    italic?: boolean;
    fill?: string;
    colspan?: number;
    rowspan?: number;
    align?: string;
    color?: string;
  },
): JSONContent {
  const marks: JSONContent["marks"] = [];
  if (options?.bold) marks.push({ type: "bold" });
  if (options?.italic) marks.push({ type: "italic" });
  if (options?.color) marks.push({ type: "textStyle", attrs: { color: options.color } });

  return {
    type: "tableCell",
    attrs: {
      ...(options?.colspan && options.colspan > 1 ? { colspan: options.colspan } : {}),
      ...(options?.rowspan && options.rowspan > 1 ? { rowspan: options.rowspan } : {}),
      ...(options?.fill ? { fill: options.fill } : {}),
      borders: {
        top: GRID_BORDER,
        bottom: GRID_BORDER,
        left: GRID_BORDER,
        right: GRID_BORDER,
      },
    },
    content: [
      {
        type: "paragraph",
        attrs: options?.align ? { textAlignment: options.align } : undefined,
        content: text
          ? [
              {
                type: "text",
                text,
                ...(marks.length > 0 ? { marks } : {}),
              },
            ]
          : [],
      },
    ],
  };
}

function row(cells: JSONContent[]): JSONContent {
  return {
    type: "tableRow",
    content: cells,
  };
}

function table(columnWidths: number[], rows: JSONContent[]): JSONContent {
  return {
    type: "table",
    attrs: {
      columnWidths,
      borders: {
        top: GRID_BORDER,
        bottom: GRID_BORDER,
        left: GRID_BORDER,
        right: GRID_BORDER,
        insideHorizontal: GRID_BORDER,
        insideVertical: GRID_BORDER,
      },
    },
    content: rows,
  };
}

/** Build Calendar 1: Monthly calendar grid with centered title and days. */
function buildCalendar1(): JSONContent {
  const widths = [60, 60, 60, 60, 60, 60, 60];
  return table(widths, [
    // Month Title
    row([cell("May", { bold: true, align: "center", fill: "D9E1F2", colspan: 7 })]),
    // Day headers
    row([
      cell("M", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("T", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("W", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("T", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("F", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("S", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("S", { bold: true, align: "center", fill: "F2F2F2" }),
    ]),
    // Dates
    row([
      cell("", { align: "center" }),
      cell("", { align: "center" }),
      cell("", { align: "center" }),
      cell("1", { align: "center" }),
      cell("2", { align: "center" }),
      cell("3", { align: "center" }),
      cell("4", { align: "center" }),
    ]),
    row([
      cell("5", { align: "center" }),
      cell("6", { align: "center" }),
      cell("7", { align: "center" }),
      cell("8", { align: "center" }),
      cell("9", { align: "center" }),
      cell("10", { align: "center" }),
      cell("11", { align: "center" }),
    ]),
    row([
      cell("12", { align: "center" }),
      cell("13", { align: "center" }),
      cell("14", { align: "center" }),
      cell("15", { align: "center" }),
      cell("16", { align: "center" }),
      cell("17", { align: "center" }),
      cell("18", { align: "center" }),
    ]),
    row([
      cell("19", { align: "center" }),
      cell("20", { align: "center" }),
      cell("21", { align: "center" }),
      cell("22", { align: "center" }),
      cell("23", { align: "center" }),
      cell("24", { align: "center" }),
      cell("25", { align: "center" }),
    ]),
    row([
      cell("26", { align: "center" }),
      cell("27", { align: "center" }),
      cell("28", { align: "center" }),
      cell("29", { align: "center" }),
      cell("30", { align: "center" }),
      cell("31", { align: "center" }),
      cell("", { align: "center" }),
    ]),
  ]);
}

/** Build Calendar 2: Left column with month title, right columns with days grid. */
function buildCalendar2(): JSONContent {
  const widths = [90, 50, 50, 50, 50, 50, 50, 50];
  return table(widths, [
    row([
      cell("MAY", { bold: true, align: "center", fill: "D9E1F2", rowspan: 6 }),
      cell("S", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("M", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("T", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("W", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("T", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("F", { bold: true, align: "center", fill: "F2F2F2" }),
      cell("S", { bold: true, align: "center", fill: "F2F2F2" }),
    ]),
    row([
      cell("", { align: "center" }),
      cell("", { align: "center" }),
      cell("", { align: "center" }),
      cell("", { align: "center" }),
      cell("1", { align: "center" }),
      cell("2", { align: "center" }),
      cell("3", { align: "center" }),
    ]),
    row([
      cell("4", { align: "center" }),
      cell("5", { align: "center" }),
      cell("6", { align: "center" }),
      cell("7", { align: "center" }),
      cell("8", { align: "center" }),
      cell("9", { align: "center" }),
      cell("10", { align: "center" }),
    ]),
    row([
      cell("11", { align: "center" }),
      cell("12", { align: "center" }),
      cell("13", { align: "center" }),
      cell("14", { align: "center" }),
      cell("15", { align: "center" }),
      cell("16", { align: "center" }),
      cell("17", { align: "center" }),
    ]),
    row([
      cell("18", { align: "center" }),
      cell("19", { align: "center" }),
      cell("20", { align: "center" }),
      cell("21", { align: "center" }),
      cell("22", { align: "center" }),
      cell("23", { align: "center" }),
      cell("24", { align: "center" }),
    ]),
    row([
      cell("25", { align: "center" }),
      cell("26", { align: "center" }),
      cell("27", { align: "center" }),
      cell("28", { align: "center" }),
      cell("29", { align: "center" }),
      cell("30", { align: "center" }),
      cell("31", { align: "center" }),
    ]),
  ]);
}

/** Build Matrix: 4-column criteria comparison matrix. */
function buildMatrix(): JSONContent {
  const widths = [130, 95, 95, 95];
  return table(widths, [
    row([
      cell("", { fill: "D9E1F2" }),
      cell("Criteria 1", { bold: true, align: "center", fill: "D9E1F2" }),
      cell("Criteria 2", { bold: true, align: "center", fill: "D9E1F2" }),
      cell("Criteria 3", { bold: true, align: "center", fill: "D9E1F2" }),
    ]),
    row([
      cell("Option A", { bold: true }),
      cell("High", { align: "center" }),
      cell("Medium", { align: "center" }),
      cell("Low", { align: "center" }),
    ]),
    row([
      cell("Option B", { bold: true }),
      cell("Medium", { align: "center" }),
      cell("High", { align: "center" }),
      cell("High", { align: "center" }),
    ]),
    row([
      cell("Option C", { bold: true }),
      cell("Low", { align: "center" }),
      cell("Low", { align: "center" }),
      cell("Medium", { align: "center" }),
    ]),
  ]);
}

/** Build Tabular List: Styled inventory / invoice table with total row. */
function buildTabularList(): JSONContent {
  const widths = [130, 170, 70, 90];
  return table(widths, [
    row([
      cell("Item", { bold: true, fill: "2F5597", color: "FFFFFF" }),
      cell("Description", { bold: true, fill: "2F5597", color: "FFFFFF" }),
      cell("Qty", { bold: true, align: "right", fill: "2F5597", color: "FFFFFF" }),
      cell("Unit Price", { bold: true, align: "right", fill: "2F5597", color: "FFFFFF" }),
    ]),
    row([
      cell("Widget A", { bold: true }),
      cell("Standard components"),
      cell("10", { align: "right" }),
      cell("$25.00", { align: "right" }),
    ]),
    row([
      cell("Widget B", { bold: true }),
      cell("Premium assembly"),
      cell("5", { align: "right" }),
      cell("$60.00", { align: "right" }),
    ]),
    row([
      cell("Widget C", { bold: true }),
      cell("Custom packaging"),
      cell("1", { align: "right" }),
      cell("$150.00", { align: "right" }),
    ]),
    row([
      cell("Total", { bold: true, fill: "F2F2F2", colspan: 2 }),
      cell("16", { bold: true, align: "right", fill: "F2F2F2" }),
      cell("$600.00", { bold: true, align: "right", fill: "F2F2F2" }),
    ]),
  ]);
}

/** Build Double Table: 2 side-by-side sub-tables. */
function buildDoubleTable(): JSONContent {
  const widths = [95, 115, 95, 115];
  return table(widths, [
    row([
      cell("Primary Key", { bold: true, fill: "E7E6E6" }),
      cell("Primary Value", { bold: true, fill: "E7E6E6" }),
      cell("Secondary Key", { bold: true, fill: "E7E6E6" }),
      cell("Secondary Value", { bold: true, fill: "E7E6E6" }),
    ]),
    row([cell("Alpha"), cell("100"), cell("Theta"), cell("500")]),
    row([cell("Beta"), cell("200"), cell("Iota"), cell("600")]),
    row([cell("Gamma"), cell("300"), cell("Kappa"), cell("700")]),
  ]);
}

/** Build Table with Subheadings: Sections with category banner rows. */
function buildSubheadings(): JSONContent {
  const widths = [180, 110, 110];
  return table(widths, [
    row([
      cell("Department / Role", { bold: true, fill: "D9E1F2" }),
      cell("Headcount", { bold: true, align: "right", fill: "D9E1F2" }),
      cell("Budget", { bold: true, align: "right", fill: "D9E1F2" }),
    ]),
    row([
      cell("Engineering Division", {
        bold: true,
        italic: true,
        fill: "F2F2F2",
        colspan: 3,
      }),
    ]),
    row([
      cell("Software Engineering"),
      cell("24", { align: "right" }),
      cell("$320,000", { align: "right" }),
    ]),
    row([
      cell("Infrastructure & DevOps"),
      cell("8", { align: "right" }),
      cell("$140,000", { align: "right" }),
    ]),
    row([
      cell("Product & Design", {
        bold: true,
        italic: true,
        fill: "F2F2F2",
        colspan: 3,
      }),
    ]),
    row([
      cell("Product Management"),
      cell("6", { align: "right" }),
      cell("$95,000", { align: "right" }),
    ]),
    row([cell("UX/UI Design"), cell("5", { align: "right" }), cell("$80,000", { align: "right" })]),
  ]);
}

export const QUICK_TABLE_PRESETS: readonly QuickTablePreset[] = [
  {
    id: "calendar1",
    name: "Calendar 1",
    nameZh: "日历 1",
    description: "Monthly calendar grid with centered title and days",
    descriptionZh: "居中标题与日期的月度日历网格",
    build: buildCalendar1,
  },
  {
    id: "calendar2",
    name: "Calendar 2",
    nameZh: "日历 2",
    description: "Monthly calendar with side title column",
    descriptionZh: "侧边标题列的月度日历网格",
    build: buildCalendar2,
  },
  {
    id: "matrix",
    name: "Matrix",
    nameZh: "矩阵",
    description: "Evaluation matrix with criteria headers",
    descriptionZh: "含评估标准标头的矩阵表格",
    build: buildMatrix,
  },
  {
    id: "tabularList",
    name: "Tabular List",
    nameZh: "表格列表",
    description: "Formatted inventory table with total row",
    descriptionZh: "带合计行的格式化列表表格",
    build: buildTabularList,
  },
  {
    id: "doubleTable",
    name: "Double Table",
    nameZh: "双表格",
    description: "Dual key-value comparison table",
    descriptionZh: "双列键值对比表格",
    build: buildDoubleTable,
  },
  {
    id: "subheadings",
    name: "Table with Subheadings",
    nameZh: "含子标题表格",
    description: "Grouped table with category subheading rows",
    descriptionZh: "带分类子标题横条的分组表格",
    build: buildSubheadings,
  },
] as const;

export function normalizeQuickTableId(id: string): string {
  const lower = id.toLowerCase().replace(/[-_\s]/g, "");
  if (lower === "calendar1") return "calendar1";
  if (lower === "calendar2") return "calendar2";
  if (lower === "matrix") return "matrix";
  if (lower === "tabularlist") return "tabularList";
  if (lower === "doubletable") return "doubleTable";
  if (lower === "subheadings" || lower === "subheading") return "subheadings";
  return "calendar1";
}

export function getQuickTableJson(id: string): JSONContent | null {
  const norm = normalizeQuickTableId(id);
  const preset = QUICK_TABLE_PRESETS.find((p) => p.id === norm);
  return preset ? preset.build() : null;
}

/** Convert presets to BuildingBlock items for the Building Blocks Organizer. */
export function getQuickTableBuildingBlocks(): BuildingBlock[] {
  return QUICK_TABLE_PRESETS.map((p) => ({
    id: `quick-table-${p.id}`,
    name: p.name,
    gallery: "custTables",
    category: "Built-In",
    description: p.description,
    savedAt: 0,
    insertMode: "paragraph",
    content: {
      openStart: 0,
      openEnd: 0,
      content: [p.build()],
    },
  }));
}

/** Generate a minimal valid Microsoft Excel (.xlsx) workbook binary bundle. */
export function createBlankExcelWorkbookBytes(): Uint8Array {
  const enc = new TextEncoder();
  return zipSync({
    "[Content_Types].xml": enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
        `</Relationships>`,
    ),
    "xl/workbook.xml": enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>` +
        `</workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
        `</Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<sheetData/>` +
        `</worksheet>`,
    ),
  });
}

/** Trigger download of an embedded OLE object's binary payload. */
export function downloadOleObject(bytes: Uint8Array, fileName: string): void {
  const isXlsx = fileName.endsWith(".xlsx");
  const mime = isXlsx
    ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    : "application/octet-stream";
  const blob = new Blob([bytes as any], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
}
