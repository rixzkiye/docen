import { describe, expect, it, vi } from "vitest";

import { updateDynamicFieldsBeforePrint } from "./commands/dialogs";
import { evaluateFormula, isCalculatedField, CALCULATED_FIELD_NAMES } from "./fields";

describe("Field Shading & Calculations: evaluateFormula (W7.4)", () => {
  it("evaluates basic arithmetic expressions", () => {
    expect(evaluateFormula("1 + 2 * 3")).toBe("7");
    expect(evaluateFormula("(10 - 4) / 2")).toBe("3");
    expect(evaluateFormula("2 ^ 3")).toBe("8");
    expect(evaluateFormula("10 % 3")).toBe("1");
    expect(evaluateFormula("15.5 + 4.5")).toBe("20");
  });

  it("evaluates built-in math and statistical functions", () => {
    expect(evaluateFormula("SUM(1, 2, 3, 4)")).toBe("10");
    expect(evaluateFormula("AVERAGE(10, 20, 30)")).toBe("20");
    expect(evaluateFormula("COUNT(5, 10, 15)")).toBe("3");
    expect(evaluateFormula("MIN(8, 3, 9, 2, 7)")).toBe("2");
    expect(evaluateFormula("MAX(8, 3, 9, 2, 7)")).toBe("9");
    expect(evaluateFormula("PRODUCT(2, 3, 4)")).toBe("24");
    expect(evaluateFormula("ROUND(3.14159, 2)")).toBe("3.14");
    expect(evaluateFormula("ABS(-42)")).toBe("42");
    expect(evaluateFormula("INT(7.85)")).toBe("7");
    expect(evaluateFormula("MOD(17, 5)")).toBe("2");
  });

  it("applies Word picture switches (\\# format)", () => {
    expect(evaluateFormula('123.456 \\# "0.00"')).toBe("123.46");
    expect(evaluateFormula('1234.5 \\# "$#,##0.00"')).toBe("$1,234.50");
    expect(evaluateFormula('0.25 \\# "0%"')).toBe("25%");
  });

  it("handles invalid or divide-by-zero safely", () => {
    expect(evaluateFormula("10 / 0")).toBe("!ZeroDivide");
    expect(evaluateFormula("INVALID_SYNTAX(")).toBe("!SyntaxError");
  });
});

describe("Field Shading: isCalculatedField (W7.4)", () => {
  it("identifies formulas and calculated dynamic fields", () => {
    expect(isCalculatedField("= 1 + 2")).toBe(true);
    expect(isCalculatedField("FORMULA SUM(A1:A5)")).toBe(true);
    expect(isCalculatedField("PAGE")).toBe(true);
    expect(isCalculatedField("NUMPAGES")).toBe(true);
    expect(isCalculatedField('DATE \\@ "yyyy-MM-dd"')).toBe(true);
    expect(isCalculatedField("TIME")).toBe(true);
    expect(isCalculatedField("AUTHOR")).toBe(true);
    expect(isCalculatedField("FILENAME")).toBe(true);
    expect(isCalculatedField("PAGEOF")).toBe(true);
  });

  it("returns false for non-calculated or structural fields", () => {
    expect(isCalculatedField('TOC \\o "1-3"')).toBe(false);
    expect(isCalculatedField('XE "Index entry"')).toBe(false);
    expect(isCalculatedField("UNKNOWN_FIELD")).toBe(false);
    expect(isCalculatedField("")).toBe(false);
  });

  it("exports all calculated field names in CALCULATED_FIELD_NAMES set", () => {
    expect(CALCULATED_FIELD_NAMES.has("=")).toBe(true);
    expect(CALCULATED_FIELD_NAMES.has("FORMULA")).toBe(true);
    expect(CALCULATED_FIELD_NAMES.has("PAGE")).toBe(true);
    expect(CALCULATED_FIELD_NAMES.has("NUMPAGES")).toBe(true);
    expect(CALCULATED_FIELD_NAMES.has("DATE")).toBe(true);
    expect(CALCULATED_FIELD_NAMES.has("TIME")).toBe(true);
  });
});

describe("Update Fields Before Print (W7.4)", () => {
  it("executes updateAllFields and update-toc prior to print compilation", () => {
    const mockDialogs = {
      updateAllFields: vi.fn(() => 5),
    };
    const mockCommands = {
      "update-toc": vi.fn(() => true),
      "update-figure-table": vi.fn(),
    };
    const pageOf = vi.fn((_pos: number) => 1);

    const result = updateDynamicFieldsBeforePrint(mockDialogs as any, mockCommands as any, pageOf);

    expect(mockDialogs.updateAllFields).toHaveBeenCalledTimes(1);
    expect(mockCommands["update-toc"]).toHaveBeenCalledWith(pageOf);
    expect(mockCommands["update-figure-table"]).toHaveBeenCalledWith(pageOf);
    expect(result).toEqual({ fieldsUpdated: 5, tocUpdated: true });
  });

  it("works gracefully when TOC or figure commands are absent", () => {
    const mockDialogs = {
      updateAllFields: vi.fn(() => 2),
    };
    const result = updateDynamicFieldsBeforePrint(mockDialogs as any, undefined);
    expect(mockDialogs.updateAllFields).toHaveBeenCalledTimes(1);
    expect(result).toEqual({ fieldsUpdated: 2, tocUpdated: false });
  });
});
