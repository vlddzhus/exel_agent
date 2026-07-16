/**
 * Тесты для чистых helper-функций из read.ts: parseNumeric, detectCellType.
 *
 * Эти функции критичны для корректной обработки RU/EN числовых данных.
 * Без них agent неправильно понимает "1 234,5" (RU) или "1,234.5" (EN).
 */
import {
  parseNumeric,
  detectCellType,
} from "../../src/taskpane/tools/read";

// ---------------------------------------------------------------------------
// parseNumeric
// ---------------------------------------------------------------------------

describe("parseNumeric — распознавание чисел", () => {
  test("число как number", () => {
    expect(parseNumeric(42)).toBe(42);
    expect(parseNumeric(3.14)).toBe(3.14);
    expect(parseNumeric(-7)).toBe(-7);
    expect(parseNumeric(0)).toBe(0);
  });

  test("число как string EN", () => {
    expect(parseNumeric("42")).toBe(42);
    expect(parseNumeric("3.14")).toBe(3.14);
    expect(parseNumeric("-7")).toBe(-7);
  });

  test("число как string RU с запятой", () => {
    expect(parseNumeric("3,14")).toBe(3.14);
    expect(parseNumeric("1234,5")).toBe(1234.5);
    expect(parseNumeric("-0,5")).toBe(-0.5);
  });

  test("число с разделителем разрядов RU (пробел)", () => {
    expect(parseNumeric("1 234")).toBe(1234);
    expect(parseNumeric("1 234,5")).toBe(1234.5);
    expect(parseNumeric("10 000")).toBe(10000);
  });

  test("число с разделителем разрядов EN (запятая)", () => {
    expect(parseNumeric("1,234.5")).toBe(1234.5);
    expect(parseNumeric("1,234,567.89")).toBe(1234567.89);
  });

  test("процент", () => {
    expect(parseNumeric("50%")).toBeCloseTo(0.5, 10);
    expect(parseNumeric("12.5%")).toBeCloseTo(0.125, 10);
    expect(parseNumeric("3,14%")).toBeCloseTo(0.0314, 10);
  });

  test("null/undefined/пустое → null", () => {
    expect(parseNumeric(null)).toBeNull();
    expect(parseNumeric(undefined)).toBeNull();
    expect(parseNumeric("")).toBeNull();
    expect(parseNumeric("   ")).toBeNull();
  });

  test("не число → null", () => {
    expect(parseNumeric("hello")).toBeNull();
    expect(parseNumeric("abc123")).toBeNull();
    expect(parseNumeric(true)).toBeNull();
    expect(parseNumeric({})).toBeNull();
  });

  test("NaN/Infinity → null", () => {
    expect(parseNumeric(NaN)).toBeNull();
    expect(parseNumeric(Infinity)).toBeNull();
    expect(parseNumeric(-Infinity)).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// detectCellType
// ---------------------------------------------------------------------------

describe("detectCellType — определение типа ячейки", () => {
  test("пустая ячейка", () => {
    expect(detectCellType(null, "", "General")).toBe("empty");
    expect(detectCellType(undefined, "", "General")).toBe("empty");
    expect(detectCellType("", "", "General")).toBe("empty");
  });

  test("число", () => {
    expect(detectCellType(42, "", "General")).toBe("number");
    expect(detectCellType(3.14, "", "0.00")).toBe("number");
    expect(detectCellType(-7, "", "#,##0")).toBe("number");
  });

  test("дата (по numberFormat)", () => {
    expect(detectCellType(44197, "", "DD.MM.YYYY")).toBe("date");
    expect(detectCellType(44197, "", "YYYY-MM-DD")).toBe("date");
    expect(detectCellType(44197, "", "MM/DD/YY")).toBe("date");
  });

  test("процент — НЕ дата (даже если есть 'd' в формате)", () => {
    expect(detectCellType(0.5, "", "0%")).toBe("number");
  });

  test("булево", () => {
    expect(detectCellType(true, "", "General")).toBe("boolean");
    expect(detectCellType(false, "", "General")).toBe("boolean");
  });

  test("текст", () => {
    expect(detectCellType("hello", "", "General")).toBe("string");
    expect(detectCellType("Иванов", "", "@")).toBe("string");
  });

  test("ошибки Excel", () => {
    expect(detectCellType("#REF!", "", "General")).toBe("error");
    expect(detectCellType("#DIV/0!", "", "General")).toBe("error");
    expect(detectCellType("#N/A", "", "General")).toBe("error");
    expect(detectCellType("#NAME?", "", "General")).toBe("error");
    expect(detectCellType("#VALUE!", "", "General")).toBe("error");
  });

  test("формула (по formula starting with =)", () => {
    expect(detectCellType(100, "=SUM(A1:A10)", "General")).toBe("formula");
    expect(detectCellType("Иван", "=IF(A1>0,\"да\",\"нет\")", "General")).toBe("formula");
  });
});
