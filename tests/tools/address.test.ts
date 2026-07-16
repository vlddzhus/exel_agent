/**
 * Тесты для _shared/address.ts — единый источник адресной арифметики.
 *
 * Эти тесты критичны: ВСЕ инструменты зависят от этого модуля.
 */
import {
  columnToLetter,
  letterToColumn,
  parseCellAddress,
  parseRangeAddress,
  computeRange,
  mergeRangeAddress,
  isValidCellAddress,
  isValidRangeAddress,
  isValidFullAddress,
  isValidSheetName,
  normalizeAddress,
  splitFullAddress,
} from "../../src/taskpane/tools/_shared/address";

describe("columnToLetter / letterToColumn", () => {
  test("0-based: 0->A, 1->B, 25->Z, 26->AA", () => {
    expect(columnToLetter(0)).toBe("A");
    expect(columnToLetter(1)).toBe("B");
    expect(columnToLetter(25)).toBe("Z");
    expect(columnToLetter(26)).toBe("AA");
    expect(columnToLetter(27)).toBe("AB");
    expect(columnToLetter(51)).toBe("AZ");
    expect(columnToLetter(52)).toBe("BA");
    expect(columnToLetter(701)).toBe("ZZ");
    expect(columnToLetter(702)).toBe("AAA");
    expect(columnToLetter(16383)).toBe("XFD");
  });

  test("letterToColumn обратное преобразование", () => {
    expect(letterToColumn("A")).toBe(0);
    expect(letterToColumn("Z")).toBe(25);
    expect(letterToColumn("AA")).toBe(26);
    expect(letterToColumn("AZ")).toBe(51);
    expect(letterToColumn("BA")).toBe(52);
    expect(letterToColumn("ZZ")).toBe(701);
    expect(letterToColumn("AAA")).toBe(702);
    expect(letterToColumn("XFD")).toBe(16383);
  });

  test("round-trip для всех валидных", () => {
    for (const letter of ["A", "B", "Z", "AA", "AZ", "BA", "ZZ", "AAA", "XFD"]) {
      expect(columnToLetter(letterToColumn(letter))).toBe(letter);
    }
  });

  test("letterToColumn нечувствителен к регистру", () => {
    expect(letterToColumn("a")).toBe(0);
    expect(letterToColumn("aa")).toBe(26);
    expect(letterToColumn("Xfd")).toBe(16383);
  });

  test("columnToLetter бросает RangeError для невалидных", () => {
    expect(() => columnToLetter(-1)).toThrow(RangeError);
    expect(() => columnToLetter(18279)).toThrow(RangeError);
    expect(() => columnToLetter(1.5)).toThrow(RangeError);
    expect(() => columnToLetter(NaN)).toThrow(RangeError);
  });

  test("letterToColumn бросает RangeError для невалидных", () => {
    expect(() => letterToColumn("")).toThrow(RangeError);
    expect(() => letterToColumn("A1")).toThrow(RangeError);
    expect(() => letterToColumn("$A")).toThrow(RangeError);
    expect(() => letterToColumn("АА")).toThrow(RangeError);
  });
});

describe("parseCellAddress", () => {
  test("парсит валидные", () => {
    expect(parseCellAddress("A1")).toEqual({ col: "A", row: 1 });
    expect(parseCellAddress("AA123")).toEqual({ col: "AA", row: 123 });
    expect(parseCellAddress("XFD1048576")).toEqual({ col: "XFD", row: 1048576 });
  });

  test("бросает Error для невалидных", () => {
    expect(() => parseCellAddress("1A")).toThrow();
    expect(() => parseCellAddress("A")).toThrow();
    expect(() => parseCellAddress("123")).toThrow();
    expect(() => parseCellAddress("")).toThrow();
    expect(() => parseCellAddress("A1:B2")).toThrow();
  });
});

describe("parseRangeAddress", () => {
  test("парсит валидные", () => {
    expect(parseRangeAddress("A1:B2")).toEqual({
      startCol: "A", startRow: 1, endCol: "B", endRow: 2,
    });
    expect(parseRangeAddress("AA10:AB20")).toEqual({
      startCol: "AA", startRow: 10, endCol: "AB", endRow: 20,
    });
  });

  test("бросает Error для невалидных", () => {
    expect(() => parseRangeAddress("A1")).toThrow();
    expect(() => parseRangeAddress("A1:B2:C3")).toThrow();
    expect(() => parseRangeAddress("A1:")).toThrow();
    expect(() => parseRangeAddress(":B2")).toThrow();
  });
});

describe("computeRange", () => {
  test("базовые случаи", () => {
    expect(computeRange("A1", 1, 1)).toBe("A1:A1");
    expect(computeRange("A1", 9, 5)).toBe("A1:E9");
    expect(computeRange("A11", 9, 5)).toBe("A11:E19");
    expect(computeRange("B3", 10, 3)).toBe("B3:D12");
  });

  test("multi-letter", () => {
    expect(computeRange("AA1", 5, 3)).toBe("AA1:AC5");
    expect(computeRange("AZ10", 2, 4)).toBe("AZ10:BC11");
  });

  test("бросает Error для неположительных", () => {
    expect(() => computeRange("A1", 0, 5)).toThrow();
    expect(() => computeRange("A1", 5, 0)).toThrow();
    expect(() => computeRange("A1", -1, 5)).toThrow();
    expect(() => computeRange("A1", 5, 1.5)).toThrow();
  });

  test("бросает Error для невалидного startCell", () => {
    expect(() => computeRange("1A", 5, 5)).toThrow();
    expect(() => computeRange("", 5, 5)).toThrow();
  });
});

describe("mergeRangeAddress", () => {
  test("базовое объединение", () => {
    expect(mergeRangeAddress("A1:B2", "C3:D4")).toBe("A1:D4");
    expect(mergeRangeAddress("A1:A1", "C3:C3")).toBe("A1:C3");
  });

  test("КРИТИЧНО: числовое сравнение Z < AA (не строковое!)", () => {
    const result = mergeRangeAddress("A1:Z10", "AA1:AB10");
    expect(result).toBe("A1:AB10");
  });

  test("вложенные диапазоны", () => {
    expect(mergeRangeAddress("A1:Z100", "B5:C10")).toBe("A1:Z100");
  });

  test("та же колонка", () => {
    expect(mergeRangeAddress("A1:A5", "A3:A10")).toBe("A1:A10");
  });
});

describe("isValidCellAddress", () => {
  test.each([
    ["A1", true], ["AA123", true], ["XFD1048576", true], ["a1", true],
    ["", false], ["A", false], ["1", false], ["1A", false],
    ["A1:B2", false], ["Лист!A1", false],
  ])("isValidCellAddress(%j) === %s", (input, expected) => {
    expect(isValidCellAddress(input)).toBe(expected);
  });
});

describe("isValidRangeAddress", () => {
  test.each([
    ["A1", true], ["A1:B2", true], ["AA1:AB2", true],
    ["A1:B2:C3", false], ["A1:", false], ["A", false], ["", false],
  ])("isValidRangeAddress(%j) === %s", (input, expected) => {
    expect(isValidRangeAddress(input)).toBe(expected);
  });
});

describe("isValidSheetName", () => {
  test.each([
    ["Лист1", true], ["МойЛист", true], ["'Мой лист'", true], ["Sheet", true],
    ["", false], ["a".repeat(32), false],
    ["Лист:1", false], ["Лист/1", false], ["Лист?1", false], ["Лист*1", false],
  ])("isValidSheetName(%j) === %s", (input, expected) => {
    expect(isValidSheetName(input)).toBe(expected);
  });
});

describe("isValidFullAddress — path traversal защита", () => {
  test.each([
    ["A1", true], ["A1:B2", true], ["Лист1!A1", true],
    ["'Мой лист'!A1:B2", true], ["Лист1!A1:B2", true],
    ["../../etc/passwd", false], ["Лист!../../secret", false],
    ["A1/../../etc", false], ["", false], ["!A1", false], ["Лист!", false],
  ])("isValidFullAddress(%j) === %s", (input, expected) => {
    expect(isValidFullAddress(input)).toBe(expected);
  });

  test("отклоняет слишком длинные адреса", () => {
    expect(isValidFullAddress("A".repeat(501))).toBe(false);
  });
});

describe("normalizeAddress", () => {
  test("поднимает регистр", () => {
    expect(normalizeAddress("a1")).toBe("A1");
    expect(normalizeAddress("a1:b2")).toBe("A1:B2");
  });

  test("сохраняет имя листа", () => {
    expect(normalizeAddress("Лист1!a1")).toBe("Лист1!A1");
    expect(normalizeAddress("'Мой лист'!a1:b2")).toBe("'Мой лист'!A1:B2");
  });
});

describe("splitFullAddress", () => {
  test("без листа", () => {
    expect(splitFullAddress("A1:B2")).toEqual({ rangeAddress: "A1:B2" });
  });

  test("с листом без кавычек", () => {
    expect(splitFullAddress("Лист1!A1")).toEqual({
      sheetName: "Лист1", rangeAddress: "A1",
    });
  });

  test("с листом в кавычках (снимает кавычки)", () => {
    expect(splitFullAddress("'Мой лист'!A1:B2")).toEqual({
      sheetName: "Мой лист", rangeAddress: "A1:B2",
    });
  });
});
