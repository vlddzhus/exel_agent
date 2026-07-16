/**
 * Тесты для детекторов аномалий из read.ts (R5 findAnomalies).
 *
 * Чистые функции: isErrorValue, detectOutliersZScore, detectOutliersIQR,
 * findColumnAnomalies. Сам инструмент findAnomaliesTool работает через Excel.run
 * и тестируется отдельно на office-addin-mock (как getWorkbookOverview).
 *
 * См. docs/03-TOOLS-SPEC.md §1 R5, эталонный сценарий №19.
 */
import {
  isErrorValue,
  detectOutliersZScore,
  detectOutliersIQR,
  findColumnAnomalies,
  type AnomalyKind,
} from "../../src/taskpane/tools/read";

// ---------------------------------------------------------------------------
// isErrorValue — значения-ошибки Excel (RU + EN)
// ---------------------------------------------------------------------------

describe("isErrorValue", () => {
  test.each([
    ["#Н/Д", true], // RU
    ["#ДЕЛ/0!", true], // RU
    ["#ЗНАЧ!", true], // RU
    ["#REF!", true], // EN
    ["#DIV/0!", true], // EN
    ["#VALUE!", true],
    ["#NAME?", true],
    ["#NUM!", true],
    ["#SPILL!", true],
    ["#CALC!", true],
    [42, false],
    ["", false],
    ["текст", false],
    [null, false],
    [undefined, false],
  ])("isErrorValue(%j) === %s", (input, expected) => {
    expect(isErrorValue(input)).toBe(expected);
  });
});

// ---------------------------------------------------------------------------
// detectOutliersZScore — только при ≥8 числах
// ---------------------------------------------------------------------------

describe("detectOutliersZScore", () => {
  test("меньше 8 чисел → пусто (недостаточно данных)", () => {
    expect(detectOutliersZScore([1, 2, 3, 4, 5, 6, 7])).toEqual([]);
  });

  test("один явный выброс среди нормальных значений", () => {
    // 7 значений около 10 + один 1000
    const nums = [10, 11, 9, 10, 12, 9, 11, 1000];
    const result = detectOutliersZScore(nums, 3);
    expect(result).toEqual([7]);
  });

  test("нет выбросов в однородной выборке", () => {
    const nums = [10, 11, 9, 10, 12, 9, 11, 10];
    expect(detectOutliersZScore(nums, 3)).toEqual([]);
  });

  test("std=0 (все равны) → пусто", () => {
    expect(detectOutliersZScore([5, 5, 5, 5, 5, 5, 5, 5])).toEqual([]);
  });

  test("выбросы с обеих сторон от среднего", () => {
    const nums = [10, 10, 10, 10, 10, 10, 10, 1000, -1000];
    const result = detectOutliersZScore(nums, 3);
    expect(result).toContain(7);
    expect(result).toContain(8);
  });
});

// ---------------------------------------------------------------------------
// detectOutliersIQR — для малых выборок
// ---------------------------------------------------------------------------

describe("detectOutliersIQR", () => {
  test("меньше 4 чисел → пусто", () => {
    expect(detectOutliersIQR([1, 2, 3])).toEqual([]);
  });

  test("явный выброс сверху в малой выборке", () => {
    const nums = [1, 2, 3, 4, 5, 6, 100];
    const result = detectOutliersIQR(nums);
    expect(result).toEqual([6]);
  });

  test("явный выброс снизу", () => {
    const nums = [10, 20, 30, 40, 50, 60, -100];
    const result = detectOutliersIQR(nums);
    expect(result).toContain(6);
  });

  test("нет выбросов в плотной выборке", () => {
    const nums = [1, 2, 3, 4, 5, 6, 7];
    expect(detectOutliersIQR(nums)).toEqual([]);
  });

  test("IQR=0 (Q1=Q3) → пусто", () => {
    expect(detectOutliersIQR([5, 5, 5, 5, 5, 5, 5])).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// findColumnAnomalies — комплексный детектор
// ---------------------------------------------------------------------------

describe("findColumnAnomalies", () => {
  test("пустые ячейки детектируются", () => {
    const result = findColumnAnomalies([1, null, 3, "", 5], "X", 0);
    const empties = result.filter((a) => a.kind === "empty");
    expect(empties.length).toBe(2);
    expect(empties.map((a) => a.row).sort((a, b) => a - b)).toEqual([2, 4]);
  });

  test("ошибки Excel детектируются", () => {
    const result = findColumnAnomalies(
      [1, "#Н/Д", 3, "#ДЕЛ/0!", 5],
      "X",
      0,
    );
    const errors = result.filter((a) => a.kind === "error");
    expect(errors.length).toBe(2);
    expect(errors.map((a) => a.row).sort((a, b) => a - b)).toEqual([2, 4]);
  });

  test("дубли в email-колонке детектируются", () => {
    const result = findColumnAnomalies(
      [
        "ivan@example.com",
        "maria@example.com",
        "ivan@example.com", // дубль строки 1
        "ivan@example.com", // дубль строки 1
      ],
      "Email",
      0,
    );
    const dups = result.filter((a) => a.kind === "duplicate");
    expect(dups.length).toBe(2);
    expect(dups.map((a) => a.row).sort((a, b) => a - b)).toEqual([3, 4]);
    // detail ссылается на первую встреченную строку
    expect(dups.find((a) => a.row === 3)?.detail).toContain("строки 1");
  });

  test("дубли НЕ детектируются в произвольном тексте (не ID)", () => {
    // "text"-колонка: одинаковые значения — нормально, не аномалия
    const result = findColumnAnomalies(
      ["привет", "мир", "привет", "привет"],
      "Комментарий",
      0,
    );
    const dups = result.filter((a) => a.kind === "duplicate");
    expect(dups.length).toBe(0);
  });

  test("дубли в ИНН-колонке детектируются (ID-тип)", () => {
    const result = findColumnAnomalies(
      ["7707083893", "7707083893", "500100732259"],
      "ИНН",
      0,
    );
    const dups = result.filter((a) => a.kind === "duplicate");
    expect(dups.length).toBe(1);
    expect(dups[0].row).toBe(2);
  });

  test("числовой выброс через IQR в малой выборке", () => {
    // 6 значений около 10 + выброс 1000 → IQR сработает (Z-score нужен ≥8)
    const result = findColumnAnomalies(
      [10, 11, 9, 10, 12, 1000],
      "Сумма",
      0,
    );
    const outliers = result.filter((a) => a.kind === "outlier");
    expect(outliers.length).toBe(1);
    expect(outliers[0].row).toBe(6);
  });

  test("несоответствие типу: строка в числовой колонке", () => {
    const result = findColumnAnomalies(
      [1, 2, "не число", 4, 5],
      "Кол-во",
      0,
    );
    const mismatches = result.filter((a) => a.kind === "type_mismatch");
    expect(mismatches.length).toBe(1);
    expect(mismatches[0].row).toBe(3);
    expect(mismatches[0].detail).toContain("integer");
  });

  test("несоответствие НЕ детектируется в text/mixed колонке", () => {
    // text — нет «ожидаемого» типа, type_mismatch не применим
    const result = findColumnAnomalies(
      ["Иван", "Мария", "Петр", "123", ""],
      "Имя",
      0,
    );
    const mismatches = result.filter((a) => a.kind === "type_mismatch");
    expect(mismatches.length).toBe(0);
  });

  test("числовая выброс через Z-score в большой выборке", () => {
    const result = findColumnAnomalies(
      [10, 11, 9, 10, 12, 9, 11, 1000],
      "X",
      0,
    );
    const outliers = result.filter((a) => a.kind === "outlier");
    expect(outliers.length).toBe(1);
    expect(outliers[0].row).toBe(8);
  });

  test("пустая колонка → только empty-аномалии", () => {
    const result = findColumnAnomalies(
      [null, "", undefined, null],
      "X",
      0,
    );
    expect(result.every((a) => a.kind === "empty")).toBe(true);
    expect(result.length).toBe(4);
  });

  test("аномалии содержат корректные column/header", () => {
    const result = findColumnAnomalies([1, null, 3], "Кол-во", 2);
    expect(result[0].column).toBe(2);
    expect(result[0].header).toBe("Кол-во");
  });

  test("чистая колонка без аномалий", () => {
    const result = findColumnAnomalies([1, 2, 3, 4, 5], "X", 0);
    expect(result.length).toBe(0);
  });

  test("row начинается с 1 (первая строка данных)", () => {
    const result = findColumnAnomalies([null, 1, null], "X", 0);
    const rows = result.map((a) => a.row).sort((a, b) => a - b);
    expect(rows).toEqual([1, 3]);
  });

  test("комбинированная колонка: несколько видов аномалий", () => {
    // integer-колонка: [1, "#Н/Д", 2, null, 9999, 3, 4, 5, 6, 7]
    const result = findColumnAnomalies(
      [1, "#Н/Д", 2, null, 9999, 3, 4, 5, 6, 7],
      "X",
      0,
    );
    const kinds = new Set(result.map((a) => a.kind));
    expect(kinds.has("error")).toBe(true);
    expect(kinds.has("empty")).toBe(true);
    // 9999 — выброс относительно 1..7
    expect(kinds.has("outlier")).toBe(true);
  });

  test("все виды аномалий валидны", () => {
    const result = findColumnAnomalies([1, null, "x", 3], "X", 0);
    const validKinds: AnomalyKind[] = [
      "empty",
      "duplicate",
      "outlier",
      "type_mismatch",
      "error",
    ];
    for (const a of result) {
      expect(validKinds).toContain(a.kind);
    }
  });
});
