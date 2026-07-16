/**
 * Тесты для write.ts — чистые helper-функции write-инструментов.
 *
 * Все функции, тестируемые здесь, не требуют Excel.run — они чистые.
 * Тесты на реальную запись в Excel (через defineTool) — в отдельных
 * наборах с office-addin-mock.
 */
import {
  normalizeRows,
  escapeFormulaInjection,
  escapeValues,
  generateFillValues,
  prepareFormulas,
} from "../../src/taskpane/tools/write";

// ===========================================================================
// normalizeRows
// ===========================================================================

describe("normalizeRows — выравнивание длин строк", () => {
  test("строки одинаковой длины → без изменений", () => {
    expect(
      normalizeRows([
        [1, 2],
        [3, 4],
      ]),
    ).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  test("строки разной длины → дополняется пустой строкой", () => {
    expect(normalizeRows([[1, 2, 3], [4], [5, 6]])).toEqual([
      [1, 2, 3],
      [4, "", ""],
      [5, 6, ""],
    ]);
  });

  test("пустой массив → пустой массив", () => {
    expect(normalizeRows([])).toEqual([]);
  });

  test("не-массив как элемент → оборачивается в массив", () => {
    const input: unknown[][] = [[1, 2], "single" as unknown as unknown[], [3]];
    expect(normalizeRows(input)).toEqual([
      [1, 2],
      ["single", ""],
      [3, ""],
    ]);
  });

  test("массив с null → null сохраняется, паддинг не нужен (одна колонка)", () => {
    expect(normalizeRows([[1], [null]])).toEqual([[1], [null]]);
  });

  test("одна строка → без изменений", () => {
    expect(normalizeRows([[1, 2, 3]])).toEqual([[1, 2, 3]]);
  });
});

// ===========================================================================
// escapeFormulaInjection
// ===========================================================================

describe("escapeFormulaInjection — экранирование инъекций формул", () => {
  test("строка с '=' → экранируется апострофом", () => {
    expect(escapeFormulaInjection("=SUM(A1:A10)")).toBe("'=SUM(A1:A10)");
  });

  test("строка с '+' → экранируется", () => {
    expect(escapeFormulaInjection("+123")).toBe("'+123");
  });

  test("строка с '-' → экранируется", () => {
    expect(escapeFormulaInjection("-text")).toBe("'-text");
  });

  test("строка с '@' → экранируется", () => {
    expect(escapeFormulaInjection("@command")).toBe("'@command");
  });

  test("обычный текст → без изменений", () => {
    expect(escapeFormulaInjection("Hello")).toBe("Hello");
    expect(escapeFormulaInjection("A1")).toBe("A1");
    expect(escapeFormulaInjection("")).toBe("");
  });

  test("число → без изменений", () => {
    expect(escapeFormulaInjection(42)).toBe(42);
    expect(escapeFormulaInjection(3.14)).toBe(3.14);
    expect(escapeFormulaInjection(0)).toBe(0);
  });

  test("boolean → без изменений", () => {
    expect(escapeFormulaInjection(true)).toBe(true);
    expect(escapeFormulaInjection(false)).toBe(false);
  });

  test("null → без изменений", () => {
    expect(escapeFormulaInjection(null)).toBeNull();
  });
});

// ===========================================================================
// escapeValues
// ===========================================================================

describe("escapeValues — экранирование 2D-массива", () => {
  test("все ячейки экранируются корректно", () => {
    const input = [
      ["=1+2", "text", 42],
      ["+command", "-value", "@ref"],
    ];
    const expected = [
      ["'=1+2", "text", 42],
      ["'+command", "'-value", "'@ref"],
    ];
    expect(escapeValues(input)).toEqual(expected);
  });

  test("пустой массив → пустой массив", () => {
    expect(escapeValues([])).toEqual([]);
  });

  test("только безопасные значения → без изменений", () => {
    const input = [
      ["a", "b"],
      [1, 2],
      [true, null],
    ];
    expect(escapeValues(input)).toEqual(input);
  });
});

// ===========================================================================
// generateFillValues
// ===========================================================================

describe("generateFillValues — генерация значений для заполнения", () => {
  test("progression 3×1 с шагом 1", () => {
    const result = generateFillValues(3, 1, "progression");
    expect(result).toEqual([[1], [2], [3]]);
  });

  test("progression 1×4 с шагом 1", () => {
    const result = generateFillValues(1, 4, "progression");
    expect(result).toEqual([[1, 2, 3, 4]]);
  });

  test("progression с startValue=10", () => {
    const result = generateFillValues(3, 1, "progression", { startValue: 10 });
    expect(result).toEqual([[10], [11], [12]]);
  });

  test("progression с шагом 5", () => {
    const result = generateFillValues(4, 1, "progression", {
      startValue: 0,
      step: 5,
    });
    expect(result).toEqual([[0], [5], [10], [15]]);
  });

  test("progression 2×3 (табличная прогрессия)", () => {
    const result = generateFillValues(2, 3, "progression", {
      startValue: 1,
      step: 1,
    });
    expect(result).toEqual([
      [1, 2, 3],
      [4, 5, 6],
    ]);
  });

  test("copy — одно значение во все ячейки", () => {
    const result = generateFillValues(2, 2, "copy", { fillValue: "test" });
    expect(result).toEqual([
      ["test", "test"],
      ["test", "test"],
    ]);
  });

  test("copy без fillValue → пустая строка", () => {
    const result = generateFillValues(2, 1, "copy");
    expect(result).toEqual([[""], [""]]);
  });

  test("value — аналогично copy", () => {
    const result = generateFillValues(2, 2, "value", { fillValue: 0 });
    expect(result).toEqual([
      [0, 0],
      [0, 0],
    ]);
  });

  test("value без fillValue → startValue (по умолчанию 1)", () => {
    const result = generateFillValues(2, 1, "value");
    expect(result).toEqual([[1], [1]]);
  });
});

// ===========================================================================
// prepareFormulas
// ===========================================================================

describe("prepareFormulas — подготовка формул (FormulaGuardian + allowlist)", () => {
  test("простая формула SUM", () => {
    const result = prepareFormulas([["SUM(A1:A10)"]]);
    expect(result.ok).toBe(true);
    expect(result.formulas).toEqual([["=SUM(A1:A10)"]]);
  });

  test("формула с ведущим '='", () => {
    const result = prepareFormulas([["=SUM(A1:A10)"]]);
    expect(result.ok).toBe(true);
    expect(result.formulas).toEqual([["=SUM(A1:A10)"]]);
  });

  test("2D массив формул", () => {
    const result = prepareFormulas([
      ["SUM(A1:A3)", "AVERAGE(B1:B3)"],
      ["MAX(C1:C3)", "MIN(D1:D3)"],
    ]);
    expect(result.ok).toBe(true);
    expect(result.formulas).toEqual([
      ["=SUM(A1:A3)", "=AVERAGE(B1:B3)"],
      ["=MAX(C1:C3)", "=MIN(D1:D3)"],
    ]);
  });

  test("пустая ячейка в массиве → пустая строка", () => {
    const result = prepareFormulas([["SUM(A1:A3)", ""]]);
    expect(result.ok).toBe(true);
    expect(result.formulas).toEqual([["=SUM(A1:A3)", ""]]);
  });

  test("несбалансированные скобки → FORMULA_INVALID", () => {
    const result = prepareFormulas([["SUM(A1:A3"]]);
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("FORMULA_INVALID");
  });

  test("заблокированная функция WEBSERVICE → FORMULA_BLOCKED", () => {
    const result = prepareFormulas([['WEBSERVICE("https://evil.com")']]);
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("FORMULA_BLOCKED");
  });

  test("HYPERLINK → FORMULA_CONFIRM (автономный режим)", () => {
    const result = prepareFormulas([
      ['HYPERLINK("https://example.com", "Link")'],
    ]);
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("FORMULA_CONFIRM");
  });

  test("формула с авто-исправлением B8B6 → B8*B6", () => {
    const result = prepareFormulas([["B8B6"]]);
    expect(result.ok).toBe(true);
    expect(result.formulas).toEqual([["=B8*B6"]]);
  });

  test("пустой массив → пустой результат", () => {
    const result = prepareFormulas([[]]);
    expect(result.ok).toBe(true);
    expect(result.formulas).toEqual([[]]);
  });

  test("русская функция СУММ → валидна", () => {
    const result = prepareFormulas([["СУММ(A1:A10)"]]);
    expect(result.ok).toBe(true);
  });

  test("смешанный: валидные + невалидные → первая ошибка", () => {
    const result = prepareFormulas([["SUM(A1:A3)", 'WEBSERVICE("bad")']]);
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("FORMULA_BLOCKED");
  });
});
