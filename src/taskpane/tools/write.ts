/**
 * write.ts — Write-инструменты агента (категория W в docs/03-TOOLS-SPEC.md §1).
 *
 * Неделя 3 Фазы 1. Принципы те же, что в read.ts (см. шапку read.ts):
 *   - Каждый инструмент = defineTool с riskLevel/requiresUndo/estimateCells.
 *   - riskLevel "moderate" — меняет данные, требует undo-снапшот ДО.
 *   - Возвращает ToolResult с summary на русском.
 *   - Валидация через _shared/validation, адреса через _shared/address.
 *   - Один context.sync() на пакет (P5).
 *   - Undo через undoManager.createBackup ДО записи (P4 data safety).
 *
 * Инструменты:
 *   W1 setValues    — записать 2D-массив с валидацией размера и нормализацией.
 *   W2 setFormula   — записать формулы с ОБЯЗАТЕЛЬНОЙ проверкой FormulaGuardian.
 *   W3 fillRange    — заполнить по шаблону: прогрессия, копирование, константа.
 *   W4 appendRows   — дописать строки вниз под usedRange.
 *   W5 clearRange   — очистить значения/формулы/форматы.
 */

import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";
import { assertCellLimit, SAFE_CELL_LIMIT } from "./_shared/performance";
import { withPerformanceGuard } from "./_shared/performance";
import { undoManager } from "./backup";
import { validateFormula } from "./formula-guardian";
import { analyzeFormulaRisk } from "./formula-allowlist";

// ============================================================================
// Helpers — чистые функции для unit-тестов
// ============================================================================

/**
 * Выравнивает длины строк 2D-массива до максимума, заполняя недостающие ячейки
 * пустой строкой. Фикс бага в legacy setValues:(values[0].length падал на
 * неровных строках, docs/03-TOOLS-SPEC.md §2 W1 проблема 1).
 *
 * [[1,2,3],[4],[5,6]] → [[1,2,3],[4,"",""],[5,6,""]]
 */
export function normalizeRows(values: unknown[][]): unknown[][] {
  if (!Array.isArray(values) || values.length === 0) return [];
  const maxCols = values.reduce(
    (max, row) => (Array.isArray(row) ? Math.max(max, row.length) : max),
    0,
  );
  return values.map((row) => {
    const r = Array.isArray(row) ? row : [row];
    const padded = [...r];
    while (padded.length < maxCols) padded.push("");
    return padded;
  });
}

/**
 * Экранирует строковые значения, начинающиеся с = + - @, добавляя апостроф.
 * Защита от случайной инъекции формул в setValues (в Excel ячейка с "=" в начале
 * становится формулой; "+"/"-"/"@" — устаревшие префиксы совместимости).
 *
 * Числа, boolean, null не трогаются. Формулы должны идти через setFormula (W2),
 * а не через setValues.
 *
 * Унаследовано из legacy setValues (range-tools.ts) с тем же правилом.
 */
export function escapeFormulaInjection(value: unknown): unknown {
  if (typeof value === "string" && /^[=+\-@]/.test(value)) {
    return `'${value}`;
  }
  return value;
}

/**
 * Применяет escapeFormulaInjection ко всем ячейкам 2D-массива.
 */
export function escapeValues(values: unknown[][]): unknown[][] {
  return values.map((row) => row.map(escapeFormulaInjection));
}

/**
 * Результат подготовки формул: либо все валидны (с исправлениями), либо первая
 * ошибка останавливает выполнение (P4 — нельзя записывать невалидное).
 */
export interface FormulaPreparationResult {
  ok: boolean;
  /** Исправленные формулы (если ok=true). */
  formulas: string[][];
  /** Первая ошибка (если ok=false). */
  error?: { code: string; message: string; formula?: string };
}

/**
 * Массово валидирует формулы через FormulaGuardian + allowlist.
 *
 * Для каждой формулы:
 *   1. validateFormula — скобки, uppercase EN-функций, авто-исправление "*" между
 *      смежными cell-refs (см. formula-guardian.ts).
 *   2. analyzeFormulaRisk — блокировка WEBSERVICE/EXEC и подтверждение HYPERLINK
 *      (см. formula-allowlist.ts).
 *
 * Блокированные формулы → FORMULA_BLOCKED. Confirm-required → FORMULA_CONFIRM
 * (пока не критично для автономного агента; возвращаем как ошибку, агент
 * попросит подтверждение через UI в Фазе 3).
 *
 * @param formulas 2D массив БЕЗ ведущего "=" (или с "=" — оба варианта ок).
 */
export function prepareFormulas(formulas: string[][]): FormulaPreparationResult {
  const result: string[][] = [];
  for (let r = 0; r < formulas.length; r++) {
    const row = formulas[r] ?? [];
    const outRow: string[] = [];
    for (let c = 0; c < row.length; c++) {
      const raw = row[c] ?? "";
      // Срезаем ведущий "=" — validateFormula работает без него.
      const body = raw.startsWith("=") ? raw.slice(1).trim() : raw.trim();
      if (!body) {
        outRow.push("");
        continue;
      }

      const validation = validateFormula(body);
      if (!validation.valid) {
        return {
          ok: false,
          formulas: [],
          error: {
            code: "FORMULA_INVALID",
            message: validation.error ?? "Невалидная формула",
            formula: raw,
          },
        };
      }

      const fixedBody = validation.fixedFormula ?? body;
      // Проверка на заблокированные/требующие подтверждения функции.
      const risk = analyzeFormulaRisk(fixedBody);
      if (risk.level === "blocked") {
        return {
          ok: false,
          formulas: [],
          error: {
            code: "FORMULA_BLOCKED",
            message: risk.description,
            formula: raw,
          },
        };
      }
      if (risk.level === "confirm") {
        // В автономном режиме (Фаза 1) — отказ; UI-подтверждение в Фазе 3.
        return {
          ok: false,
          formulas: [],
          error: {
            code: "FORMULA_CONFIRM",
            message: risk.description,
            formula: raw,
          },
        };
      }

      outRow.push(`=${fixedBody}`);
    }
    result.push(outRow);
  }
  return { ok: true, formulas: result };
}

// ============================================================================
// W1. setValues — записать 2D-массив с валидацией размера и нормализацией
// (docs/03-TOOLS-SPEC.md §1 W1, §2 W1)
// ============================================================================

export const setValuesTool = defineTool({
  name: "setValues",
  description: `Записывает 2D-массив значений в диапазон. Каждая внутренняя строка = одна строка листа.
Используй, когда нужно записать конкретные значения (текст/числа/даты), НЕ формулы.
Адрес вида "A1", "A1:B10" или "Лист1!A1:B5". Строки разной длины выравниваются автоматически.
Лимит: ${SAFE_CELL_LIMIT} ячеек за вызов. Перед записью создаётся undo-снапшот (кнопка «Отменить всё» вернёт как было).
Строки, начинающиеся с = + - @, экранируются (чтобы не стать случайной формулой — для формул используй setFormula).`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Адрес левой верхней ячейки или диапазона: "A1" или "A1:D10" или "Лист1!A1:B5"',
      },
      values: {
        type: "array",
        items: { type: "array" },
        description: '2D массив: каждая внутренняя строка = строка листа. [[1,2,3],["а","б","в"]]',
      },
    },
    required: ["address", "values"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: (args: Record<string, unknown>) => {
    const values = args.values as unknown[][];
    if (!Array.isArray(values)) return 0;
    return values.length * (Array.isArray(values[0]) ? values[0].length : 1);
  },

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    const rawValues = args.values as unknown[][];
    if (!Array.isArray(rawValues) || rawValues.length === 0) {
      return {
        ok: false,
        summary: "Пустой массив значений",
        error: {
          code: "EMPTY_VALUES",
          message: "values должен быть непустым 2D-массивом",
          retryable: false,
        },
      };
    }

    const values = normalizeRows(rawValues);
    const rows = values.length;
    const cols = values[0]?.length ?? 0;
    if (cols === 0) {
      return {
        ok: false,
        summary: "Пустые строки значений",
        error: {
          code: "EMPTY_ROWS",
          message: "каждая строка values должна содержать хотя бы одно значение",
          retryable: false,
        },
      };
    }

    const cellCount = rows * cols;
    try {
      assertCellLimit(cellCount, "setValues");
    } catch (e) {
      return {
        ok: false,
        summary: `Превышен лимит ячеек: ${cellCount} > ${SAFE_CELL_LIMIT}`,
        error: {
          code: "RANGE_TOO_LARGE",
          message: e instanceof Error ? e.message : String(e),
          retryable: false,
        },
      };
    }

    const safeValues = escapeValues(values);

    // Undo-снапшот ДО записи (P4). Через withPerformanceGuard для одного sync.
    await undoManager.createBackup(address, "setValues", {
      description: `Записаны значения в ${address} (${cellCount} ячеек)`,
    });

    try {
      await withPerformanceGuard(async (context) => {
        const startCell = getRangeSafe(context, address).getCell(0, 0);
        const targetRange = startCell.getResizedRange(rows - 1, cols - 1);
        targetRange.values = safeValues;
        await context.sync();
      });
    } catch (e) {
      return {
        ok: false,
        summary: `Ошибка записи в ${address}`,
        error: {
          code: "WRITE_FAILED",
          message: e instanceof Error ? e.message : String(e),
          retryable: true,
        },
      };
    }

    return {
      ok: true,
      summary: `Записано ${rows}×${cols} (${cellCount} ячеек) в ${address}`,
      data: { address, rows, cols, cellCount },
      cellsAffected: cellCount,
    };
  },
});

toolRegistry.registerDefinition(setValuesTool);

// ============================================================================
// W2. setFormula — записать формулы с ОБЯЗАТЕЛЬНОЙ проверкой FormulaGuardian
// (docs/03-TOOLS-SPEC.md §1 W2, §2 W2, §4)
// ============================================================================

export const setFormulaTool = defineTool({
  name: "setFormula",
  description: `Записывает формулы в диапазон. Каждый элемент — строка-формула вида "=SUM(A1:A10)" или "SUM(A1:A10)".
Используй для ВЫЧИСЛЕНИЙ (сумма, среднее, ВПР, ЕСЛИ и т.д.), а не для текста.
КАЖДАЯ формула проверяется: сбалансированность скобок, авто-исправление "*" между смежными ссылками (B8B6 → B8*B6), блокировка опасных функций (WEBSERVICE, EXEC — эксфильтрация данных).
Перед записью создаётся undo-снапшот. Лимит: ${SAFE_CELL_LIMIT} ячеек.
Поддерживаются русские имена функций (СУММ, ЕСЛИ, ВПР).`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Адрес левой верхней ячейки или диапазона: "C5" или "C5:C10"',
      },
      formulas: {
        type: "array",
        items: { type: "array" },
        description: '2D массив формул: [["=SUM(A1:A3)"],["=B1*2"]]. С или без ведущего "=".',
      },
    },
    required: ["address", "formulas"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: (args: Record<string, unknown>) => {
    const formulas = args.formulas as unknown[][];
    if (!Array.isArray(formulas)) return 0;
    return formulas.length * (Array.isArray(formulas[0]) ? formulas[0].length : 1);
  },

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    const rawFormulas = args.formulas as unknown[][];
    if (!Array.isArray(rawFormulas) || rawFormulas.length === 0) {
      return {
        ok: false,
        summary: "Пустой массив формул",
        error: {
          code: "EMPTY_FORMULAS",
          message: "formulas должен быть непустым 2D-массивом",
          retryable: false,
        },
      };
    }

    // Нормализуем в 2D строковый массив.
    const normalized: string[][] = rawFormulas.map((row) =>
      (Array.isArray(row) ? row : [row]).map((cell) => {
        if (cell === null || cell === undefined) return "";
        return typeof cell === "string" ? cell : String(cell);
      }),
    );
    const rows = normalized.length;
    const cols = normalized[0]?.length ?? 0;
    const cellCount = rows * cols;
    if (cols === 0) {
      return {
        ok: false,
        summary: "Пустые строки формул",
        error: {
          code: "EMPTY_ROWS",
          message: "formulas должен содержать хотя бы одну колонку",
          retryable: false,
        },
      };
    }

    try {
      assertCellLimit(cellCount, "setFormula");
    } catch (e) {
      return {
        ok: false,
        summary: `Превышен лимит: ${cellCount} > ${SAFE_CELL_LIMIT}`,
        error: {
          code: "RANGE_TOO_LARGE",
          message: e instanceof Error ? e.message : String(e),
          retryable: false,
        },
      };
    }

    // Валидация и подготовка формул (P4 — нельзя писать невалидное).
    const prepared = prepareFormulas(normalized);
    if (!prepared.ok) {
      return {
        ok: false,
        summary: `Формула отклонена: ${prepared.error?.code}`,
        error: {
          code: prepared.error?.code ?? "FORMULA_ERROR",
          message: prepared.error?.message ?? "Ошибка формулы",
          retryable: false,
        },
      };
    }

    // Undo ДО записи.
    await undoManager.createBackup(address, "setFormula", {
      description: `Записаны формулы в ${address} (${cellCount} ячеек)`,
    });

    try {
      await withPerformanceGuard(async (context) => {
        const startCell = getRangeSafe(context, address).getCell(0, 0);
        const targetRange = startCell.getResizedRange(rows - 1, cols - 1);
        targetRange.formulas = prepared.formulas;
        await context.sync();
      });
    } catch (e) {
      return {
        ok: false,
        summary: `Ошибка записи формул в ${address}`,
        error: {
          code: "WRITE_FAILED",
          message: e instanceof Error ? e.message : String(e),
          retryable: true,
        },
      };
    }

    return {
      ok: true,
      summary: `Записано ${cellCount} формул(а/ы) в ${address}`,
      data: { address, rows, cols, cellCount },
      cellsAffected: cellCount,
    };
  },
});

toolRegistry.registerDefinition(setFormulaTool);

// ============================================================================
// W3. fillRange — заполнить по шаблону: прогрессия, копирование, константа
// (docs/03-TOOLS-SPEC.md §1 W3)
// ============================================================================

export type FillType = "progression" | "copy" | "value";

/**
 * Генерирует значения для заполнения диапазона по шаблону.
 * @param rows, cols — размеры целевого диапазона.
 * @param fillType   — тип заполнения.
 * @param startValue — стартовое значение (для progression/value).
 * @param step       — шаг прогрессии (по умолчанию 1).
 * @param fillValue  — значение/массив для копирования (для copy/value).
 */
export function generateFillValues(
  rows: number,
  cols: number,
  fillType: FillType,
  options: {
    startValue?: number;
    step?: number;
    fillValue?: unknown;
  } = {},
): unknown[][] {
  const result: unknown[][] = [];
  const start = options.startValue ?? 1;
  const step = options.step ?? 1;

  for (let r = 0; r < rows; r++) {
    const row: unknown[] = [];
    for (let c = 0; c < cols; c++) {
      const index = r * cols + c;
      if (fillType === "progression") {
        row.push(start + index * step);
      } else if (fillType === "copy") {
        // Копируем одно значение во все ячейки.
        row.push(options.fillValue ?? "");
      } else {
        // value — то же, что copy (одно значение во весь диапазон).
        row.push(options.fillValue ?? start);
      }
    }
    result.push(row);
  }
  return result;
}

export const fillRangeTool = defineTool({
  name: "fillRange",
  description: `Заполняет диапазон по шаблону: арифметическая прогрессия (1,2,3...), копирование одного значения во все ячейки, или константа.
Используй для "заполни колонку числами от 1 до 100", "поставь 0 во все ячейки", "номера строк 1,2,3...".
Перед заполнением создаётся undo-снапшот.`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Адрес диапазона: "A1:A100" или "A1:D10"',
      },
      fillType: {
        type: "string",
        enum: ["progression", "copy", "value"],
        description: "progression — арифметическая прогрессия; copy/value — одно значение во все ячейки.",
      },
      startValue: {
        type: "number",
        description: "Стартовое значение для progression (по умолчанию 1).",
      },
      step: {
        type: "number",
        description: "Шаг прогрессии (по умолчанию 1).",
      },
      fillValue: {
        description: 'Значение для заполнения (copy/value): число, строка или boolean.',
      },
    },
    required: ["address", "fillType"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: (args: Record<string, unknown>) => {
    // Без адресной арифметики — оценка через размер, при выполнении точная.
    const address = String(args.address ?? "");
    try {
      // chunkLarge / estimateRangeSize доступны, но чтобы не плодить зависимость,
      // используем минимальную оценку.
      return address.includes(":") ? 100 : 1;
    } catch {
      return 1;
    }
  },

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    const fillType = String(args.fillType ?? "") as FillType;
    if (!["progression", "copy", "value"].includes(fillType)) {
      return {
        ok: false,
        summary: `Неизвестный тип заполнения: ${fillType}`,
        error: {
          code: "INVALID_FILL_TYPE",
          message: "fillType должен быть progression, copy или value",
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("rowCount, columnCount, address");
      await context.sync();

      const rows = range.rowCount;
      const cols = range.columnCount;
      const cellCount = rows * cols;
      try {
        assertCellLimit(cellCount, "fillRange");
      } catch (e) {
        return {
          ok: false,
          summary: `Превышен лимит: ${cellCount} > ${SAFE_CELL_LIMIT}`,
          error: {
            code: "RANGE_TOO_LARGE",
            message: e instanceof Error ? e.message : String(e),
            retryable: false,
          },
        };
      }

      const values = generateFillValues(rows, cols, fillType, {
        startValue: typeof args.startValue === "number" ? args.startValue : undefined,
        step: typeof args.step === "number" ? args.step : undefined,
        fillValue: args.fillValue,
      });

      // Undo ДО записи.
      await undoManager.createBackup(address, "fillRange", {
        description: `Заполнен диапазон ${address} (${cellCount} ячеек, ${fillType})`,
      });

      const safeValues = escapeValues(values);
      range.values = safeValues;
      await context.sync();

      return {
        ok: true,
        summary: `Заполнено ${cellCount} ячеек в ${range.address} (${fillType})`,
        data: { address: range.address, rows, cols, cellCount, fillType },
        cellsAffected: cellCount,
      };
    });
  },
});

toolRegistry.registerDefinition(fillRangeTool);

// ============================================================================
// W4. appendRows — дописать строки вниз под usedRange
// (docs/03-TOOLS-SPEC.md §1 W4)
// ============================================================================

export const appendRowsTool = defineTool({
  name: "appendRows",
  description: `Дописывает строки данных вниз ПОД используемым диапазоном активного листа (или указанного листа).
Используй, когда нужно добавить новые строки в конец таблицы, не перезаписывая существующие данные.
Первая колонка используется для определения конца (ищет последнюю строку с данными).
Перед записью создаётся undo-снапшот.`,
  parameters: {
    type: "object",
    properties: {
      sheetName: {
        type: "string",
        description: "Имя листа (опционально). По умолчанию — активный лист.",
      },
      values: {
        type: "array",
        items: { type: "array" },
        description: "2D массив строк для добавления: [[1,'Иван'],[2,'Мария']]",
      },
    },
    required: ["values"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: (args: Record<string, unknown>) => {
    const values = args.values as unknown[][];
    if (!Array.isArray(values)) return 0;
    return values.length * (Array.isArray(values[0]) ? values[0].length : 1);
  },

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const rawValues = args.values as unknown[][];
    if (!Array.isArray(rawValues) || rawValues.length === 0) {
      return {
        ok: false,
        summary: "Пустой массив значений",
        error: {
          code: "EMPTY_VALUES",
          message: "values должен быть непустым 2D-массивом",
          retryable: false,
        },
      };
    }

    const values = normalizeRows(rawValues);
    const rows = values.length;
    const cols = values[0]?.length ?? 0;
    if (cols === 0) {
      return {
        ok: false,
        summary: "Пустые строки значений",
        error: {
          code: "EMPTY_ROWS",
          message: "каждая строка должна содержать хотя бы одно значение",
          retryable: false,
        },
      };
    }

    const cellCount = rows * cols;
    try {
      assertCellLimit(cellCount, "appendRows");
    } catch (e) {
      return {
        ok: false,
        summary: `Превышен лимит: ${cellCount} > ${SAFE_CELL_LIMIT}`,
        error: {
          code: "RANGE_TOO_LARGE",
          message: e instanceof Error ? e.message : String(e),
          retryable: false,
        },
      };
    }

    const sheetName = typeof args.sheetName === "string" ? args.sheetName : undefined;

    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("rowCount, columnCount, address, isNullObject");
      sheet.load("name");
      await context.sync();

      // Стартовая строка для дописывания: под последней строкой usedRange.
      // Если лист пустой — с первой строки.
      let startRow = 1;
      if (!usedRange.isNullObject && usedRange.rowCount > 0) {
        // usedRange.address вида "Лист!A1:D10" — берём endRow из адреса.
        const addrMatch = (usedRange.address || "").match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
        if (addrMatch) {
          // Если диапазон — endRow в группе 4, иначе — группа 2 (одиночная ячейка).
          startRow = addrMatch[4]
            ? parseInt(addrMatch[4], 10) + 1
            : parseInt(addrMatch[2], 10) + 1;
        } else {
          startRow = usedRange.rowCount + 1;
        }
      }

      // Целевой диапазон: с cols колонок начиная с A (или с колонки usedRange).
      const startCol = !usedRange.isNullObject && usedRange.columnCount > 0 ? "A" : "A";
      const targetAddress = `${sheet.name}!${startCol}${startRow}:${startCol}${startRow + rows - 1}`;

      const safeValues = escapeValues(values);

      // Undo ДО записи.
      await undoManager.createBackup(targetAddress, "appendRows", {
        sheetName: sheet.name,
        description: `Дописано ${rows} строк в «${sheet.name}» с ${startCol}${startRow}`,
      });

      const targetRange = sheet.getRange(`${startCol}${startRow}:${startCol}${startRow + rows - 1}`);
      // Расширяем до cols колонок.
      const fullTarget = targetRange.getResizedRange(0, cols - 1);
      fullTarget.values = safeValues;
      await context.sync();

      return {
        ok: true,
        summary: `Дописано ${rows} строк в «${sheet.name}» начиная с ${startCol}${startRow} (${cellCount} ячеек)`,
        data: {
          sheetName: sheet.name,
          startCell: `${startCol}${startRow}`,
          rows,
          cols,
          cellCount,
        },
        cellsAffected: cellCount,
      };
    });
  },
});

toolRegistry.registerDefinition(appendRowsTool);

// ============================================================================
// W5. clearRange — очистить значения/формулы/форматы
// (docs/03-TOOLS-SPEC.md §1 W5)
// ============================================================================

export type ClearWhat = "all" | "values" | "formats";

export const clearRangeTool = defineTool({
  name: "clearRange",
  description: `Очищает диапазон: значения и форматы (all), только значения/формулы (values), или только форматы (formats).
Используй для "очисти колонку B", "удали данные из A1:D10".
Перед очисткой создаётся undo-снапшот (кнопка «Отменить всё» вернёт данные).`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Адрес диапазона: "A1:D10" или "Лист1!B:B"',
      },
      clearWhat: {
        type: "string",
        enum: ["all", "values", "formats"],
        description: "all — values+formats; values — только значения/формулы; formats — только форматы. По умолчанию all.",
      },
    },
    required: ["address"],
  },
  // clearRange — деструктивная операция (удаляет данные), поэтому dangerous.
  // Legacy-версия требовала подтверждения (requiresConfirmation=true).
  riskLevel: "dangerous",
  requiresUndo: true,
  estimateCells: () => 100,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    const clearWhat = (String(args.clearWhat ?? "all") as ClearWhat);

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("address, rowCount, columnCount");
      await context.sync();

      const cellCount = range.rowCount * range.columnCount;

      // Undo ДО очистки.
      await undoManager.createBackup(address, "clearRange", {
        description: `Очищен диапазон ${range.address} (${cellCount} ячеек, ${clearWhat})`,
      });

      // Office.ClearRangesOptions нет в строгом API — используем enum-строку.
      // "all" → range.clear(); "values" → clear("Contents"); "formats" → clear("Formats").
      if (clearWhat === "all") {
        range.clear();
      } else if (clearWhat === "values") {
        // Excel.ClearApplyTo.contents — очистить содержимое (значения + формулы).
        range.clear(Excel.ClearApplyTo.contents);
      } else if (clearWhat === "formats") {
        range.clear(Excel.ClearApplyTo.formats);
      } else {
        return {
          ok: false,
          summary: `Неизвестный режим очистки: ${clearWhat}`,
          error: {
            code: "INVALID_CLEAR_WHAT",
            message: "clearWhat должен быть all, values или formats",
            retryable: false,
          },
        };
      }
      await context.sync();

      return {
        ok: true,
        summary: `Очищено ${cellCount} ячеек в ${range.address} (${clearWhat})`,
        data: { address: range.address, cellCount, clearWhat },
        cellsAffected: cellCount,
      };
    });
  },
});

toolRegistry.registerDefinition(clearRangeTool);
