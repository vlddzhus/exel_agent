/**
 * transform.ts — Transform-инструменты агента (категория T в docs/03-TOOLS-SPEC.md §1).
 *
 * Неделя 4 Фазы 1.
 *
 * Инструменты:
 *   T1 sortData          — сортировка по 1-3 колонкам.
 *   T2 filterData        — фильтр по значениям/условиям (скрытие, не удаление).
 *   T3 removeDuplicates  — удаление дублей по колонкам с отчётом.
 *   T4 splitTextToColumns — разбить по разделителю (ФИО, адреса).
 *   T5 normalizeText     — trim, регистр, удалить лишние пробелы.
 *   T6 lookup            — VLOOKUP/XLOOKUP как операция.
 */
import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";
import { undoManager } from "./backup";

// ============================================================================
// T1. sortData — сортировка по 1-3 колонкам
// (docs/03-TOOLS-SPEC.md §1 T1)
// ============================================================================

export const sortDataTool = defineTool({
  name: "sortData",
  description: `Сортирует диапазон по одной или нескольким колонкам.
Параметры:
  - address: адрес диапазона (обязательно)
  - sortColumns: массив { column: число (0-based), order: "asc" | "desc" }
  - hasHeaders: true (по умолчанию) — первая строка не участвует в сортировке
Используй для "отсортируй по сумме", "сортировка по убыванию", "упорядочь по дате сначала новые".`,
  parameters: {
    type: "object",
    properties: {
      address: { type: "string", description: 'Адрес диапазона: "A1:D100"' },
      sortColumns: {
        type: "array",
        items: {
          type: "object",
          properties: {
            column: { type: "number", description: "0-based индекс колонки" },
            order: {
              type: "string",
              enum: ["asc", "desc"],
              description: "Направление сортировки",
            },
          },
          required: ["column"],
        },
        description: "Колонки для сортировки (1-3)",
      },
      hasHeaders: {
        type: "boolean",
        description: "Первая строка — заголовки (не участвуют в сортировке)",
      },
    },
    required: ["address", "sortColumns"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "address обязателен",
        error: {
          code: "MISSING_ADDRESS",
          message: "address обязателен",
          retryable: false,
        },
      };
    }
    const sortColumns = (args.sortColumns ?? []) as {
      column: number;
      order?: string;
    }[];
    if (!sortColumns.length) {
      return {
        ok: false,
        summary: "sortColumns не может быть пустым",
        error: {
          code: "EMPTY_COLUMNS",
          message: "Укажите хотя бы одну колонку",
          retryable: false,
        },
      };
    }
    const hasHeaders = args.hasHeaders !== false;

    await undoManager.createBackup(address, "sortData", {
      description: `Сортировка ${address}`,
    });

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      const fields = sortColumns.map((c) => ({
        key: c.column,
        ascending: c.order !== "desc",
      }));
      range.load("address, rowCount, columnCount");
      await context.sync();
      (
        range.sort as unknown as {
          apply: (fields: unknown[], hasHeaders: boolean) => void;
        }
      ).apply(fields as unknown[], hasHeaders);
      await context.sync();
      return {
        ok: true,
        summary: `Диапазон ${address} отсортирован по ${fields.length} колонкe`,
        data: { address, sortColumns: fields, hasHeaders },
      };
    });
  },
});

toolRegistry.registerDefinition(sortDataTool);

// ============================================================================
// T2. filterData — фильтр по значениям/условиям
// (docs/03-TOOLS-SPEC.md §1 T2)
// ============================================================================

export const filterDataTool = defineTool({
  name: "filterData",
  description: `Включает или отключает автофильтр на диапазоне.
Параметры:
  - address: адрес диапазона (обязательно)
  - clear: true — отключить фильтр, false (по умолчанию) — включить
  - filters: массив колоночных фильтров (опционально)
    { column: число (0-based), values? ["знач1","знач2"] }
Используй для "отфильтруй", "покажи только оплачено", "убери фильтры".`,
  parameters: {
    type: "object",
    properties: {
      address: { type: "string", description: 'Адрес диапазона: "A1:D100"' },
      clear: { type: "boolean", description: "true — отключить фильтр" },
      filters: {
        type: "array",
        items: {
          type: "object",
          properties: {
            column: { type: "number", description: "0-based индекс колонки" },
            values: {
              type: "array",
              items: { type: "string" },
              description: "Значения для фильтрации",
            },
          },
          required: ["column"],
        },
        description: "Колоночные фильтры",
      },
    },
    required: ["address"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "address обязателен",
        error: {
          code: "MISSING_ADDRESS",
          message: "address обязателен",
          retryable: false,
        },
      };
    }
    const clear = args.clear === true;
    const filters = (args.filters ?? []) as {
      column: number;
      values?: string[];
    }[];

    if (!clear) {
      await undoManager.createBackup(address, "filterData", {
        description: `Фильтр ${address}`,
      });
    }

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      if (clear) {
        (sheet.autoFilter as any).clear();
        await context.sync();
        return {
          ok: true,
          summary: "Автофильтр отключён",
          data: { address, active: false },
        };
      }

      const range = getRangeSafe(context, address);
      if (filters.length > 0) {
        for (const f of filters) {
          if (f.values?.length) {
            (sheet.autoFilter as any).apply(range, f.column, {
              filterType: "Values",
              values: f.values,
            });
          }
        }
      } else {
        sheet.autoFilter.apply(range);
      }
      await context.sync();
      return {
        ok: true,
        summary: `Автофильтр включён на ${address}`,
        data: { address, active: true, filtersCount: filters.length },
      };
    });
  },
});

toolRegistry.registerDefinition(filterDataTool);

// ============================================================================
// T3. removeDuplicates — удаление дублей по колонкам с отчётом
// (docs/03-TOOLS-SPEC.md §1 T3, §2 T3)
// ============================================================================

export const removeDuplicatesTool = defineTool({
  name: "removeDuplicates",
  description: `Удаляет дубликаты строк в диапазоне по указанным колонкам.
Параметры:
  - address: адрес диапазона (обязательно)
  - columns: массив 0-based индексов колонок для сравнения
  - hasHeaders: true (по умолчанию) — первая строка — заголовок
Используй для "удали дубликаты", "убери повторы", "оставь уникальные".`,
  parameters: {
    type: "object",
    properties: {
      address: { type: "string", description: 'Адрес диапазона: "A1:D100"' },
      columns: {
        type: "array",
        items: { type: "number" },
        description: "0-based индексы колонок для сравнения",
      },
      hasHeaders: {
        type: "boolean",
        description: "Первая строка — заголовки (не участвует)",
      },
    },
    required: ["address", "columns"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "address обязателен",
        error: {
          code: "MISSING_ADDRESS",
          message: "address обязателен",
          retryable: false,
        },
      };
    }
    const columns = (args.columns ?? []) as number[];
    if (!columns.length) {
      return {
        ok: false,
        summary: "columns не может быть пустым",
        error: {
          code: "EMPTY_COLUMNS",
          message: "Укажите хотя бы одну колонку",
          retryable: false,
        },
      };
    }
    const hasHeaders = args.hasHeaders !== false;

    await undoManager.createBackup(address, "removeDuplicates", {
      description: `Удаление дублей ${address}`,
    });

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("address, rowCount, columnCount");
      await context.sync();

      const result = range.removeDuplicates(columns, hasHeaders);
      if (result) result.load("removed");
      await context.sync();
      const removed = result?.removed ?? 0;
      return {
        ok: true,
        summary: `Удалено ${removed} дубликатов из ${address}`,
        data: { address, removed, columns, hasHeaders },
        cellsAffected: removed,
      };
    });
  },
});

toolRegistry.registerDefinition(removeDuplicatesTool);

// ============================================================================
// T4. splitTextToColumns — разбить по разделителю (ФИО, адреса)
// (docs/03-TOOLS-SPEC.md §1 T4, §2 T4)
// ============================================================================

const DELIMITER_MAP: Record<string, string> = {
  auto: "",
  space: " ",
  comma: ",",
  semicolon: ";",
  dot: ".",
};

export const splitTextToColumnsTool = defineTool({
  name: "splitTextToColumns",
  description: `Разбивает текст в колонках по разделителю.
Параметры:
  - address: адрес диапазона с одной колонкой (обязательно)
  - delimiter: "space" | "comma" | "semicolon" | "dot" | "auto" (по умолчанию)
  - maxColumns: максимальное число колонок результата (опционально)
  - targetStartCell: левая верхняя ячейка для результата (опционально)
Используй для "разбей ФИО", "раздели адрес", "текст в колонки".`,
  parameters: {
    type: "object",
    properties: {
      address: { type: "string", description: 'Адрес: "A2:A100"' },
      delimiter: {
        type: "string",
        enum: ["auto", "space", "comma", "semicolon", "dot"],
        description: "Разделитель",
      },
      maxColumns: {
        type: "number",
        description: "Максимум колонок (опционально)",
      },
      targetStartCell: {
        type: "string",
        description: 'Начальная ячейка для результата: "B2" (опционально)',
      },
    },
    required: ["address"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "address обязателен",
        error: {
          code: "MISSING_ADDRESS",
          message: "address обязателен",
          retryable: false,
        },
      };
    }
    const delimiter = (args.delimiter as string) || "auto";
    const delimeterStr = DELIMITER_MAP[delimiter];
    if (!delimeterStr && delimiter !== "auto") {
      return {
        ok: false,
        summary: `Неизвестный разделитель: "${delimiter}"`,
        error: {
          code: "INVALID_DELIMITER",
          message: "Используйте space/comma/semicolon/dot/auto",
          retryable: false,
        },
      };
    }
    const maxColumns =
      typeof args.maxColumns === "number" ? args.maxColumns : undefined;
    const targetStartCell =
      typeof args.targetStartCell === "string"
        ? args.targetStartCell
        : undefined;

    if (targetStartCell) {
      await undoManager.createBackup(targetStartCell, "splitTextToColumns", {
        description: `Разбивка текста в ${address}`,
      });
    } else {
      await undoManager.createBackup(address, "splitTextToColumns", {
        description: `Разбивка текста в ${address}`,
      });
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("address, rowCount, columnCount");
      await context.sync();

      if (range.rowCount < 2) {
        return {
          ok: false,
          summary: `Диапазон ${address} содержит менее 2 строк`,
          error: {
            code: "TOO_FEW_ROWS",
            message: "Нужно минимум 2 строки (заголовок + данные)",
            retryable: false,
          },
        };
      }

      const textQualifier =
        (Excel as unknown as { TextQualifier: Record<string, number> })
          .TextQualifier?.textQualifierDoubleQuote ?? 1;

      const options: {
        delimiter: string;
        textQualifier: number;
        consecutiveDelimiter: boolean;
        numColumns?: number;
      } = {
        delimiter: delimeterStr,
        textQualifier,
        consecutiveDelimiter: delimiter === "auto" || delimiter === "space",
      };
      if (maxColumns !== undefined) options.numColumns = maxColumns;

      if (targetStartCell) {
        const target = getRangeSafe(context, targetStartCell);
        target.load("address");
        await context.sync();
        (
          range as unknown as { textToColumns: (opts: unknown) => void }
        ).textToColumns(options as unknown);
      } else {
        (
          range as unknown as { textToColumns: (opts: unknown) => void }
        ).textToColumns(options as unknown);
      }
      await context.sync();

      return {
        ok: true,
        summary: `Текст разбит по разделителю "${delimiter}" в ${address}`,
        data: { address, delimiter, maxColumns, targetStartCell },
      };
    });
  },
});

toolRegistry.registerDefinition(splitTextToColumnsTool);

// ============================================================================
// T5. normalizeText — trim, регистр, лишние пробелы
// (docs/03-TOOLS-SPEC.md §1 T5)
// ============================================================================

type NormalizeOp =
  "trim" | "uppercase" | "lowercase" | "propercase" | "cleanWhitespace";

/** @internal exported for testing */
export function normalizeValue(value: unknown, ops: NormalizeOp[]): unknown {
  if (typeof value !== "string") return value;
  let s = value;
  for (const op of ops) {
    switch (op) {
      case "trim":
        s = s.trim();
        break;
      case "uppercase":
        s = s.toUpperCase();
        break;
      case "lowercase":
        s = s.toLowerCase();
        break;
      case "propercase":
        s = s.replace(
          /\w\S*/g,
          (w) => w[0].toUpperCase() + w.slice(1).toLowerCase(),
        );
        break;
      case "cleanWhitespace":
        s = s.replace(/\s+/g, " ").trim();
        break;
    }
  }
  return s;
}

export const normalizeTextTool = defineTool({
  name: "normalizeText",
  description: `Нормализует текст в диапазоне: обрезка пробелов, регистр, удаление лишних пробелов.
Параметры:
  - address: адрес диапазона (обязательно)
  - operations: массив операций из: "trim", "uppercase", "lowercase", "propercase", "cleanWhitespace"
Используй для "сделай заглавными", "убери лишние пробелы", "приведи к нижнему регистру".`,
  parameters: {
    type: "object",
    properties: {
      address: { type: "string", description: 'Адрес: "B2:B100"' },
      operations: {
        type: "array",
        items: {
          type: "string",
          enum: [
            "trim",
            "uppercase",
            "lowercase",
            "propercase",
            "cleanWhitespace",
          ],
        },
        description: "Операции нормализации",
      },
    },
    required: ["address", "operations"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "address обязателен",
        error: {
          code: "MISSING_ADDRESS",
          message: "address обязателен",
          retryable: false,
        },
      };
    }
    const operations = (args.operations ?? []) as NormalizeOp[];
    if (!operations.length) {
      return {
        ok: false,
        summary: "operations не может быть пустым",
        error: {
          code: "EMPTY_OPS",
          message: "Укажите хотя бы одну операцию",
          retryable: false,
        },
      };
    }

    await undoManager.createBackup(address, "normalizeText", {
      description: `Нормализация текста ${address}`,
    });

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("values, rowCount, columnCount");
      await context.sync();

      const oldValues = range.values as unknown[][];
      const newValues = oldValues.map((row) =>
        row.map((cell) => normalizeValue(cell, operations)),
      );
      range.values = newValues as any[][];
      await context.sync();

      return {
        ok: true,
        summary: `Текст нормализован в ${address}`,
        data: {
          address,
          operations,
          cellCount: newValues.length * (newValues[0]?.length ?? 0),
        },
      };
    });
  },
});

toolRegistry.registerDefinition(normalizeTextTool);

// ============================================================================
// T6. lookup — VLOOKUP-подобная операция
// (docs/03-TOOLS-SPEC.md §1 T6)
// ============================================================================

export const lookupTool = defineTool({
  name: "lookup",
  description: `Выполняет поиск значения в таблице и записывает результат в указанную колонку.
Параметры:
  - lookupAddress: адрес диапазона, где искать (обязательно, с заголовками)
  - lookupColumn: 0-based индекс колонки для поиска (обязательно)
  - resultColumn: 0-based индекс колонки, откуда брать результат (обязательно)
  - lookupValue: искомое значение (обязательно)
  - writeTo: адрес ячейки, куда записать результат (обязательно)
  - exactMatch: true (по умолчанию) — точное совпадение
Используй для "найди цену товара", "подтяни ставку из таблицы".`,
  parameters: {
    type: "object",
    properties: {
      lookupAddress: {
        type: "string",
        description: 'Диапазон-таблица: "A2:C100"',
      },
      lookupColumn: {
        type: "number",
        description: "0-based индекс колонки для поиска",
      },
      resultColumn: {
        type: "number",
        description: "0-based индекс колонки результата",
      },
      lookupValue: { type: "string", description: "Искомое значение" },
      writeTo: { type: "string", description: 'Ячейка для результата: "D2"' },
      exactMatch: {
        type: "boolean",
        description: "Точное совпадение (по умолчанию true)",
      },
    },
    required: [
      "lookupAddress",
      "lookupColumn",
      "resultColumn",
      "lookupValue",
      "writeTo",
    ],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const lookupAddress = String(args.lookupAddress ?? "");
    if (!lookupAddress) {
      return {
        ok: false,
        summary: "lookupAddress обязателен",
        error: {
          code: "MISSING_ADDRESS",
          message: "lookupAddress обязателен",
          retryable: false,
        },
      };
    }
    const lookupColumn = args.lookupColumn as number | undefined;
    if (lookupColumn === undefined || lookupColumn < 0) {
      return {
        ok: false,
        summary: "lookupColumn обязателен и ≥ 0",
        error: {
          code: "INVALID_COLUMN",
          message: "lookupColumn",
          retryable: false,
        },
      };
    }
    const resultColumn = args.resultColumn as number | undefined;
    if (resultColumn === undefined || resultColumn < 0) {
      return {
        ok: false,
        summary: "resultColumn обязателен и ≥ 0",
        error: {
          code: "INVALID_COLUMN",
          message: "resultColumn",
          retryable: false,
        },
      };
    }
    const lookupValue = String(args.lookupValue ?? "");
    if (!lookupValue) {
      return {
        ok: false,
        summary: "lookupValue обязателен",
        error: {
          code: "MISSING_VALUE",
          message: "lookupValue",
          retryable: false,
        },
      };
    }
    const writeTo = String(args.writeTo ?? "");
    if (!writeTo) {
      return {
        ok: false,
        summary: "writeTo обязателен",
        error: { code: "MISSING_TARGET", message: "writeTo", retryable: false },
      };
    }
    // ponytail: exactMatch param accepted but always exact; fuzzy can be added when users request it
    // const exactMatch = args.exactMatch !== false;

    await undoManager.createBackup(writeTo, "lookup", {
      description: `Поиск значения в ${lookupAddress}`,
    });

    return Excel.run(async (context) => {
      const tableRange = getRangeSafe(context, lookupAddress);
      tableRange.load("values, rowCount, columnCount");
      await context.sync();

      const values = tableRange.values as unknown[][];
      let foundValue: unknown = null;
      for (let r = 1; r < values.length; r++) {
        const row = values[r];
        if (row && String(row[lookupColumn]) === lookupValue) {
          foundValue = row[resultColumn];
          break;
        }
      }

      if (foundValue === null) {
        return {
          ok: false,
          summary: `Значение "${lookupValue}" не найдено в ${lookupAddress}`,
          error: {
            code: "NOT_FOUND",
            message: "Значение не найдено",
            retryable: false,
          },
        };
      }

      const targetRange = getRangeSafe(context, writeTo);
      targetRange.values = [[foundValue]];
      await context.sync();

      return {
        ok: true,
        summary: `Найдено и записано значение в ${writeTo}`,
        data: { lookupValue, result: foundValue, writeTo },
        cellsAffected: 1,
      };
    });
  },
});

toolRegistry.registerDefinition(lookupTool);
