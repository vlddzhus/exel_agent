/**
 * auto-designer.ts — Автоматический дизайнер таблиц.
 *
 * Итерация «Профессиональное форматирование». Главный инструмент для запросов
 * «сделай красиво», «оформи таблицу», «приведи в порядок».
 *
 * Логика (см. knowledge/sections/excel-templates.md — палитры):
 *   1. Читает данные диапазона (values + numberFormats).
 *   2. Классифицирует каждую колонку: text/number/currency/percent/date/boolean.
 *   3. По доминирующим типам определяет профиль таблицы:
 *      financial | inventory | schedule | generic.
 *   4. Применяет палитру (header fill/font, чередующиеся строки, границы,
 *      числовые форматы по типам колонок, autoFit).
 *
 * Безопасность: новый инструмент, существующие не трогаются. Регистрация через
 * defineTool + toolRegistry.registerDefinition. requiresUndo=true (создаёт backup).
 */
import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";
import { undoManager } from "./backup";

// ============================================================================
// Палитры дизайна (синхронизированы с knowledge/sections/excel-templates.md)
// ============================================================================

export type DesignIntent =
  | "professional"
  | "financial"
  | "dashboard"
  | "minimal"
  | "auto";

interface Palette {
  /** Заливка шапки. */
  headerFill: string;
  /** Цвет текста шапки. */
  headerFontColor: string;
  /** Цвет границ. */
  borderColor: string;
  /** Чередующиеся строки (fill чётных строк). */
  bandFill: string;
  /** Заливка "позитив" (итог/сумма). */
  totalFill: string;
  /** Цвет текста для итога. */
  totalFontColor: string;
  /** Шрифт. */
  fontName: string;
  /** Размер шрифта тела. */
  bodyFontSize: number;
  /** Размер шрифта шапки. */
  headerFontSize: number;
}

const PALETTES: Record<Exclude<DesignIntent, "auto">, Palette> = {
  // Corporate Blue — универсальный профессиональный стиль
  professional: {
    headerFill: "#2B579A",
    headerFontColor: "#FFFFFF",
    borderColor: "#D6E4F0",
    bandFill: "#F2F7FB",
    totalFill: "#1F3864",
    totalFontColor: "#FFFFFF",
    fontName: "Segoe UI",
    bodyFontSize: 10,
    headerFontSize: 11,
  },
  // Financial Report — строгий, для денег/процентов
  financial: {
    headerFill: "#1A1A1A",
    headerFontColor: "#FFFFFF",
    borderColor: "#808080",
    bandFill: "#F5F5F5",
    totalFill: "#000000",
    totalFontColor: "#FFFFFF",
    fontName: "Calibri",
    bodyFontSize: 10,
    headerFontSize: 11,
  },
  // Dashboard — яркий, для KPI/дашбордов
  dashboard: {
    headerFill: "#27AE60",
    headerFontColor: "#FFFFFF",
    borderColor: "#BDC3C7",
    bandFill: "#EAF7EE",
    totalFill: "#1E8449",
    totalFontColor: "#FFFFFF",
    fontName: "Segoe UI",
    bodyFontSize: 10,
    headerFontSize: 11,
  },
  // Minimal — минималистичный, без чередования
  minimal: {
    headerFill: "#F2F2F2",
    headerFontColor: "#1A1A1A",
    borderColor: "#D0D0D0",
    bandFill: "#FFFFFF",
    totalFill: "#E0E0E0",
    totalFontColor: "#1A1A1A",
    fontName: "Calibri",
    bodyFontSize: 10,
    headerFontSize: 10,
  },
};

// ============================================================================
// Определение типа колонки
// ============================================================================

type ColumnType =
  | "text"
  | "number"
  | "currency"
  | "percent"
  | "date"
  | "boolean"
  | "id";

/**
 * Классифицирует колонку по её значениям.
 * Берёт до 20 непустых значений, определяет доминирующий тип.
 */
function detectColumnType(values: unknown[]): ColumnType {
  const sample = values.filter((v) => v !== null && v !== undefined && v !== "").slice(0, 20);
  if (sample.length === 0) return "text";

  let numCount = 0;
  let dateCount = 0;
  let boolCount = 0;
  let currencyHint = 0;
  let percentHint = 0;
  let idHint = 0;

  for (const v of sample) {
    if (v instanceof Date) {
      dateCount++;
      continue;
    }
    if (typeof v === "number") {
      numCount++;
      // Эвристики: крупные числа → деньги, <1 и >0 → процент (если много таких)
      if (Math.abs(v) >= 1000) currencyHint++;
      if (v > 0 && v < 1) percentHint++;
      continue;
    }
    if (typeof v === "boolean") {
      boolCount++;
      continue;
    }
    const s = String(v).trim();
    // Дата-паттерны (ISO, DD.MM.YYYY, ...)
    if (/^\d{4}-\d{2}-\d{2}/.test(s) || /^\d{1,2}[./]\d{1,2}[./]\d{2,4}/.test(s)) {
      dateCount++;
      continue;
    }
    // Булевы
    if (/^(true|false|да|нет|yes|no|true|ложь|истина)$/i.test(s)) {
      boolCount++;
      continue;
    }
    // Числовая строка
    if (/^-?\d[\d\s.,]*$/.test(s)) {
      numCount++;
      const num = Number(s.replace(/\s/g, "").replace(",", "."));
      if (!isNaN(num) && Math.abs(num) >= 1000) currencyHint++;
      if (!isNaN(num) && num > 0 && num < 1) percentHint++;
      continue;
    }
    // ID-паттерны (буквенно-цифровые коды одинаковой длины)
    if (/^[A-Z0-9-]{4,}$/i.test(s) && s.length >= 4) {
      idHint++;
    }
  }

  const total = sample.length;
  if (boolCount / total >= 0.6) return "boolean";
  if (dateCount / total >= 0.6) return "date";
  if (numCount / total >= 0.6) {
    if (percentHint / Math.max(numCount, 1) >= 0.6) return "percent";
    if (currencyHint / Math.max(numCount, 1) >= 0.4) return "currency";
    return "number";
  }
  if (idHint / total >= 0.7) return "id";
  return "text";
}

/** Формат числа по умолчанию для типа колонки. */
const FORMAT_BY_TYPE: Record<ColumnType, string> = {
  text: "General",
  number: "#,##0.00",
  currency: '#,##0.00 ₽;[Red]-#,##0.00 ₽',
  percent: "0.0%",
  date: "DD.MM.YYYY",
  boolean: "General",
  id: "@",
};

/**
 * Определяет профиль таблицы по набору типов колонок.
 * Используется когда intent="auto".
 */
function detectProfile(types: ColumnType[]): Exclude<DesignIntent, "auto"> {
  const counts: Record<string, number> = {};
  for (const t of types) counts[t] = (counts[t] ?? 0) + 1;

  const hasCurrency = (counts.currency ?? 0) > 0;
  const hasPercent = (counts.percent ?? 0) > 0;
  const numericShare =
    ((counts.currency ?? 0) + (counts.percent ?? 0) + (counts.number ?? 0)) /
    Math.max(types.length, 1);

  // Финансовый профиль: есть деньги/проценты и много чисел
  if ((hasCurrency || hasPercent) && numericShare >= 0.4) return "financial";

  // Профиль инвентаризации: ID + числа
  if ((counts.id ?? 0) >= 1 && numericShare >= 0.3) return "professional";

  // Дашборд: много чисел + булевы
  if (numericShare >= 0.5 || (counts.boolean ?? 0) >= 1) return "dashboard";

  // По умолчанию — professional
  return "professional";
}

// ============================================================================
// Инструмент applyAutoDesign
// ============================================================================

export const applyAutoDesignTool = defineTool({
  name: "applyAutoDesign",
  description: `Автоматически оформляет диапазон как профессиональную таблицу.
Определяет тип данных в каждой колонке (числа/деньги/проценты/даты/текст) и применяет соответствующий дизайн: шапка с заливкой, чередующиеся строки, границы, правильные числовые форматы, авто-ширина.
Intent (стиль дизайна):
  - "auto" (по умолчанию) — автоопределение по данным
  - "professional" — корпоративный синий (универсальный)
  - "financial" — строгий чёрный для финансовых отчётов
  - "dashboard" — яркий зелёный для KPI/дашбордов
  - "minimal" — минимализм без чередования
Используй ДЛЯ ЗАПРОСОВ "сделай красиво", "оформи таблицу", "приведи в порядок", "профессиональный вид".`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Диапазон таблицы С заголовком: "A1:F100"',
      },
      intent: {
        type: "string",
        enum: ["auto", "professional", "financial", "dashboard", "minimal"],
        description: "Стиль дизайна (по умолчанию auto)",
      },
      hasHeader: {
        type: "boolean",
        description: "Первая строка — заголовки (по умолчанию true)",
      },
    },
    required: ["address"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: (args: Record<string, unknown>) => {
    const a = String(args.address ?? "");
    return a.includes(":") ? 200 : 1;
  },

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: {
          code: "MISSING_ADDRESS",
          message: "address обязателен",
          retryable: false,
        },
      };
    }
    const intentArg = String(args.intent ?? "auto") as DesignIntent;
    const hasHeader = args.hasHeader !== false;

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("rowCount, columnCount, address, values, numberFormat");
      await undoManager.createBackup(address, "applyAutoDesign", {
        description: `AutoDesign ${address} (${intentArg})`,
      });
      await context.sync();

      const rowCount = range.rowCount;
      const colCount = range.columnCount;
      const values = (range as any).values as unknown[][];
      const cellCount = rowCount * colCount;

      if (rowCount < 1 || colCount < 1) {
        return {
          ok: true,
          summary: `Диапазон ${range.address} пуст — нечего оформлять`,
          data: { address: range.address, skipped: true },
        };
      }

      // ── Шаг 1: определяем тип каждой колонки ──
      const headerRow = hasHeader ? 1 : 0;
      const dataStartRow = headerRow;
      const columnTypes: ColumnType[] = [];
      for (let c = 0; c < colCount; c++) {
        const colValues: unknown[] = [];
        for (let r = dataStartRow; r < rowCount; r++) {
          colValues.push(values[r]?.[c]);
        }
        columnTypes.push(detectColumnType(colValues));
      }

      // ── Шаг 2: выбираем палитру ──
      const intent: Exclude<DesignIntent, "auto"> =
        intentArg === "auto" ? detectProfile(columnTypes) : intentArg;
      const palette = PALETTES[intent];

      // ── Шаг 3: применяем дизайн ──
      // 3a. Шапка
      if (hasHeader && rowCount >= 1) {
        const headerRange = range.getCell(0, 0).getResizedRange(0, colCount - 1);
        const hf = headerRange.format;
        hf.fill.color = palette.headerFill;
        hf.font.color = palette.headerFontColor;
        hf.font.bold = true;
        hf.font.size = palette.headerFontSize;
        hf.font.name = palette.fontName;
        hf.horizontalAlignment = "Center";
        hf.verticalAlignment = "Center";
        // Толстая нижняя граница шапки
        const bottom = hf.borders.getItem("EdgeBottom" as any);
        bottom.style = "Continuous" as any;
        bottom.color = palette.headerFill;
      }

      // 3b. Тело таблицы: шрифт + чередование + границы
      if (rowCount > dataStartRow) {
        const bodyRange = range
          .getCell(dataStartRow, 0)
          .getResizedRange(rowCount - dataStartRow - 1, colCount - 1);
        const bf = bodyRange.format;
        bf.font.name = palette.fontName;
        bf.font.size = palette.bodyFontSize;
        bf.font.color = "#1A1A1A";
        bf.verticalAlignment = "Center" as any;

        // Чередующиеся строки (кроме minimal)
        if (intent !== "minimal") {
          for (let r = 0; r < rowCount - dataStartRow; r++) {
            if (r % 2 === 1) {
              const rowRange = bodyRange
                .getCell(r, 0)
                .getResizedRange(0, colCount - 1);
              rowRange.format.fill.color = palette.bandFill;
            }
          }
        }

        // Тонкие границы
        const edges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
        for (const edge of edges) {
          const b = bf.borders.getItem(edge as any);
          b.style = "Continuous" as any;
          b.color = palette.borderColor;
        }
      }

      // 3c. Числовые форматы по типам колонок
      const formatsByCol = columnTypes.map((t) => FORMAT_BY_TYPE[t]);
      for (let c = 0; c < colCount; c++) {
        const fmt = formatsByCol[c];
        if (fmt !== "General" && fmt !== "@") {
          const colRange = range
            .getCell(dataStartRow, c)
            .getResizedRange(rowCount - dataStartRow - 1, 0);
          if (columnTypes[c] === "number") {
            colRange.numberFormat = [[fmt]];
            colRange.format.horizontalAlignment = "Right" as any;
          } else if (columnTypes[c] === "currency") {
            colRange.numberFormat = [[fmt]];
            colRange.format.horizontalAlignment = "Right" as any;
          } else if (columnTypes[c] === "percent") {
            colRange.numberFormat = [[fmt]];
            colRange.format.horizontalAlignment = "Right" as any;
          } else if (columnTypes[c] === "date") {
            colRange.numberFormat = [[fmt]];
            colRange.format.horizontalAlignment = "Center" as any;
          }
        }
      }

      // 3d. Внешние границы (рамка вокруг всей таблицы)
      const outerEdges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
      for (const edge of outerEdges) {
        const b = range.format.borders.getItem(edge as any);
        b.style = "Continuous" as any;
        b.color = palette.headerFill;
      }

      // 3e. Авто-ширина колонок
      try {
        range.format.autofitColumns();
        if (hasHeader) {
          // Чуть выше шапка
          const headerRange = range
            .getCell(0, 0)
            .getResizedRange(0, colCount - 1);
          headerRange.format.rowHeight = 22;
        }
      } catch {
        // autofit может быть недоступен в некоторых контекстах — не критично
      }

      await context.sync();

      return {
        ok: true,
        summary: `Применён авто-дизайн "${intent}" к ${range.address} (${cellCount} ячеек, ${colCount} колонок: ${columnTypes.join(", ")})`,
        data: {
          address: range.address,
          intent,
          detectedIntent: intentArg === "auto" ? intent : undefined,
          cellCount,
          rowCount,
          colCount,
          columnTypes,
          palette: {
            headerFill: palette.headerFill,
            bandFill: palette.bandFill,
            fontName: palette.fontName,
          },
        },
        cellsAffected: cellCount,
      };
    });
  },
});

toolRegistry.registerDefinition(applyAutoDesignTool);
