/**
 * format.ts — Format-инструменты агента (категория F в docs/03-TOOLS-SPEC.md §1).
 *
 * Неделя 3 Фазы 1.
 *
 * Инструменты:
 *   F1 applyCellFormat     — шрифт, заливка, границы, выравнивание.
 *   F2 applyNumberFormat   — числовой/денежный/процентный/дата-формат.
 *   F3 applyConditionalFormat — colorScale, dataBar, top10, highlightCell.
 *   F4 formatAsTable       — превратить диапазон в именованную таблицу.
 *   F5 autoFitColumns      — авто-ширина колонок.
 */
import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";
import { undoManager } from "./backup";

// ============================================================================
// F1. applyCellFormat
// (docs/03-TOOLS-SPEC.md §1 F1)
// ============================================================================

export interface CellFormatOptions {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontColor?: string;
  fillColor?: string;
  horizontalAlignment?: "left" | "center" | "right";
  verticalAlignment?: "top" | "center" | "bottom";
  wrapText?: boolean;
  borderTop?: { style: string; color: string };
  borderBottom?: { style: string; color: string };
  borderLeft?: { style: string; color: string };
  borderRight?: { style: string; color: string };
}

const BORDER_EDGES = [
  "EdgeTop",
  "EdgeBottom",
  "EdgeLeft",
  "EdgeRight",
] as const;

export const applyCellFormatTool = defineTool({
  name: "applyCellFormat",
  description: `Применяет форматирование к диапазону: шрифт (bold/italic/size/color), заливка (fillColor), выравнивание (horizontal/vertical/wrap), границы (borderTop/Bottom/Left/Right).
Используй для "сделай жирным", "залей жёлтым", "выровняй по центру".
Формат границ: style = "Continuous"/"Dash"/"Dot", color = CSS hex "#FF0000".`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Адрес: "A1:D10" или "Лист1!B:B"',
      },
      format: {
        type: "object",
        properties: {
          bold: { type: "boolean" },
          italic: { type: "boolean" },
          fontSize: { type: "number" },
          fontColor: { type: "string" },
          fillColor: { type: "string" },
          horizontalAlignment: {
            type: "string",
            enum: ["left", "center", "right"],
          },
          verticalAlignment: {
            type: "string",
            enum: ["top", "center", "bottom"],
          },
          wrapText: { type: "boolean" },
          borderTop: {
            type: "object",
            properties: {
              style: { type: "string" },
              color: { type: "string" },
            },
          },
          borderBottom: {
            type: "object",
            properties: {
              style: { type: "string" },
              color: { type: "string" },
            },
          },
          borderLeft: {
            type: "object",
            properties: {
              style: { type: "string" },
              color: { type: "string" },
            },
          },
          borderRight: {
            type: "object",
            properties: {
              style: { type: "string" },
              color: { type: "string" },
            },
          },
        },
      },
    },
    required: ["address", "format"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: (args: Record<string, unknown>) => {
    const a = String(args.address ?? "");
    return a.includes(":") ? 100 : 1;
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
    const fmt = (args.format ?? {}) as CellFormatOptions;

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("rowCount, columnCount, address");
      await context.sync();
      const cellCount = range.rowCount * range.columnCount;

      await undoManager.createBackup(address, "applyCellFormat", {
        description: `Форматирование ${range.address} (${cellCount} ячеек)`,
      });

      const rf = range.format;
      if (fmt.bold !== undefined) rf.font.bold = fmt.bold;
      if (fmt.italic !== undefined) rf.font.italic = fmt.italic;
      if (fmt.fontSize !== undefined) rf.font.size = fmt.fontSize;
      if (fmt.fontColor !== undefined) rf.font.color = fmt.fontColor;
      if (fmt.fillColor !== undefined) rf.fill.color = fmt.fillColor;
      const hMap: Record<string, string> = {
        left: "Left",
        center: "Center",
        right: "Right",
      };
      const vMap: Record<string, string> = {
        top: "Top",
        center: "Center",
        bottom: "Bottom",
      };
      if (fmt.horizontalAlignment)
        rf.horizontalAlignment = hMap[fmt.horizontalAlignment] as any;
      if (fmt.verticalAlignment)
        rf.verticalAlignment = vMap[fmt.verticalAlignment] as any;
      if (fmt.wrapText !== undefined) rf.wrapText = fmt.wrapText;

      for (const edge of BORDER_EDGES) {
        const b = fmt[
          `border${edge.replace("Edge", "")}` as keyof CellFormatOptions
        ] as { style: string; color: string } | undefined;
        if (b) {
          rf.borders.getItem(edge).style = b.style as any;
          rf.borders.getItem(edge).color = b.color;
        }
      }

      await context.sync();
      return {
        ok: true,
        summary: `Применено форматирование к ${range.address} (${cellCount} ячеек)`,
        data: { address: range.address, cellCount },
        cellsAffected: cellCount,
      };
    });
  },
});

toolRegistry.registerDefinition(applyCellFormatTool);

// ============================================================================
// F2. applyNumberFormat
// (docs/03-TOOLS-SPEC.md §1 F2)
// ============================================================================

const NUMBER_FORMAT_PRESETS: Record<string, string> = {
  number: "#,##0.00",
  integer: "#,##0",
  currency: "#,##0.00₽",
  percent: "0.00%",
  date: "DD.MM.YYYY",
  dateShort: "DD.MM.YY",
  dateLong: "DD MMMM YYYY",
  datetime: "DD.MM.YYYY HH:mm",
  time: "HH:mm:ss",
  scientific: "0.00E+00",
  text: "@",
};

export const applyNumberFormatTool = defineTool({
  name: "applyNumberFormat",
  description: `Применяет числовой формат к диапазону.
Предустановки: number, integer, currency, percent, date, datetime, time, scientific, text.
Или произвольный Excel-формат: "#,##0.00", "DD.MM.YYYY".`,
  parameters: {
    type: "object",
    properties: {
      address: { type: "string", description: 'Адрес: "A1:D10"' },
      format: {
        type: "string",
        description:
          'Предустановка или Excel-формат: "currency", "percent", "DD.MM.YYYY"',
      },
    },
    required: ["address", "format"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 100,

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
    const rawFormat = String(args.format ?? "");
    if (!rawFormat) {
      return {
        ok: false,
        summary: "Не указан формат",
        error: {
          code: "MISSING_FORMAT",
          message: "format обязателен",
          retryable: false,
        },
      };
    }
    const numberFormat = NUMBER_FORMAT_PRESETS[rawFormat] ?? rawFormat;

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("rowCount, columnCount, address");
      await context.sync();
      const cellCount = range.rowCount * range.columnCount;

      await undoManager.createBackup(address, "applyNumberFormat", {
        description: `Числовой формат ${range.address}: ${rawFormat}`,
      });

      range.numberFormat = [[numberFormat]];
      await context.sync();

      return {
        ok: true,
        summary: `Применён формат "${rawFormat}" к ${range.address} (${cellCount} ячеек)`,
        data: { address: range.address, cellCount, format: rawFormat },
        cellsAffected: cellCount,
      };
    });
  },
});

toolRegistry.registerDefinition(applyNumberFormatTool);

// ============================================================================
// F3. applyConditionalFormat
// (docs/03-TOOLS-SPEC.md §1 F3)
// ============================================================================

export type CfRuleType = "colorScale" | "dataBar" | "top10" | "highlightCell";

export interface CfRule {
  type: CfRuleType;
  minColor?: string;
  maxColor?: string;
  midColor?: string;
  fillColor?: string;
  showBarOnly?: boolean;
  rank?: number;
  bottom?: boolean;
  percent?: boolean;
  operator?: string;
  value1?: number | string;
  value2?: number | string;
  fontColor?: string;
}

export const applyConditionalFormatTool = defineTool({
  name: "applyConditionalFormat",
  description: `Применяет условное форматирование к диапазону.
Типы:
  - colorScale: градиент (minColor, midColor, maxColor)
  - dataBar: столбцы (fillColor, showBarOnly)
  - top10: топ N (rank, bottom, percent)
  - highlightCell: по условию (operator: greaterThan/lessThan/equalTo/between/containsText, value1, value2)
Используй для "выдели цветом больше 1000", "градиент от красного к зелёному", "топ-10".`,
  parameters: {
    type: "object",
    properties: {
      address: { type: "string", description: 'Адрес: "A1:D10"' },
      rules: {
        type: "array",
        items: {
          type: "object",
          properties: {
            type: {
              type: "string",
              enum: ["colorScale", "dataBar", "top10", "highlightCell"],
            },
            minColor: { type: "string" },
            maxColor: { type: "string" },
            midColor: { type: "string" },
            fillColor: { type: "string" },
            showBarOnly: { type: "boolean" },
            rank: { type: "number" },
            bottom: { type: "boolean" },
            percent: { type: "boolean" },
            operator: { type: "string" },
            value1: { type: ["number", "string"] },
            value2: { type: ["number", "string"] },
            fontColor: { type: "string" },
          },
          required: ["type"],
        },
      },
    },
    required: ["address", "rules"],
  },
  riskLevel: "moderate",
  requiresUndo: false,
  estimateCells: () => 100,

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
    const rules = args.rules as CfRule[];
    if (!Array.isArray(rules) || rules.length === 0) {
      return {
        ok: false,
        summary: "Не указаны правила",
        error: {
          code: "MISSING_RULES",
          message: "rules должен быть непустым массивом",
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);

      for (const rule of rules) {
        switch (rule.type) {
          case "colorScale": {
            const cf = range.conditionalFormats.add("ColorScale" as any);
            const cs = (cf as any).colorScale;
            cs.criteria = {
              minimum: {
                type: "LowestValue",
                color: rule.minColor ?? "#F8696B",
              },
              maximum: {
                type: "HighestValue",
                color: rule.maxColor ?? "#63BE7B",
              },
            };
            if (rule.midColor) {
              cs.criteria.midpoint = {
                type: "Percent",
                value: 50,
                color: rule.midColor,
              };
            }
            break;
          }
          case "dataBar": {
            const cf = range.conditionalFormats.add("DataBar" as any);
            (cf as any).bar.fill.color = rule.fillColor ?? "#5cb85c";
            if (rule.showBarOnly) (cf as any).bar.showBarOnly = true;
            break;
          }
          case "top10": {
            const cf = range.conditionalFormats.add("Top10" as any);
            const t10 = (cf as any).top10;
            t10.rank = rule.rank ?? 10;
            if (rule.bottom) t10.bottom = true;
            if (rule.percent) t10.percent = true;
            break;
          }
          case "highlightCell": {
            const cf = range.conditionalFormats.add("CellValue" as any);
            const cv = (cf as any).cellValue;
            const op = rule.operator ?? "greaterThan";
            if (op === "between") {
              cv.rule = {
                operator: "Between",
                formula1: String(rule.value1 ?? ""),
                formula2: String(rule.value2 ?? ""),
              };
            } else if (op === "containsText") {
              cv.rule = {
                operator: "ContainsText",
                formula1: String(rule.value1 ?? ""),
                formula2: "",
              };
            } else {
              const opMap: Record<string, string> = {
                greaterThan: "GreaterThan",
                lessThan: "LessThan",
                equalTo: "EqualTo",
              };
              cv.rule = {
                operator: opMap[op] ?? "GreaterThan",
                formula1: String(rule.value1 ?? "0"),
                formula2: "",
              };
            }
            if (rule.fillColor)
              (cf as any).cellValue.format.fill.color = rule.fillColor;
            if (rule.fontColor)
              (cf as any).cellValue.format.font.color = rule.fontColor;
            break;
          }
        }
      }
      await context.sync();

      return {
        ok: true,
        summary: `Применено ${rules.length} правил(о) условного форматирования к ${address}`,
        data: { address, rulesCount: rules.length },
      };
    });
  },
});

toolRegistry.registerDefinition(applyConditionalFormatTool);

// ============================================================================
// F4. formatAsTable — превратить диапазон в именованную таблицу
// (docs/03-TOOLS-SPEC.md §1 F4)
// ============================================================================

export const formatAsTableTool = defineTool({
  name: "formatAsTable",
  description: `Превращает диапазон в именованную таблицу Excel с авто-стилем.
Первая строка = заголовки (если hasHeaders=true).
Стили: "TableStyleLight1"-"Light21", "TableStyleMedium1"-"Medium28", "TableStyleDark1"-"Dark11".
Используй для "сделай таблицу", "отформатируй как таблицу".`,
  parameters: {
    type: "object",
    properties: {
      address: { type: "string", description: 'Адрес: "A1:D100"' },
      hasHeaders: {
        type: "boolean",
        description: "Первая строка — заголовки (по умолчанию true)",
      },
      tableName: { type: "string", description: "Имя таблицы (опционально)" },
      style: { type: "string", description: 'Стиль: "TableStyleLight1"' },
    },
    required: ["address"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 200,

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
    const hasHeaders = args.hasHeaders !== false;
    const tableName =
      typeof args.tableName === "string" ? args.tableName : undefined;
    const style = typeof args.style === "string" ? args.style : undefined;

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(address);
      const table = sheet.tables.add(range, hasHeaders);
      if (tableName) table.name = tableName;
      if (style) table.style = style;
      table.load("name, style");
      await context.sync();

      return {
        ok: true,
        summary: `Создана таблица "${table.name}"`,
        data: { name: table.name, style: table.style },
        cellsAffected: 0,
      };
    });
  },
});

toolRegistry.registerDefinition(formatAsTableTool);

// ============================================================================
// F5. autoFitColumns — авто-ширина колонок
// (docs/03-TOOLS-SPEC.md §1 F5)
// ============================================================================

export const autoFitColumnsTool = defineTool({
  name: "autoFitColumns",
  description: `Авто-ширина колонок по содержимому.
Можно указать диапазон или весь лист (без address).
Используй для "подгони ширину колонок", "авто-ширина".`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Диапазон: "A:D" или "A1:D10" (опционально)',
      },
    },
    required: [],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rawAddress =
        typeof args.address === "string" ? args.address : undefined;

      if (rawAddress) {
        const range = getRangeSafe(context, rawAddress);
        range.format.autofitColumns();
        await context.sync();
        return {
          ok: true,
          summary: `Авто-ширина колонок в ${rawAddress}`,
          data: { address: rawAddress },
        };
      }

      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject, address");
      await context.sync();
      if (!usedRange.isNullObject) {
        usedRange.format.autofitColumns();
        await context.sync();
        return {
          ok: true,
          summary: `Авто-ширина по всему листу "${sheet.name}"`,
          data: { sheetName: sheet.name },
        };
      }
      return {
        ok: true,
        summary: `Лист "${sheet.name}" пуст`,
        data: { sheetName: sheet.name },
      };
    });
  },
});

toolRegistry.registerDefinition(autoFitColumnsTool);

// ============================================================================
// F6. setColumnWidths — явная установка ширины колонок
// ============================================================================

export const setColumnWidthsTool = defineTool({
  name: "setColumnWidths",
  description: `Устанавливает ширину одной или нескольких колонок вручную.
Укажите address как "A:A" для одной колонки или "A:D" для нескольких (все получат одинаковую ширину).
Для разной ширины вызывайте повторно для каждой колонки.
Используйте когда autoFitColumns даёт слишком широкие колонки (напр. при длинных заголовках).`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Колонка или диапазон: "A:A", "B:B" или "A:D" (все колонки получат одинаковую ширину)',
      },
      width: {
        type: "number",
        description: "Ширина в символах (стандартная ~8.43, для данных ~15-20, для коротких текстов ~10-12)",
      },
    },
    required: ["address", "width"],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }
    const width = Number(args.width);
    if (!Number.isFinite(width) || width <= 0) {
      return {
        ok: false,
        summary: "Некорректная ширина",
        error: { code: "INVALID_WIDTH", message: "width должен быть положительным числом", retryable: false },
      };
    }

    return Excel.run(async (context) => {
      // Валидация адреса через getRangeSafe
      getRangeSafe(context, address);
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Для каждой колонки в адресе — устанавливаем ширину
      const colMatch = address.match(/^([A-Z]+):([A-Z]+)$/);
      if (colMatch) {
        const startCol = colMatch[1];
        const endCol = colMatch[2];
        const cols = getColumnLetters(startCol, endCol);
        for (const col of cols) {
          sheet.getRange(`${col}:${col}`).format.columnWidth = width;
        }
      } else {
        // Одиночная колонка или диапазон — ширина для первой колонки
        const firstCol = address.replace(/\d.*$/, "");
        if (firstCol) {
          sheet.getRange(`${firstCol}:${firstCol}`).format.columnWidth = width;
        }
      }
      await context.sync();

      return {
        ok: true,
        summary: `Установлена ширина колонок ${address} = ${width}`,
        data: { address, width },
      };
    });
  },
});

toolRegistry.registerDefinition(setColumnWidthsTool);

/**
 * Вспомогательная: список колонок от A до ZZ.
 * getColumnLetters("A", "D") → ["A","B","C","D"]
 */
function getColumnLetters(from: string, to: string): string[] {
  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  const result: string[] = [];
  let current = from;
  while (true) {
    result.push(current);
    if (current === to) break;
    // Инкремент колонки
    const chars = current.split("");
    let i = chars.length - 1;
    while (i >= 0) {
      const idx = letters.indexOf(chars[i]);
      if (idx < 25) {
        chars[i] = letters[idx + 1];
        break;
      }
      chars[i] = "A";
      i--;
    }
    if (i < 0) chars.unshift("A");
    current = chars.join("");
    if (current.length > 3) break; // предохранитель
  }
  return result;
}
