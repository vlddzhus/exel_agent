import { toolRegistry } from "./registry";
import { withPerformanceGuard } from "./performance-guard";
import { getRangeSafe } from "./address-helper";
import {
  validateFormula,
  letterToColumn,
  columnToLetter,
  parseCellAddress,
  parseRangeAddress,
} from "./formula-guardian";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Merge sourceCell into targetRange, ensuring source is ALWAYS included.
 *
 * Office.js autoFill requires destination range INCLUDING the source cell.
 * Исправленная версия: использует числовое сравнение колонок вместо строкового.
 *
 * Примеры:
 *   source "B2", target "B3:B100" → "B2:B100"
 *   source "D2", target "A1:E100" → "A1:E100" (уже покрывает source)
 *   source "Z5", target "AA1:AA10" → "Z5:AA10" (корректное сравнение AA vs Z)
 */
function mergeRangeAddress(sourceCell: string, targetRange: string): string {
  const src = parseCellAddress(sourceCell);
  const tgt = parseRangeAddress(targetRange);

  const srcColNum = letterToColumn(src.col);
  const tgtStartColNum = letterToColumn(tgt.startCol);
  const tgtEndColNum = letterToColumn(tgt.endCol);

  const mergedStartColNum = Math.min(srcColNum, tgtStartColNum);
  const mergedStartRow = Math.min(src.row, tgt.startRow);
  const mergedEndColNum = Math.max(srcColNum, tgtEndColNum);
  const mergedEndRow = Math.max(src.row, tgt.endRow);

  const mergedStartCol = columnToLetter(mergedStartColNum);
  const mergedEndCol = columnToLetter(mergedEndColNum);

  return `${mergedStartCol}${mergedStartRow}:${mergedEndCol}${mergedEndRow}`;
}

// ---------------------------------------------------------------------------
// fillFormula
// ---------------------------------------------------------------------------

toolRegistry.register(
  "fillFormula",
  "Write a formula to a cell and auto-fill it down a column. Uses Excel autoFill engine so relative references adjust correctly.",
  {
    type: "object",
    properties: {
      sheetName: {
        type: "string",
        description: 'Sheet name (e.g., "Sheet1"). Use active sheet if empty.',
      },
      sourceCell: {
        type: "string",
        description: 'Cell to write the formula into first (e.g., "B2")',
      },
      targetRange: {
        type: "string",
        description:
          'Full range to fill. Should start at the row after sourceCell (e.g., "B3:B100" for sourceCell="B2")',
      },
      formula: {
        type: "string",
        description:
          'Excel formula without leading "=" (e.g., "SUM(A2:C2)" or "A2*B2")',
      },
    },
    required: ["sourceCell", "targetRange", "formula"],
  },
  async (args) => {
    const sourceCell = args.sourceCell as string;
    const targetRange = args.targetRange as string;
    const formula = args.formula as string;
    const sheetName = args.sheetName as string | undefined;

    // Валидация формулы перед записью
    const validation = validateFormula(formula);
    if (!validation.valid) {
      throw new Error(
        `Некорректная формула: "${formula}". ${validation.error || "Неизвестная ошибка"}` +
          (validation.fixedFormula
            ? ` Возможное исправление: "${validation.fixedFormula}"`
            : ""),
      );
    }

    const finalFormula = validation.fixedFormula || formula;
    const mergedRange = mergeRangeAddress(sourceCell, targetRange);

    return withPerformanceGuard(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const source = sheet.getRange(sourceCell);
      const destination = sheet.getRange(mergedRange);

      source.formulas = [[`=${finalFormula}`]];
      await context.sync();

      source.autoFill(destination, Excel.AutoFillType.fillDefault);
      await context.sync();
    }).then(() =>
      JSON.stringify({
        success: true,
        sourceCell,
        mergedRange,
        formula: finalFormula,
        fillType: "fillDefault",
      }),
    );
  },
);

// ---------------------------------------------------------------------------
// setFormula
// ---------------------------------------------------------------------------

toolRegistry.register(
  "setFormula",
  "Set a formula in a specific cell or range. Validates and auto-fixes formula before writing.",
  {
    type: "object",
    properties: {
      cellAddress: {
        type: "string",
        description: 'Cell address like "A1" or "Sheet1!B2"',
      },
      formula: {
        type: "string",
        description: 'Excel formula without leading "=", e.g., "SUM(A1:A10)"',
      },
    },
    required: ["cellAddress", "formula"],
  },
  async (args) => {
    const cellAddress = args.cellAddress as string;
    const formula = args.formula as string;

    // Валидация формулы перед записью
    const validation = validateFormula(formula);
    if (!validation.valid) {
      throw new Error(
        `Некорректная формула: "${formula}". ${validation.error || "Неизвестная ошибка"}` +
          (validation.fixedFormula
            ? ` Возможное исправление: "${validation.fixedFormula}"`
            : ""),
      );
    }

    const finalFormula = validation.fixedFormula || formula;

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, cellAddress);
      range.formulas = [[`=${finalFormula}`]];
      await context.sync();
      return JSON.stringify({
        success: true,
        cell: cellAddress,
        formula: finalFormula,
      });
    });
  },
);

// ---------------------------------------------------------------------------
// applyFormat — enhanced: string (entire range) OR 2D array (per-cell)
// ---------------------------------------------------------------------------

toolRegistry.register(
  "applyFormat",
  "Apply number format to a range. Accepts either a single format string (applied to entire range) or a 2D array for per-cell formatting.",
  {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Range address (e.g., "B2:B100")',
      },
      format: {
        type: "string",
        description:
          'Number format string like "#,##0.00", "$#,##0", "0%" for whole range, ' +
          'or a JSON stringified 2D array like \'[["$#,##0.00","0%"]]\' for per-cell formatting',
      },
    },
    required: ["address", "format"],
  },
  async (args) => {
    const address = args.address as string;
    const format = args.format as string;
    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);

      // Пытаемся распарсить как JSON-массив (2D формат)
      // Если не парсится — используем как строку формата для всего диапазона
      let parsedFormat: unknown;
      try {
        parsedFormat = JSON.parse(format);
      } catch {
        parsedFormat = format;
      }

      if (typeof parsedFormat === "string") {
        range.numberFormat = [[parsedFormat]];
      } else if (Array.isArray(parsedFormat)) {
        range.numberFormat = parsedFormat as any[][];
      } else {
        range.numberFormat = [[String(parsedFormat)]];
      }

      await context.sync();
      return JSON.stringify({ success: true, range: address, format });
    });
  },
);

// ---------------------------------------------------------------------------
// mergeCells — new tool
// ---------------------------------------------------------------------------

toolRegistry.register(
  "mergeCells",
  "Merge cells in a specified range. Only the top-left cell value is kept; all other values are cleared.",
  {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Range address to merge (e.g., "A1:C1" for a heading row)',
      },
    },
    required: ["address"],
  },
  async (args) => {
    const address = args.address as string;
    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.merge();
      await context.sync();
      return JSON.stringify({ success: true, merged: address });
    });
  },
);

// ---------------------------------------------------------------------------
// setCellFormat — new tool for font / border / alignment
// ---------------------------------------------------------------------------

toolRegistry.register(
  "setCellFormat",
  "Set font, alignment, border, or fill formatting on a cell or range.",
  {
    type: "object",
    properties: {
      address: { type: "string", description: 'Range address (e.g., "A1:E1")' },
      bold: { type: "boolean", description: "Make text bold (true/false)" },
      fontSize: {
        type: "number",
        description: "Font size in points (e.g., 12, 14, 16)",
      },
      fontColor: {
        type: "string",
        description:
          'Font color: CSS color name ("green", "red", "blue") or hex "#FF0000"',
      },
      horizontalAlignment: {
        type: "string",
        enum: ["left", "center", "right"],
        description: "Horizontal text alignment",
      },
      verticalAlignment: {
        type: "string",
        enum: ["top", "center", "bottom"],
        description: "Vertical text alignment",
      },
      borderTop: {
        type: "string",
        enum: ["thin", "medium", "thick", "none"],
        description: "Top border style",
      },
      borderBottom: {
        type: "string",
        enum: ["thin", "medium", "thick", "none"],
        description: "Bottom border style",
      },
      borderLeft: {
        type: "string",
        enum: ["thin", "medium", "thick", "none"],
        description: "Left border style",
      },
      borderRight: {
        type: "string",
        enum: ["thin", "medium", "thick", "none"],
        description: "Right border style",
      },
      wrapText: { type: "boolean", description: "Wrap text within cells" },
      fillColor: {
        type: "string",
        description:
          'Background fill color: CSS name or hex (e.g., "#FFFF00" for yellow)',
      },
    },
    required: ["address"],
  },
  async (args) => {
    const address = args.address as string;
    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      const fmt = range.format;

      // Font
      if (args.bold !== undefined) fmt.font.bold = args.bold as boolean;
      if (args.fontSize !== undefined) fmt.font.size = args.fontSize as number;
      if (args.fontColor !== undefined) {
        fmt.font.color = normalizeColor(args.fontColor as string);
      }

      // Horizontal alignment
      if (args.horizontalAlignment !== undefined) {
        const alignMap: Record<string, Excel.HorizontalAlignment> = {
          left: Excel.HorizontalAlignment.left,
          center: Excel.HorizontalAlignment.center,
          right: Excel.HorizontalAlignment.right,
        };
        fmt.horizontalAlignment = alignMap[args.horizontalAlignment as string];
      }

      // Vertical alignment
      if (args.verticalAlignment !== undefined) {
        const alignMap: Record<string, Excel.VerticalAlignment> = {
          top: Excel.VerticalAlignment.top,
          center: Excel.VerticalAlignment.center,
          bottom: Excel.VerticalAlignment.bottom,
        };
        fmt.verticalAlignment = alignMap[args.verticalAlignment as string];
      }

      // Borders
      if (args.borderTop !== undefined)
        setBorder(fmt, "top", args.borderTop as string);
      if (args.borderBottom !== undefined)
        setBorder(fmt, "bottom", args.borderBottom as string);
      if (args.borderLeft !== undefined)
        setBorder(fmt, "left", args.borderLeft as string);
      if (args.borderRight !== undefined)
        setBorder(fmt, "right", args.borderRight as string);

      // Wrap text
      if (args.wrapText !== undefined) fmt.wrapText = args.wrapText as boolean;

      // Fill color
      if (args.fillColor !== undefined) {
        fmt.fill.color = normalizeColor(args.fillColor as string);
      }

      await context.sync();

      const appliedKeys = Object.entries(args)
        .filter(([k, v]) => k !== "address" && v !== undefined)
        .map(([k]) => k);

      return JSON.stringify({
        success: true,
        range: address,
        applied: appliedKeys,
      });
    });
  },
);

// ---------------------------------------------------------------------------
// Internal helpers for setCellFormat
// ---------------------------------------------------------------------------

/**
 * Приводит название цвета к формату, понятному Excel.
 * Если это известное CSS-имя — преобразует в hex.
 * Если уже hex — оставляет как есть.
 */
function normalizeColor(color: string): string {
  const lower = color.toLowerCase().trim();

  // Известные CSS-цвета → hex
  const namedColors: Record<string, string> = {
    green: "#00FF00",
    red: "#FF0000",
    blue: "#0000FF",
    yellow: "#FFFF00",
    orange: "#FFA500",
    purple: "#800080",
    pink: "#FFC0CB",
    brown: "#A52A2A",
    black: "#000000",
    white: "#FFFFFF",
    gray: "#808080",
    grey: "#808080",
    darkgreen: "#006400",
    darkred: "#8B0000",
    darkblue: "#00008B",
    lightgray: "#D3D3D3",
    lightgrey: "#D3D3D3",
    cyan: "#00FFFF",
    magenta: "#FF00FF",
    navy: "#000080",
    teal: "#008080",
    olive: "#808000",
    maroon: "#800000",
    silver: "#C0C0C0",
    lime: "#00FF00",
    aqua: "#00FFFF",
    fuchsia: "#FF00FF",
  };

  if (namedColors[lower]) return namedColors[lower];
  if (color.startsWith("#")) return color;

  // Возвращаем как есть — Excel может понять
  return color;
}

/**
 * Устанавливает стиль границы для указанной стороны ячейки.
 * Office.js API: RangeBorderCollection.getItem(BorderIndex) → RangeBorder
 * Свойства: style (BorderLineStyle), weight (BorderWeight)
 */
function setBorder(fmt: Excel.RangeFormat, side: string, style: string): void {
  const border = fmt.borders.getItem(sideIndex(side));
  if (style === "none") {
    border.style = Excel.BorderLineStyle.none;
    return;
  }

  border.style = Excel.BorderLineStyle.continuous;

  const weightMap: Record<string, Excel.BorderWeight> = {
    thin: Excel.BorderWeight.thin,
    medium: Excel.BorderWeight.medium,
    thick: Excel.BorderWeight.thick,
  };
  border.weight = weightMap[style] || Excel.BorderWeight.thin;
}

/**
 * Преобразует строковое имя стороны в BorderIndex.
 */
function sideIndex(side: string): Excel.BorderIndex {
  const map: Record<string, Excel.BorderIndex> = {
    top: Excel.BorderIndex.edgeTop,
    bottom: Excel.BorderIndex.edgeBottom,
    left: Excel.BorderIndex.edgeLeft,
    right: Excel.BorderIndex.edgeRight,
  };
  return map[side] || Excel.BorderIndex.edgeTop;
}
