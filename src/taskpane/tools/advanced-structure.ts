/**
 * advanced-structure.ts — Продвинутые структурные инструменты (Tier 1).
 *
 * Инструменты:
 *   A1 manageRowsColumns — вставка/удаление строк и столбцов.
 *   A2 mergeCells / unmergeCells — объединение/разъединение ячеек.
 *   A3 manageSheetProtection — защита/снятие защиты листа.
 */

import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";
import { undoManager } from "./backup";

// ============================================================================
// A1. manageRowsColumns — вставка/удаление строк и столбцов
// ============================================================================

export const manageRowsColumnsTool = defineTool({
  name: "manageRowsColumns",
  description: `Вставляет или удаляет строки/столбцы на листе.
Действия:
  - insertRows: вставить строки. count (по умолчанию 1), position (номер строки, 1-based, по умолчанию после последней строки usedRange)
  - deleteRows: удалить строки. count (по умолчанию 1), position (номер строки, 1-based)
  - insertColumns: вставить столбцы. count (по умолчанию 1), position (номер колонки, 1-based)
  - deleteColumns: удалить столбцы. count (по умолчанию 1), position (номер колонки, 1-based)
  - address: альтернатива position через адрес диапазона ("A5" = перед строкой 5)
Используй для "вставь строку между 5 и 6", "удали колонку B", "добавь 3 строки вниз".`,
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["insertRows", "deleteRows", "insertColumns", "deleteColumns"],
        description: "Действие: insertRows / deleteRows / insertColumns / deleteColumns",
      },
      count: {
        type: "number",
        description: "Количество строк/колонок (по умолчанию 1)",
      },
      position: {
        type: "number",
        description: "Номер (1-based): строки для insertRows/deleteRows, или колонки для insertColumns/deleteColumns",
      },
      address: {
        type: "string",
        description: 'Адрес ячейки: "A5" = перед строкой 5, "B1" = перед колонкой B. Альтернатива position.',
      },
      sheetName: {
        type: "string",
        description: "Имя листа (опционально, по умолчанию активный)",
      },
    },
    required: ["action"],
  },
  riskLevel: "dangerous",
  requiresUndo: true,
  estimateCells: () => 100,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const action = String(args.action ?? "");
    if (!action) {
      return {
        ok: false,
        summary: "action обязателен",
        error: { code: "MISSING_ACTION", message: "action обязателен", retryable: false },
      };
    }

    const count = Math.max(1, Math.floor(Number(args.count ?? 1)));
    const sheetName = typeof args.sheetName === "string" ? args.sheetName : undefined;

    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      // Определяем позицию: из position, address или авто
      let targetRow = 1;
      let targetCol = 1;
      const rawPosition = Number(args.position);
      const rawAddress = typeof args.address === "string" ? args.address : undefined;

      if (Number.isFinite(rawPosition) && rawPosition >= 1) {
        if (action === "insertRows" || action === "deleteRows") {
          targetRow = rawPosition;
        } else {
          targetCol = rawPosition;
        }
      } else if (rawAddress) {
        const range = getRangeSafe(context, rawAddress);
        range.load("rowIndex, columnIndex, address");
        await context.sync();
        targetRow = range.rowIndex + 1;
        targetCol = range.columnIndex + 1;
      } else {
        // Авто: последняя строка/колонка usedRange
        const used = sheet.getUsedRangeOrNullObject();
        used.load("rowCount, columnCount, isNullObject");
        await context.sync();
        if (action === "insertRows" || action === "deleteRows") {
          targetRow = used.isNullObject ? 1 : used.rowCount + 1;
        } else {
          targetCol = used.isNullObject ? 1 : used.columnCount + 1;
        }
      }

      await undoManager.createBackup(sheetName ?? sheet.name, action, {
        description: `${action} × ${count} на листе "${sheet.name}"`,
      });

      // Office.js API: insert/delete строк и колонок
      if (action === "insertRows") {
        const endRow = targetRow + count - 1;
        sheet.getRange(`${targetRow}:${endRow}`).insert("Down");
      } else if (action === "deleteRows") {
        const endRow = targetRow + count - 1;
        sheet.getRange(`${targetRow}:${endRow}`).delete("Up");
      } else if (action === "insertColumns") {
        const startLetter = columnIndexToLetter(targetCol);
        const endLetter = columnIndexToLetter(targetCol + count - 1);
        sheet.getRange(`${startLetter}:${endLetter}`).insert("Right");
      } else if (action === "deleteColumns") {
        const startLetter = columnIndexToLetter(targetCol);
        const endLetter = columnIndexToLetter(targetCol + count - 1);
        sheet.getRange(`${startLetter}:${endLetter}`).delete("Left");
      }

      await context.sync();

      return {
        ok: true,
        summary: `${action}: ${count} шт. на листе "${sheet.name}"`,
        data: { action, count, sheetName: sheet.name },
        cellsAffected: count * 10,
      };
    });
  },
});

toolRegistry.registerDefinition(manageRowsColumnsTool);

/** Преобразует 1-based индекс колонки в буквенное обозначение (1→A, 27→AA) */
function columnIndexToLetter(index: number): string {
  let result = "";
  let n = index;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result || "A";
}

// ============================================================================
// A2. mergeCells / unmergeCells — объединение/разъединение ячеек
// ============================================================================

export const mergeCellsTool = defineTool({
  name: "mergeCells",
  description: `Объединяет ячейки в указанном диапазоне в одну.
Содержимым результирующей ячейки становится значение верхней левой ячейки диапазона.
Используй для "объедини ячейки A1:C1 для заголовка", "сделай шапку во всю ширину".`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Диапазон: "A1:C1" или "A1:E1" для объединения заголовка',
      },
      sheetName: {
        type: "string",
        description: "Имя листа (опционально)",
      },
    },
    required: ["address"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 10,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    return Excel.run(async (context) => {
      const sheet = typeof args.sheetName === "string"
        ? context.workbook.worksheets.getItem(args.sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = getRangeSafe(context, address);
      range.load("address, rowCount, columnCount");
      await context.sync();
      const cellCount = range.rowCount * range.columnCount;

      await undoManager.createBackup(address, "mergeCells", {
        description: `Объединение ${range.address} (${cellCount} ячеек)`,
      });

      range.merge(true);
      await context.sync();

      return {
        ok: true,
        summary: `Объединены ячейки ${range.address} (${cellCount} → 1)`,
        data: { address: range.address, cellCount },
        cellsAffected: cellCount,
      };
    });
  },
});

toolRegistry.registerDefinition(mergeCellsTool);

export const unmergeCellsTool = defineTool({
  name: "unmergeCells",
  description: `Разъединяет ранее объединённые ячейки обратно в отдельные ячейки.
Значение остаётся только в верхней левой ячейке, остальные становятся пустыми.
Используй для "разъедини ячейки", "отмени объединение".`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Диапазон с объединёнными ячейками: "A1:C1"',
      },
      sheetName: {
        type: "string",
        description: "Имя листа (опционально)",
      },
    },
    required: ["address"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 10,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес",
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    return Excel.run(async (context) => {
      const sheet = typeof args.sheetName === "string"
        ? context.workbook.worksheets.getItem(args.sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = getRangeSafe(context, address);
      range.load("address, rowCount, columnCount");
      await context.sync();
      const cellCount = range.rowCount * range.columnCount;

      await undoManager.createBackup(address, "unmergeCells", {
        description: `Разъединение ${range.address} (${cellCount} ячеек)`,
      });

      range.unmerge();
      await context.sync();

      return {
        ok: true,
        summary: `Разъединены ячейки ${range.address} (1 → ${cellCount})`,
        data: { address: range.address, cellCount },
        cellsAffected: cellCount,
      };
    });
  },
});

toolRegistry.registerDefinition(unmergeCellsTool);

// ============================================================================
// A3. manageSheetProtection — защита/снятие защиты листа
// ============================================================================

export const manageSheetProtectionTool = defineTool({
  name: "manageSheetProtection",
  description: `Управляет защитой листа: protect (включить) или unprotect (отключить).
При protect можно указать password (опционально) и разрешения (allowInsertRows, allowDeleteColumns и т.д.).
Используй для "защити лист паролем", "сними защиту", "заблокируй лист от изменений".`,
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["protect", "unprotect"],
        description: "protect — включить защиту, unprotect — отключить",
      },
      password: {
        type: "string",
        description: "Пароль для защиты (опционально, только для protect)",
      },
      sheetName: {
        type: "string",
        description: "Имя листа (опционально, по умолчанию активный)",
      },
      allowInsertRows: { type: "boolean", description: "Разрешить вставку строк" },
      allowDeleteRows: { type: "boolean", description: "Разрешить удаление строк" },
      allowInsertColumns: { type: "boolean", description: "Разрешить вставку колонок" },
      allowDeleteColumns: { type: "boolean", description: "Разрешить удаление колонок" },
      allowSort: { type: "boolean", description: "Разрешить сортировку" },
      allowFilter: { type: "boolean", description: "Разрешить фильтры" },
      allowFormatCells: { type: "boolean", description: "Разрешить форматирование ячеек" },
      allowFormatRows: { type: "boolean", description: "Разрешить форматирование строк" },
      allowFormatColumns: { type: "boolean", description: "Разрешить форматирование колонок" },
    },
    required: ["action"],
  },
  riskLevel: "moderate",
  requiresUndo: false,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const action = String(args.action ?? "");
    if (!action || !["protect", "unprotect"].includes(action)) {
      return {
        ok: false,
        summary: "action должен быть protect или unprotect",
        error: { code: "INVALID_ACTION", message: "action обязателен", retryable: false },
      };
    }

    const password = typeof args.password === "string" ? args.password : undefined;
    const sheetName = typeof args.sheetName === "string" ? args.sheetName : undefined;

    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      if (action === "unprotect") {
        sheet.protection.unprotect(password ?? undefined);
        await context.sync();
        return {
          ok: true,
          summary: `Защита снята с листа "${sheet.name}"`,
          data: { sheetName: sheet.name, protected: false },
        };
      }

      // protect
      const protectionOptions: any = {};
      if (typeof args.allowInsertRows === "boolean") protectionOptions.allowInsertRows = args.allowInsertRows;
      if (typeof args.allowDeleteRows === "boolean") protectionOptions.allowDeleteRows = args.allowDeleteRows;
      if (typeof args.allowInsertColumns === "boolean") protectionOptions.allowInsertColumns = args.allowInsertColumns;
      if (typeof args.allowDeleteColumns === "boolean") protectionOptions.allowDeleteColumns = args.allowDeleteColumns;
      if (typeof args.allowSort === "boolean") protectionOptions.allowSort = args.allowSort;
      if (typeof args.allowFilter === "boolean") protectionOptions.allowAutoFilter = args.allowFilter;
      if (typeof args.allowFormatCells === "boolean") protectionOptions.allowFormatCells = args.allowFormatCells;
      if (typeof args.allowFormatRows === "boolean") protectionOptions.allowFormatRows = args.allowFormatRows;
      if (typeof args.allowFormatColumns === "boolean") protectionOptions.allowFormatColumns = args.allowFormatColumns;

      sheet.protection.protect(
        Object.keys(protectionOptions).length > 0 ? protectionOptions : undefined,
        password ?? undefined,
      );
      await context.sync();

      return {
        ok: true,
        summary: `Лист "${sheet.name}" защищён${password ? " (с паролем)" : ""}`,
        data: { sheetName: sheet.name, protected: true, hasPassword: !!password },
      };
    });
  },
});

toolRegistry.registerDefinition(manageSheetProtectionTool);
