/**
 * data-validation.ts — Data Validation инструменты (выпадающие списки в Excel).
 *
 * Инструменты:
 *   V1 addDataValidation — добавить проверку данных (выпадающий список).
 */

import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";
import { undoManager } from "./backup";

// ============================================================================
// V1. addDataValidation — выпадающий список значений (Data Validation)
// ============================================================================

export type ValidationType = "list" | "wholeNumber" | "decimal" | "textLength" | "date" | "custom";

export interface DataValidationOptions {
  /** Тип проверки. "list" — выпадающий список (по умолчанию). */
  type?: ValidationType;
  /** Для list: строка значений через запятую, например "Да,Нет" */
  listSource?: string;
  /** Для list: ссылка на диапазон-источник, например "Sheet2!A1:A10" */
  listSourceRange?: string;
  /** Показывать выпадающую стрелку (по умолчанию true) */
  inCellDropDown?: boolean;
  /** Разрешать пустое значение (по умолчанию true) */
  allowBlank?: boolean;
  /** Заголовок подсказки */
  promptTitle?: string;
  /** Текст подсказки */
  promptMessage?: string;
  /** Заголовок ошибки */
  errorTitle?: string;
  /** Текст ошибки */
  errorMessage?: string;
  /** Стиль ошибки: Stop, Warning, Information */
  errorStyle?: "Stop" | "Warning" | "Information";
}

export const addDataValidationTool = defineTool({
  name: "addDataValidation",
  description: `Добавляет проверку данных (выпадающий список) к диапазону ячеек.
Используй когда нужно создать выбор из вариантов: "Да/Нет", "мужской/женский", цвета, статусы и т.п.
Основное применение — type="list" с listSource="Да,Нет" (значения через запятую).
Управляет: allowBlank (пустое разрешено), inCellDropDown (стрелка), подсказки (prompt), ошибки (error).`,
  parameters: {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Адрес: "A1:A100" или "Лист1!B2:B50"',
      },
      type: {
        type: "string",
        enum: ["list", "wholeNumber", "decimal", "textLength", "date", "custom"],
        description: 'Тип проверки. "list" — выпадающий список (по умолчанию).',
      },
      listSource: {
        type: "string",
        description: 'Для type="list": значения через запятую, напр. "Да,Нет" или "Красный,Зелёный,Синий"',
      },
      listSourceRange: {
        type: "string",
        description: 'Для type="list": ссылка на диапазон-источник, напр. "Sheet2!A1:A10". Альтернатива listSource.',
      },
      inCellDropDown: {
        type: "boolean",
        description: "Показывать стрелку выпадающего списка (по умолчанию true)",
      },
      allowBlank: {
        type: "boolean",
        description: "Разрешать пустые значения (по умолчанию true)",
      },
      promptTitle: { type: "string", description: "Заголовок всплывающей подсказки" },
      promptMessage: { type: "string", description: "Текст всплывающей подсказки" },
      errorTitle: { type: "string", description: "Заголовок сообщения об ошибке" },
      errorMessage: { type: "string", description: "Текст сообщения об ошибке" },
      errorStyle: {
        type: "string",
        enum: ["Stop", "Warning", "Information"],
        description: "Стиль ошибки: Stop — блокирует ввод, Warning — предупреждает, Information — информирует",
      },
    },
    required: ["address"],
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
        error: { code: "MISSING_ADDRESS", message: "address обязателен", retryable: false },
      };
    }

    const validationType = String(args.type ?? "list") as ValidationType;
    const inCellDropDown = args.inCellDropDown !== false;
    const allowBlank = args.allowBlank !== false;

    const listSource = typeof args.listSource === "string" ? args.listSource : undefined;
    const listSourceRange = typeof args.listSourceRange === "string" ? args.listSourceRange : undefined;
    const promptTitle = typeof args.promptTitle === "string" ? args.promptTitle : undefined;
    const promptMessage = typeof args.promptMessage === "string" ? args.promptMessage : undefined;
    const errorTitle = typeof args.errorTitle === "string" ? args.errorTitle : undefined;
    const errorMessage = typeof args.errorMessage === "string" ? args.errorMessage : undefined;
    const errorStyle = (typeof args.errorStyle === "string" ? args.errorStyle : "Stop") as "Stop" | "Warning" | "Information";

    // Валидация: для list нужен хотя бы один источник
    if (validationType === "list" && !listSource && !listSourceRange) {
      return {
        ok: false,
        summary: "Для list требуется listSource или listSourceRange",
        error: {
          code: "MISSING_LIST_SOURCE",
          message: "Укажите listSource (значения через запятую) или listSourceRange (диапазон)",
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("rowCount, columnCount, address");
      await context.sync();

      const cellCount = range.rowCount * range.columnCount;

      await undoManager.createBackup(address, "addDataValidation", {
        description: `Data Validation на ${range.address} (${cellCount} ячеек, ${validationType})`,
      });

      const dv = range.dataValidation;

      // Строим правило валидации
      const rule: any = {};

      if (validationType === "list") {
        if (listSource) {
          rule.list = { inCellDropDown, source: listSource };
        } else if (listSourceRange) {
          rule.list = { inCellDropDown, source: listSourceRange };
        }
      } else if (validationType === "wholeNumber") {
        rule.wholeNumber = {};
      } else if (validationType === "decimal") {
        rule.decimal = {};
      } else if (validationType === "textLength") {
        rule.textLength = {};
      } else if (validationType === "date") {
        rule.date = {};
      } else if (validationType === "custom") {
        rule.custom = {};
      }

      dv.rule = rule;
      dv.ignoreBlanks = allowBlank;

      if (promptTitle || promptMessage) {
        dv.prompt = {
          title: promptTitle ?? "",
          message: promptMessage ?? "",
          showPrompt: true,
        };
      }

      if (errorTitle || errorMessage) {
        dv.errorAlert = {
          title: errorTitle ?? "",
          message: errorMessage ?? "",
          style: errorStyle,
          showAlert: true,
        };
      }

      await context.sync();

      return {
        ok: true,
        summary: `Добавлена проверка данных (${validationType}) на ${range.address} (${cellCount} ячеек)`,
        data: { address: range.address, cellCount, type: validationType },
        cellsAffected: cellCount,
      };
    });
  },
});

toolRegistry.registerDefinition(addDataValidationTool);
