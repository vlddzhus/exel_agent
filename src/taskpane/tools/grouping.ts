/**
 * grouping.ts — Группировка строк/столбцов (категория Outline/Group).
 *
 * Итерация «Расширение инструментов». Новый инструмент:
 *   GR1 manageGrouping — group/ungroup/clearOutline/summaryBelow.
 *
 * Безопасность: новый файл, существующие инструменты не трогаются.
 * Регистрация через defineTool + toolRegistry.registerDefinition (единый API).
 *
 * Office.js: Range.group(groupBy: "ByRows" | "ByColumns") создаёт уровень
 * структуры. Range.ungroup() — убирает. Worksheet.outline — настройки
 * (summaryBottom, summaryRight) и clearMethods().
 *
 * Группировка = «сворачиваемые секции» в финмоделях и отчётах: кнопки +/- слева
 * от номеров строк позволяют скрыть детали, оставив итоги.
 */
import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";
import { undoManager } from "./backup";

type GroupingAction =
  | "groupRows"
  | "groupColumns"
  | "ungroupRows"
  | "ungroupColumns"
  | "clearOutline";

export const manageGroupingTool = defineTool({
  name: "manageGrouping",
  description: `Группировка строк/столбцов — сворачиваемые секции (outline).
Действия:
  - groupRows: сгруппировать строки диапазона (появятся кнопки +/- слева)
  - groupColumns: сгруппировать столбцы диапазона (кнопки +/- сверху)
  - ungroupRows: убрать группировку строк
  - ungroupColumns: убрать группировку столбцов
  - clearOutline: очистить всю структуру на листе
Опционально summaryBelow (true — итоги под группой, false — над),
summaryRight (true — итоги справа от группы столбцов, false — слева).
Используй для "сгруппируй строки 5-10", "сделай сворачиваемые секции", "спрячь детали под итогами".`,
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["groupRows", "groupColumns", "ungroupRows", "ungroupColumns", "clearOutline"],
        description: "Действие над группировкой",
      },
      address: {
        type: "string",
        description: 'Диапазон для group/ungroup: "A5:A10" (строки 5-10), "C:F" (колонки C-F). Для clearOutline — игнорируется.',
      },
      sheetName: {
        type: "string",
        description: "Имя листа (опционально, по умолчанию активный)",
      },
      summaryBelow: {
        type: "boolean",
        description: "Итоги строк под группой (true) или над (false). По умолчанию true.",
      },
      summaryRight: {
        type: "boolean",
        description: "Итоги колонок справа (true) или слева (false). По умолчанию true.",
      },
    },
    required: ["action"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 10,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const action = String(args.action ?? "") as GroupingAction;
    const validActions: GroupingAction[] = [
      "groupRows",
      "groupColumns",
      "ungroupRows",
      "ungroupColumns",
      "clearOutline",
    ];
    if (!action || !validActions.includes(action)) {
      return {
        ok: false,
        summary: "Некорректный action",
        error: {
          code: "INVALID_ACTION",
          message: `action должен быть одним из: ${validActions.join("/")}`,
          retryable: false,
        },
      };
    }

    const sheetName = typeof args.sheetName === "string" ? args.sheetName : undefined;
    const rawAddress = typeof args.address === "string" ? args.address : undefined;

    // clearOutline не требует адреса; остальным нужен
    if (action !== "clearOutline" && !rawAddress) {
      return {
        ok: false,
        summary: "address обязателен",
        error: {
          code: "MISSING_ADDRESS",
          message: "address обязателен для group/ungroup",
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");

      await undoManager.createBackup(sheetName ?? "active", `manageGrouping:${action}`, {
        description: `${action} на листе "${(sheet as any).name}"`,
      });

      // Применяем настройки outline (если указаны)
      if (typeof args.summaryBelow === "boolean") {
        (sheet as any).outline.summaryBelow = args.summaryBelow;
      }
      if (typeof args.summaryRight === "boolean") {
        (sheet as any).outline.summaryRight = args.summaryRight;
      }

      // ── clearOutline ──
      if (action === "clearOutline") {
        (sheet as any).outline.clearMethods?.();
        // Если clearMethods недоступен — пробуем через usedRange
        try {
          const used = sheet.getUsedRangeOrNullObject();
          used.load("isNullObject");
          await context.sync();
          if (!(used as any).isNullObject) {
            (used as any).ungroup();
          }
        } catch {
          // ignore — clearMethods уже выполнен
        }
        await context.sync();
        return {
          ok: true,
          summary: `Структура очищена на листе "${(sheet as any).name}"`,
          data: { action, sheetName: (sheet as any).name },
        };
      }

      // ── group / ungroup ──
      const range = getRangeSafe(context, rawAddress as string);
      const isRows = action === "groupRows" || action === "ungroupRows";
      const isUngroup = action === "ungroupRows" || action === "ungroupColumns";

      if (isUngroup) {
        (range as any).ungroup();
      } else {
        (range as any).group(isRows ? "ByRows" : "ByColumns");
      }
      await context.sync();

      return {
        ok: true,
        summary: `${action}: ${rawAddress} на листе "${(sheet as any).name}"`,
        data: {
          action,
          address: rawAddress,
          sheetName: (sheet as any).name,
          axis: isRows ? "rows" : "columns",
        },
        cellsAffected: 10,
      };
    });
  },
});

toolRegistry.registerDefinition(manageGroupingTool);
