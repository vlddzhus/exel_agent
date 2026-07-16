/**
 * named-ranges.ts — Именованные диапазоны (категория Names).
 *
 * Итерация «Расширение инструментов». Новый инструмент:
 *   NR1 manageNamedRanges — add/list/delete/get именованных диапазонов.
 *
 * Безопасность: новый файл, существующие инструменты не трогаются.
 * Регистрация через defineTool + toolRegistry.registerDefinition (единый API).
 *
 * Office.js: workbook.names — коллекция NamedItem. Каждый может ссылаться
 * на Range или константу/формулу.
 */
import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { undoManager } from "./backup";

type NamedRangeAction = "add" | "list" | "delete" | "get";

export const manageNamedRangesTool = defineTool({
  name: "manageNamedRanges",
  description: `Управление именованными диапазонами workbook (имена для ячеек/диапазонов).
Действия:
  - add: создать имя, ссылающееся на диапазон. Параметры: name, refersTo ("=Sheet1!$A$1:$D$10" или "A1:D10"), comment (опц.).
  - list: показать все имена (имя → ссылка).
  - get: показать детали одного имени.
  - delete: удалить имя.
Используй для "назови диапазон A1:D10 как SalesData", "создай имя для итоговой ячейки", "удали имя".
Имена удобны для формул: =SUM(SalesData) вместо =SUM(A1:D10).`,
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["add", "list", "delete", "get"],
        description: "Действие: add / list / delete / get",
      },
      name: {
        type: "string",
        description: 'Имя диапазона: "SalesData", "TaxRate" (без пробелов, не начинается с цифры)',
      },
      refersTo: {
        type: "string",
        description: 'Ссылка для add: "=Sheet1!$A$1:$D$10" или просто "A1:D10"',
      },
      comment: {
        type: "string",
        description: "Комментарий к имени (опционально)",
      },
    },
    required: ["action"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const action = String(args.action ?? "") as NamedRangeAction;
    if (!action || !["add", "list", "delete", "get"].includes(action)) {
      return {
        ok: false,
        summary: "Некорректный action",
        error: {
          code: "INVALID_ACTION",
          message: "action должен быть add/list/delete/get",
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const names = context.workbook.names;

      // ── LIST ──
      if (action === "list") {
        names.load("name, type, value, comment");
        await context.sync();
        const items = (names as any).items as any[];
        if (!items || items.length === 0) {
          return {
            ok: true,
            summary: `В книге нет именованных диапазонов`,
            data: { action, count: 0, names: [] },
          };
        }
        const list = items.map((it) => ({
          name: it.name,
          type: it.type,
          value: it.value,
          comment: it.comment,
        }));
        return {
          ok: true,
          summary: `Найдено ${list.length} именованных диапазонов`,
          data: { action, count: list.length, names: list },
        };
      }

      // Остальные действия требуют name
      const name = String(args.name ?? "");
      if (!name) {
        return {
          ok: false,
          summary: "name обязателен",
          error: {
            code: "MISSING_NAME",
            message: "name обязателен для add/delete/get",
            retryable: false,
          },
        };
      }

      // ── ADD ──
      if (action === "add") {
        const refersToRaw = String(args.refersTo ?? "");
        if (!refersToRaw) {
          return {
            ok: false,
            summary: "refersTo обязателен для add",
            error: {
              code: "MISSING_REFERS_TO",
              message: 'refersTo обязателен (например "A1:D10" или "=Sheet1!A1")',
              retryable: false,
            },
          };
        }
        // Нормализуем: добавляем ведущий "=", если отсутствует
        const refersTo = refersToRaw.startsWith("=")
          ? refersToRaw
          : `=${refersToRaw}`;
        const comment =
          typeof args.comment === "string" ? args.comment : undefined;

        await undoManager.createBackup(name, "manageNamedRanges:add", {
          description: `Создание имени "${name}" → ${refersTo}`,
        });

        try {
          const namedItem = names.add(name, refersTo, comment);
          namedItem.load("name, type, value, comment");
          await context.sync();
          return {
            ok: true,
            summary: `Создано имя "${name}" → ${refersTo}`,
            data: {
              action: "add",
              name: (namedItem as any).name,
              type: (namedItem as any).type,
              value: (namedItem as any).value,
              comment: (namedItem as any).comment,
            },
          };
        } catch (e: any) {
          return {
            ok: false,
            summary: `Не удалось создать имя "${name}"`,
            error: {
              code: "ADD_FAILED",
              message: e?.message ?? "Имя уже существует или невалидно",
              retryable: false,
            },
          };
        }
      }

      // ── GET ──
      if (action === "get") {
        try {
          const item = names.getItem(name);
          item.load("name, type, value, comment");
          await context.sync();
          return {
            ok: true,
            summary: `Имя "${name}" → ${(item as any).value}`,
            data: {
              action: "get",
              name: (item as any).name,
              type: (item as any).type,
              value: (item as any).value,
              comment: (item as any).comment,
            },
          };
        } catch {
          return {
            ok: false,
            summary: `Имя "${name}" не найдено`,
            error: {
              code: "NAME_NOT_FOUND",
              message: `Имя "${name}" не существует`,
              retryable: false,
            },
          };
        }
      }

      // ── DELETE ──
      if (action === "delete") {
        await undoManager.createBackup(name, "manageNamedRanges:delete", {
          description: `Удаление имени "${name}"`,
        });
        try {
          const item = names.getItem(name);
          item.delete();
          await context.sync();
          return {
            ok: true,
            summary: `Удалено имя "${name}"`,
            data: { action: "delete", name },
          };
        } catch {
          return {
            ok: false,
            summary: `Имя "${name}" не найдено`,
            error: {
              code: "NAME_NOT_FOUND",
              message: `Имя "${name}" не существует`,
              retryable: false,
            },
          };
        }
      }

      // Недостижимо (вышли раньше через валидацию action)
      return {
        ok: false,
        summary: "Неизвестное действие",
        error: {
          code: "UNKNOWN_ACTION",
          message: `Неподдерживаемое действие: ${action}`,
          retryable: false,
        },
      };
    });
  },
});

toolRegistry.registerDefinition(manageNamedRangesTool);
