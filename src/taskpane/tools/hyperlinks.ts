/**
 * hyperlinks.ts — Гиперссылки (категория Hyperlinks).
 *
 * Итерация «Расширение инструментов». Новый инструмент:
 *   HL1 manageHyperlinks — add/get/remove гиперссылок в ячейках.
 *
 * Безопасность: новый файл, существующие инструменты не трогаются.
 * Регистрация через defineTool + toolRegistry.registerDefinition (единый API).
 *
 * Office.js: Range.hyperlink — объект { address, documentReference, screenTip, textToDisplay }.
 * Для добавления URL-ссылки: range.hyperlink = { address: "https://..." }.
 * Для внутренней ссылки: { documentReference: "Sheet1!A1" }.
 * Для удаления: range.hyperlink = null (или {}).
 */
import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";
import { undoManager } from "./backup";

type HyperlinkAction = "add" | "get" | "remove";

export const manageHyperlinksTool = defineTool({
  name: "manageHyperlinks",
  description: `Управление гиперссылками в ячейках.
Действия:
  - add: добавить ссылку. Параметры: address (ячейка), target (URL или "Лист!A1" для внутренней), screenTip (опц.), textToDisplay (опц.).
  - get: прочитать текущую ссылку.
  - remove: удалить ссылку.
Примеры target:
  - URL: "https://example.com"
  - Внутренняя: "Sheet2!A1" (переход внутри книги)
  - Email: "mailto:info@example.com"
Используй для "добавь ссылку на сайт", "сделай навигацию по листам", "ссылку на email".`,
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["add", "get", "remove"],
        description: "Действие: add / get / remove",
      },
      address: {
        type: "string",
        description: 'Адрес ячейки: "A1" или "Лист!B2"',
      },
      target: {
        type: "string",
        description: 'URL или ссылка: "https://example.com", "Sheet2!A1", "mailto:a@b.com"',
      },
      screenTip: {
        type: "string",
        description: "Всплывающая подсказка (опционально)",
      },
      textToDisplay: {
        type: "string",
        description: "Отображаемый текст ссылки (опционально)",
      },
    },
    required: ["action", "address"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 1,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const action = String(args.action ?? "") as HyperlinkAction;
    if (!action || !["add", "get", "remove"].includes(action)) {
      return {
        ok: false,
        summary: "Некорректный action",
        error: {
          code: "INVALID_ACTION",
          message: "action должен быть add/get/remove",
          retryable: false,
        },
      };
    }

    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "Не указан адрес ячейки",
        error: {
          code: "MISSING_ADDRESS",
          message: "address обязателен",
          retryable: false,
        },
      };
    }

    return Excel.run(async (context) => {
      const range = getRangeSafe(context, address);
      range.load("address");

      // ── ADD ──
      if (action === "add") {
        const target = String(args.target ?? "");
        if (!target) {
          return {
            ok: false,
            summary: "target обязателен для add",
            error: {
              code: "MISSING_TARGET",
              message: "target (URL или ссылка) обязателен",
              retryable: false,
            },
          };
        }

        await undoManager.createBackup(address, "manageHyperlinks:add", {
          description: `Гиперссылка в ${address} → ${target}`,
        });

        // Различаем URL и внутреннюю ссылку.
        // Если target начинается с протокола/mailto/# — это address,
        // иначе считаем documentReference (внутренняя).
        const isExternal =
          /^[a-z][a-z0-9+.-]*:/i.test(target) || target.startsWith("#");
        const hyperlink: Record<string, string> = {};
        if (isExternal) {
          hyperlink.address = target.startsWith("#")
            ? target.slice(1)
            : target;
        } else {
          hyperlink.documentReference = target;
        }
        if (typeof args.screenTip === "string") hyperlink.screenTip = args.screenTip;
        if (typeof args.textToDisplay === "string") {
          hyperlink.textToDisplay = args.textToDisplay;
        }

        (range as any).hyperlink = hyperlink;
        await context.sync();
        return {
          ok: true,
          summary: `Добавлена ссылка в ${(range as any).address} → ${target}`,
          data: { action: "add", address: (range as any).address, target, type: isExternal ? "external" : "internal" },
        };
      }

      // ── GET ──
      if (action === "get") {
        (range as any).load("hyperlink");
        await context.sync();
        const hl = (range as any).hyperlink;
        if (!hl) {
          return {
            ok: true,
            summary: `В ${(range as any).address} нет гиперссылки`,
            data: { action: "get", address: (range as any).address, hyperlink: null },
          };
        }
        return {
          ok: true,
          summary: `Ссылка в ${(range as any).address}: ${hl.address ?? hl.documentReference ?? "(нет)"}`,
          data: { action: "get", address: (range as any).address, hyperlink: hl },
        };
      }

      // ── REMOVE ──
      await undoManager.createBackup(address, "manageHyperlinks:remove", {
        description: `Удаление гиперссылки из ${address}`,
      });
      (range as any).hyperlink = null;
      await context.sync();
      return {
        ok: true,
        summary: `Удалена ссылка из ${(range as any).address}`,
        data: { action: "remove", address: (range as any).address },
      };
    });
  },
});

toolRegistry.registerDefinition(manageHyperlinksTool);
