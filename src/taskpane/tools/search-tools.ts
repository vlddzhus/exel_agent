/**
 * search-tools.ts — Поиск и замена (категория Search).
 *
 * Итерация «Расширение инструментов». Новый инструмент:
 *   SR1 findAndReplace — поиск значений/текста и замена по листу или диапазону.
 *
 * Безопасность: новый файл, существующие инструменты не трогаются.
 * Регистрация через defineTool + toolRegistry.registerDefinition (единый API).
 *
 * Office.js: Range.find(text, criteria) → RangeAreas. Для замены итерируем
 * найденные ячейки и записываем новые значения через range.values.
 */
import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";
import { undoManager } from "./backup";

type FindReplaceAction = "find" | "findAll" | "replace" | "replaceAll";

interface FindCriteria {
  matchCase?: boolean;
  matchWholeWord?: boolean;
  searchDirection?: "Forward" | "Backwards";
}

/**
 * Найти все совпадения в диапазоне.
 * Возвращает массив адресов найденных ячеек (загруженных через context.sync).
 */
async function findAllMatches(
  context: Excel.RequestContext,
  range: Excel.Range,
  findWhat: string,
  criteria: FindCriteria,
): Promise<string[]> {
  const addresses: string[] = [];
  // Office.js: первый find возвращает первое совпадение (RangeAreas),
  // последующие find с теми же критериями — следующие совпадения.
  // Цикл ограничен 1000 итерациями (защита от зацикливания).
  const MAX_ITERATIONS = 1000;
  let matchedRange: Excel.RangeAreas | null = null;

  for (let i = 0; i < MAX_ITERATIONS; i++) {
    const found = matchedRange
      ? (range as any).find(findWhat, {
          ...criteria,
          matchSequential: true,
        } as any)
      : (range as any).find(findWhat, criteria as any);
    // load для проверки существования
    found.load("address,isNullObject");
    await context.sync();
    if ((found as any).isNullObject) break;

    const addr = (found as any).address as string;
    if (addresses.includes(addr)) break; // зацикливание — выходим
    addresses.push(addr);
    matchedRange = found as any;
  }
  return addresses;
}

export const findAndReplaceTool = defineTool({
  name: "findAndReplace",
  description: `Поиск и замена текста/значений по листу или диапазону.
Действия:
  - find: найти первое совпадение (вернёт адрес)
  - findAll: найти все совпадения (вернёт список адресов)
  - replace: заменить первое совпадение
  - replaceAll: заменить все совпадения
Критерии (опционально): matchCase (учитывать регистр), matchWholeWord (слово целиком).
Если address не указан — поиск по всему usedRange активного листа.
Используй для "найди все ячейки со словом X", "замени A на B", "найди пустые".`,
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["find", "findAll", "replace", "replaceAll"],
        description: "Действие: find / findAll / replace / replaceAll",
      },
      findWhat: {
        type: "string",
        description: "Что искать (текст или значение)",
      },
      replaceWith: {
        type: "string",
        description: "На что заменять (для replace/replaceAll)",
      },
      address: {
        type: "string",
        description: 'Диапазон поиска: "A1:D100" (опционально — весь лист)',
      },
      matchCase: {
        type: "boolean",
        description: "Учитывать регистр (по умолчанию false)",
      },
      matchWholeWord: {
        type: "boolean",
        description: "Слово целиком (по умолчанию false)",
      },
    },
    required: ["action", "findWhat"],
  },
  riskLevel: "dangerous",
  requiresUndo: true,
  estimateCells: (args: Record<string, unknown>) => {
    const action = String(args.action ?? "");
    return action === "replace" || action === "replaceAll" ? 100 : 0;
  },

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const action = String(args.action ?? "") as FindReplaceAction;
    if (!action || !["find", "findAll", "replace", "replaceAll"].includes(action)) {
      return {
        ok: false,
        summary: "Некорректный action",
        error: {
          code: "INVALID_ACTION",
          message: 'action должен быть find/findAll/replace/replaceAll',
          retryable: false,
        },
      };
    }
    const findWhat = String(args.findWhat ?? "");
    if (!findWhat && action !== "find" && action !== "findAll") {
      // findWhat может быть пустым только при явном поиске пустых (зарезервировано)
      return {
        ok: false,
        summary: "Не указано findWhat",
        error: {
          code: "MISSING_FIND_WHAT",
          message: "findWhat обязателен",
          retryable: false,
        },
      };
    }
    const replaceWith =
      action === "replace" || action === "replaceAll"
        ? String(args.replaceWith ?? "")
        : undefined;
    if (
      (action === "replace" || action === "replaceAll") &&
      args.replaceWith === undefined
    ) {
      return {
        ok: false,
        summary: "Не указано replaceWith",
        error: {
          code: "MISSING_REPLACE_WITH",
          message: "replaceWith обязателен для replace/replaceAll",
          retryable: false,
        },
      };
    }

    const criteria: FindCriteria = {};
    if (args.matchCase === true) criteria.matchCase = true;
    if (args.matchWholeWord === true) criteria.matchWholeWord = true;
    criteria.searchDirection = "Forward";

    const rawAddress = typeof args.address === "string" ? args.address : undefined;
    const isMutation = action === "replace" || action === "replaceAll";

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Определяем диапазон поиска
      let range: Excel.Range;
      if (rawAddress) {
        range = getRangeSafe(context, rawAddress);
      } else {
        const used = sheet.getUsedRangeOrNullObject();
        used.load("isNullObject, address");
        await context.sync();
        if ((used as any).isNullObject) {
          return {
            ok: true,
            summary: `Лист пуст — нечего искать`,
            data: { action, matches: 0, addresses: [] },
          };
        }
        range = used as unknown as Excel.Range;
      }

      if (isMutation) {
        await undoManager.createBackup(rawAddress ?? "usedRange", "findAndReplace", {
          description: `${action}: "${findWhat}" → "${replaceWith}"`,
        });
      }

      // Find-only: вернуть адреса совпадений
      if (action === "find" || action === "findAll") {
        const addresses = await findAllMatches(
          context,
          range,
          findWhat,
          criteria,
        );
        const limit = action === "find" ? 1 : addresses.length;
        const result = addresses.slice(0, limit);
        return {
          ok: true,
          summary:
            result.length === 0
              ? `Совпадений "${findWhat}" не найдено`
              : `Найдено ${result.length} совпадений "${findWhat}"`,
          data: {
            action,
            findWhat,
            matches: result.length,
            addresses: result,
          },
        };
      }

      // Replace / ReplaceAll: найти → записать новые значения
      const addresses = await findAllMatches(
        context,
        range,
        findWhat,
        criteria,
      );
      const toReplace = action === "replace" ? addresses.slice(0, 1) : addresses;

      let replacedCount = 0;
      for (const addr of toReplace) {
        try {
          const cell = sheet.getRange(addr);
          cell.load("values");
          await context.sync();
          const currentVal = (cell as any).values?.[0]?.[0];
          // Точное или частичное замена
          let newVal: string;
          if (criteria.matchWholeWord && String(currentVal) !== findWhat) {
            continue;
          }
          if (String(currentVal) === findWhat) {
            newVal = replaceWith as string;
          } else {
            // Частичная замена внутри строки
            const flags = criteria.matchCase ? "g" : "gi";
            newVal = String(currentVal).replace(
              new RegExp(escapeRegExp(findWhat), flags),
              replaceWith as string,
            );
          }
          (cell as any).values = [[newVal]];
          replacedCount++;
        } catch {
          // Пропускаем ячейку, которую не удалось изменить
        }
      }
      await context.sync();

      return {
        ok: true,
        summary:
          replacedCount === 0
            ? `Нечего заменять (нет совпадений "${findWhat}")`
            : `${action}: заменено ${replacedCount} ячеек "${findWhat}" → "${replaceWith}"`,
        data: {
          action,
          findWhat,
          replaceWith,
          replaced: replacedCount,
          totalMatches: addresses.length,
        },
        cellsAffected: replacedCount,
      };
    });
  },
});

/** Экранирование спецсимволов RegExp для безопасной частичной замены. */
function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

toolRegistry.registerDefinition(findAndReplaceTool);
