/**
 * useAgent-core.ts — чистая (non-React) логика агентского цикла на клиенте.
 *
 * Вынесена из useAgent.ts, чтобы быть тестируемой в node-среде без jsdom.
 * useAgent.ts — тонкая React-обёртка (state, fetch, SSE-парсинг).
 *
 * Главная инварианта (контракт с бэкендом POST /api/agent/stream):
 * тело запроса ОБЯЗАНО содержать tools (схемы инструментов) и system message
 * (агентский промпт). Без них у LLM физически нет инструментов для вызова.
 */

import { toolRegistry, type ToolResult } from "../tools/registry";
import { AGENT_PROMPT } from "../agent/agent-prompt";

// Максимальное время выполнения одного инструмента (мс).
// Office.js может зависнуть при неинициализированном соединении.
const TOOL_EXECUTION_TIMEOUT = 15_000;

// Автоматический повтор при ошибке «Excel в режиме редактирования».
const EDIT_MODE_RETRIES = 5;
const EDIT_MODE_RETRY_DELAY_MS = 1_500;

// ──────────────────────────────────────────────────────────────────────────
// Тип схемы инструмента в OpenAI function-calling формате.
// Совпадает с toolRegistry.getSchemas() (см. registry.ts:174).
// ──────────────────────────────────────────────────────────────────────────

export interface ToolFunctionSchema {
  type: "function";
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
}

export interface StreamRequestBody {
  messages: Array<{ role: "system" | "user"; content: string }>;
  tools: ToolFunctionSchema[];
}

/**
 * Строит тело запроса для POST /api/agent/stream.
 *
 * ВАЖНО: порядок messages — system первым (бэкенд agent.ts:250 извлекает
 * system-сообщение из массива по role==='system'), затем user.
 * tools — массив схем из toolRegistry.getSchemas().
 */
export function buildRequestBody(
  userText: string,
  tools: ToolFunctionSchema[],
): StreamRequestBody {
  return {
    messages: [
      { role: "system", content: AGENT_PROMPT },
      { role: "user", content: userText },
    ],
    tools,
  };
}

/**
 * Безопасный парсинг аргументов tool-call.
 * @returns parsed args или null при невалидном JSON.
 */
export function parseToolArgs(rawArgs: string): Record<string, unknown> | null {
  try {
    const parsed = JSON.parse(rawArgs);
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
      return parsed as Record<string, unknown>;
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * Выполняет tool-call локально (через Office.js внутри toolRegistry).
 *
 * Устойчивость к ошибкам (исправление «скрытой проблемы №4» — silent catch):
 * - если args — невалидный JSON → ToolResult с ok:false, код INVALID_ARGS.
 * - если инструмент не найден → ToolResult с ok:false.
 * - если execute бросает (нет Excel-среды и т.п.) → ToolResult с ok:false,
 *   код EXECUTE_FAILED. НИКОГДА не пробрасывает исключение наружу:
 *   в цикле useAgent это гарантировало бы уход tool-result, 120с таймаута
 *   на бэкенде и бесполезное сообщение «Клиент не вернул результат вовремя».
 *
 * Динамический import инструмента через глобальный toolRegistry — чтобы ядро
 * оставалось platform-agnostic и тестируемым в node.
 */
export async function executeToolCall(
  name: string,
  rawArgs: string,
): Promise<ToolResult> {
  const def = toolRegistry.getDefinition(name);
  const tool = toolRegistry.getTool(name);

  if (!def && !tool) {
    return {
      ok: false,
      summary: `Tool not found: ${name}`,
      error: {
        code: "TOOL_NOT_FOUND",
        message: `Инструмент «${name}» не зарегистрирован`,
        retryable: false,
      },
    };
  }

  const args = parseToolArgs(rawArgs);
  if (args === null) {
    return {
      ok: false,
      summary: `Невалидные аргументы для ${name}: ${rawArgs}`,
      error: {
        code: "INVALID_ARGS",
        message: `Не удалось разобрать JSON аргументов: ${rawArgs}`,
        retryable: false,
      },
    };
  }

  // Флаг «Excel в режиме редактирования» — агент подождёт и повторит.
  // Office.js блокирует все API-вызовы, пока пользователь не выйдет из ячейки.
  // Без retry пользователю приходилось бы отправлять команду заново.
  const isEditModeError = (result: ToolResult): boolean =>
    !result.ok &&
    result.error?.code === "EXECUTE_FAILED" &&
    result.error?.retryable &&
    /edit|редактирован|введите|enter/i.test(result.error?.message ?? "");

  for (let attempt = 0; attempt <= EDIT_MODE_RETRIES; attempt++) {
    try {
      if (def) {
        if (typeof Excel === "undefined") {
          return {
            ok: false,
            summary: `Office.js не инициализирован — ${name} не может быть выполнен`,
            error: {
              code: "OFFICE_NOT_READY",
              message: "Excel JavaScript API не доступен. Перезапустите надстройку.",
              retryable: true,
            },
          };
        }
        const result = await Promise.race([
          def.execute(args),
          new Promise<never>((_, reject) =>
            setTimeout(
              () => reject(new Error(`Tool execution timed out after ${TOOL_EXECUTION_TIMEOUT}ms`)),
              TOOL_EXECUTION_TIMEOUT,
            ),
          ),
        ]);
        return result;
      }
      const raw = await tool!.fn(args);
      try {
        return JSON.parse(raw) as ToolResult;
      } catch {
        return { ok: true, summary: raw };
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      const result: ToolResult = {
        ok: false,
        summary: `Ошибка выполнения ${name}: ${msg}`,
        error: {
          code: "EXECUTE_FAILED",
          message: msg,
          retryable: true,
        },
      };
      if (isEditModeError(result) && attempt < EDIT_MODE_RETRIES) {
        // Ждём, пока пользователь выйдет из режима редактирования
        await new Promise((r) => setTimeout(r, EDIT_MODE_RETRY_DELAY_MS));
        continue;
      }
      return result;
    }
  }

  // Fallback (недостижимо, но TypeScript требует return)
  return {
    ok: false,
    summary: `Ошибка выполнения ${name}`,
    error: { code: "EXECUTE_FAILED", message: "Все попытки исчерпаны", retryable: true },
  };
}

/**
 * Извлекает имя инструмента и args из SSE tool_call-события.
 * Бэкенд шлёт формат OpenAI: { id, type:"function", function:{name, arguments} }.
 * Совместимо и с Anthropic-стилем { toolCallId, toolName, input } (запасной путь).
 */
export function extractToolCallInfo(data: {
  function?: { name?: string; arguments?: string };
  name?: string;
  arguments?: string;
  id?: string;
  toolCallId?: string;
}): { name: string; args: string; id: string } {
  const name = data.function?.name || data.name || "";
  const args = data.function?.arguments || data.arguments || "{}";
  const id = data.id || data.toolCallId || "";
  return { name, args, id };
}

/**
 * Извлекает адрес диапазона из args — для подсветки ячеек в Excel.
 */
export function extractRangeAddress(args: Record<string, unknown>): string {
  const addr = args.range || args.address || args.target || args.sourceCell;
  return typeof addr === "string" ? addr : "";
}
