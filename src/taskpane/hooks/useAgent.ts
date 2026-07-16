import { useCallback, useRef } from "react";
import { useChatStore } from "../stores/chatStore";
import { useLiveActivityStore } from "../stores/liveActivityStore";
import { useStreaming } from "./useStreaming";
import {
  buildRequestBody,
  executeToolCall,
  extractToolCallInfo,
  extractRangeAddress,
  parseToolArgs,
} from "./useAgent-core";
import { toolRegistry } from "../tools/registry";
import { highlightRange } from "../utils/cell-highlight";

// ──────────────────────────────────────────────────────────────────────────
// Backend URL.
// Локальный dev: https://localhost:4000 (см. backend/.env PORT, USE_HTTPS=true).
// ВАЖНО: протокол HTTPS обязателен — taskpane грузится с https://localhost:3000,
// и HTTP-fetch заблокируется как mixed content (NETWORK_ERROR: Failed to fetch).
// Можно переопределить через localStorage для отладки.
//
// Нормализуем сохранённое значение: в старых сборках в localStorage могло
// остаться http://localhost:4000 — оно даст мгновенный Failed to fetch.
// Если там http:// на localhost — перезаписываем на https:// один раз.
// ──────────────────────────────────────────────────────────────────────────
const BACKEND_URL = (() => {
  const stored = localStorage.getItem("backend_url");
  if (stored && /^http:\/\/localhost/i.test(stored)) {
    const fixed = stored.replace(/^http:/i, "https:");
    localStorage.setItem("backend_url", fixed);
    console.info("[useAgent] normalized backend_url http→https:", fixed);
    return fixed;
  }
  return stored || "https://localhost:4000";
})();

/**
 * Сколько мс ждать появления requestId, если status-событие ещё не пришло
 * (теоретическая перестановка порядка SSE-событий при буферизации).
 * Бэкенд всегда шлёт status{requestId} ПЕРЕД tool_call, поэтому обычно
 * requestId уже известен; этот guard — защита от редких race-condition.
 */
const REQUEST_ID_WAIT_MS = 3000;

export function useAgent() {
  const abortRef = useRef<AbortController | null>(null);
  const requestIdRef = useRef<string>("");
  const { handleEvent } = useStreaming();
  const addUserMessage = useChatStore((s) => s.addUserMessage);
  const setProcessing = useChatStore((s) => s.setProcessing);
  const start = useLiveActivityStore((s) => s.start);
  const reset = useLiveActivityStore((s) => s.reset);
  const setChanges = useLiveActivityStore((s) => s.setChanges);

  /**
   * Выполняет tool-call локально (через Office.js) и отправляет результат
   * бэкенду. Изолирован от SSE-reader-цикла: вызывается через `void`
   * (fire-and-forget), чтобы медленный Excel не блокировал чтение стрима.
   *
   * Устойчивость:
   *  - executeToolCall НИКОГДА не бросает (возвращает ToolResult с ok:false).
   *  - requestId может прийти позже status — ждём до REQUEST_ID_WAIT_MS.
   *  - POST к бэкенду падает мягко (catch с логированием, без swallow).
   * Без этого бэкенд ждёт 120с и эмитит TOOL_TIMEOUT.
   */
  const runToolAndReport = useCallback(
    async (
      name: string,
      args: string,
      id: string,
      token: string,
    ): Promise<void> => {
      const result = await executeToolCall(name, args);
      if (result.ok) {
        setChanges();
      }

      // Подсветка затронутого диапазона (не блокирует отправку результата).
      const parsed = parseToolArgs(args);
      if (parsed) {
        const addr = extractRangeAddress(parsed);
        if (addr) {
          highlightRange(addr).catch((e) =>
            console.warn(`[useAgent] highlightRange failed for ${addr}:`, e),
          );
        }
      }

      // Гарантированно отправляем tool-result бэкенду — даже при ошибке
      // выполнения. Если requestId ещё не захвачен (status не пришёл),
      // коротко ждём — иначе бэкенд повиснет на 120с.
      if (!id) {
        console.error(
          `[useAgent] tool_call без id — не могу отправить результат для ${name}`,
        );
        return;
      }
      if (!requestIdRef.current) {
        const deadline = Date.now() + REQUEST_ID_WAIT_MS;
        while (!requestIdRef.current && Date.now() < deadline) {
          await new Promise((r) => setTimeout(r, 50));
        }
      }
      if (!requestIdRef.current) {
        console.error(
          `[useAgent] requestId так и не пришёл за ${REQUEST_ID_WAIT_MS}мс — ` +
            `результат для ${name} (${id}) не отправлен, бэкенд даст TOOL_TIMEOUT`,
        );
        return;
      }

      fetch(`${BACKEND_URL}/api/agent/tool-result`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${token}`,
        },
        body: JSON.stringify({
          requestId: requestIdRef.current,
          toolCallId: id,
          result,
        }),
      }).catch((e) => {
        console.error(
          `[useAgent] tool-result POST failed for ${name} (${id}):`,
          e,
        );
      });
    },
    [setChanges],
  );

  const sendMessage = useCallback(
    async (text: string) => {
      addUserMessage(text);
      setProcessing(true);
      reset();
      start();

      const controller = new AbortController();
      abortRef.current = controller;

      // Схемы инструментов — берём из реестра ОДИН раз за запрос.
      // Реестр заполнен через side-effect import "../tools" в index.tsx.
      const tools = toolRegistry.getSchemas();
      const body = buildRequestBody(text, tools);

      const token = localStorage.getItem("auth_token") || "";

      try {
        const res = await fetch(`${BACKEND_URL}/api/agent/stream`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${token}`,
          },
          body: JSON.stringify(body),
          signal: controller.signal,
        });

        if (!res.ok) {
          handleEvent("error", {
            code: "HTTP_ERROR",
            message: `HTTP ${res.status}`,
          });
          return;
        }

        const reader = res.body?.getReader();
        if (!reader) {
          handleEvent("error", {
            code: "NO_RESPONSE_BODY",
            message: "Сервер вернул пустой ответ",
          });
          return;
        }

        const decoder = new TextDecoder();
        let buffer = "";
        let currentEvent = "";

        while (true) {
          const { done, value } = await reader.read();
          if (done) break;

          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split("\n");
          buffer = lines.pop() || "";

          for (const line of lines) {
            if (line.startsWith("event: ")) {
              currentEvent = line.slice(7).trim();
            } else if (line.startsWith("data: ")) {
              const payload = line.slice(6);
              let data: unknown;
              try {
                data = JSON.parse(payload);
              } catch (e) {
                // Раньше: пустой catch {} молча глотал невалидный JSON из стрима.
                // Теперь: логируем, чтобы SSE-протокол можно было отлаживать.
                console.warn("[useAgent] malformed SSE data, skipping:", payload, e);
                continue;
              }
              const evt = data as Record<string, unknown>;

              if (currentEvent === "status" && evt.requestId) {
                requestIdRef.current = String(evt.requestId);
                handleEvent("status", evt);
              } else if (currentEvent === "tool_call") {
                handleEvent("tool_call", evt);

                // Выполняем инструмент локально (Office.js) и шлём результат
                // бэкенду. НЕ блокируем SSE-reader: если Office.js медленный,
                // стрим должен продолжать читаться (heartbeat, следующие tool_call).
                // Запускаем как самостоятельный промис, ошибки — внутрь логирования.
                const { name, args, id } = extractToolCallInfo(
                  evt as Parameters<typeof extractToolCallInfo>[0],
                );
                void runToolAndReport(name, args, id, token);
              } else {
                handleEvent(currentEvent, evt);
              }
            }
          }
        }
      } catch (err: unknown) {
        const e = err as { name?: string; message?: string };
        if (e.name !== "AbortError") {
          handleEvent("error", {
            code: "NETWORK_ERROR",
            message: e.message || "Сетевая ошибка",
          });
        }
      } finally {
        setProcessing(false);
        abortRef.current = null;
      }
    },
    [addUserMessage, setProcessing, start, reset, handleEvent, runToolAndReport],
  );

  const cancel = useCallback(() => {
    abortRef.current?.abort();
    useLiveActivityStore.getState().cancel();
    setProcessing(false);
  }, [setProcessing]);

  return { sendMessage, cancel };
}
