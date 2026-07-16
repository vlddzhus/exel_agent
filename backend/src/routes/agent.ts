import { Router, Request, Response } from "express";
import { generateText, LanguageModel } from "ai";
import type { ToolSet } from "ai";
import { getProviderEntries, convertTools } from "../utils/provider-factory";
import { getFilteredProviderChain, TaskType } from "../utils/model-router";
import type { UserTier } from "../utils/model-router";
import { pruneMessages } from "../utils/prune";
import { redactSecrets } from "../utils/secrets-redactor";
import {
  registerRequest,
  registerToolCall,
  resolveToolCall,
  cleanupRequest,
} from "../utils/pending-tools";
import {
  TIER_LIMITS,
  getUsage,
  getUsageSummary,
  incrementUsage,
} from "../utils/usage-store";
import {
  isDevBypassEnabled,
  isUsageBlocked,
  resolveTier,
  resolveUserId,
  DEV_DAILY_LIMIT,
} from "../config/dev-bypass";
import pino from "pino";

const logger = pino({ level: process.env.LOG_LEVEL || "info" });
const router = Router();

// ── Helpers ──

function emit(res: Response, event: string, data: unknown): void {
  if (!res.destroyed) {
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  }
}

async function tryGenerate(
  system: string | undefined,
  messages: unknown[],
  tools: ToolSet | undefined,
  signal: AbortSignal,
) {
  const entries = getProviderEntries();
  const tried: string[] = [];
  for (const entry of entries) {
    tried.push(entry.name);
    try {
      return await generateText({
        model: entry.model as LanguageModel,
        system,
        messages: messages as any,
        tools,
        abortSignal: signal,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      logger.warn({ provider: entry.name, error: msg }, "provider failed");
      if (tried.length < entries.length) continue;
      throw err;
    }
  }
  throw new Error("All providers exhausted");
}

// ── GET /usage — current user usage stats ──

router.get("/usage", async (req: Request, res: Response) => {
  const userId = resolveUserId(req.userId);
  const tier = resolveTier(req.tier);
  const limit = isDevBypassEnabled() ? DEV_DAILY_LIMIT : (TIER_LIMITS[tier] ?? TIER_LIMITS.free);
  const usage = await getUsageSummary(userId);

  res.json({
    userId,
    tier,
    used: usage.count,
    limit,
    totalCost: usage.cost,
    resetDate: usage.resetDate,
    recentRequests: usage.recentRequests,
  });
});

// ── POST /complete — non-streaming completion ──

router.post("/complete", async (req: Request, res: Response) => {
  const startTime = Date.now();
  const { messages, tools: toolDefs, model } = req.body;
  const userId = req.userId || "anonymous";

  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    res.status(400).json({
      error: "messages array is required",
      code: "INVALID_MESSAGES",
    });
    return;
  }

  const tier = resolveTier(req.tier);
  const limit = isDevBypassEnabled() ? DEV_DAILY_LIMIT : (TIER_LIMITS[tier] ?? TIER_LIMITS.free);
  const usage = await getUsage(userId);
  if (isUsageBlocked(usage.count, limit)) {
    res.status(429).json({
      error: `Достигнут дневной лимит (${limit} запросов). Обновите тариф.`,
      code: "DAILY_LIMIT_REACHED",
    });
    return;
  }

  const pruned = pruneMessages(messages);

  try {
    const system =
      (pruned.find((m: { role: string }) => m.role === "system")
        ?.content as string) || undefined;
    const result = await tryGenerate(
      system,
      pruned.filter((m: { role: string }) => m.role !== "system"),
      convertTools(toolDefs),
      AbortSignal.timeout(120000),
    );

    const elapsed = Date.now() - startTime;
    logger.info(
      { userId, tokens: result.usage?.totalTokens, elapsed },
      "complete",
    );

    await incrementUsage(userId, {
      provider: "unknown",
      model: model || "unknown",
      inputTokens: result.usage.inputTokens,
      outputTokens: result.usage.outputTokens,
    });

    const response = {
      id: undefined,
      choices: [
        {
          index: 0,
          message: {
            role: "assistant",
            content: result.text || null,
            tool_calls: (result.toolCalls ?? []).map((tc) => ({
              id: tc.toolCallId,
              type: "function" as const,
              function: {
                name: tc.toolName,
                arguments: JSON.stringify(tc.input),
              },
            })),
          },
          finish_reason:
            result.finishReason === "tool-calls"
              ? "tool_calls"
              : (result.finishReason ?? "stop"),
        },
      ],
      usage: {
        total_tokens: result.usage.totalTokens,
        input_tokens: result.usage.inputTokens,
        output_tokens: result.usage.outputTokens,
      },
    };

    res.json(response);
  } catch (error: unknown) {
    const msg = error instanceof Error ? error.message : String(error);
    const safeMsg = redactSecrets(msg);
    logger.error({ userId, error: safeMsg }, "complete error");

    if (msg.includes("429") || msg.includes("rate limit")) {
      res.status(429).json({
        error: "Rate limit reached",
        code: "RATE_LIMIT",
        retryAfter: 30,
      });
      return;
    }
    if (
      msg.includes("timeout") ||
      msg.includes("timed out") ||
      msg.includes("ETIMEDOUT")
    ) {
      res.status(504).json({ error: "Request timed out", code: "TIMEOUT" });
      return;
    }
    res.status(502).json({
      error: "Upstream API error",
      code: "UPSTREAM_ERROR",
      message: safeMsg,
    });
  }
});

// ── POST /stream — SSE streaming endpoint ──

router.post("/stream", async (req: Request, res: Response) => {
  const { messages, tools: toolDefs, model } = req.body;
  const userId = req.userId || "anonymous";

  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    res
      .status(400)
      .json({ error: "messages array is required", code: "INVALID_MESSAGES" });
    return;
  }

  res.writeHead(200, {
    "Content-Type": "text/event-stream",
    "Cache-Control": "no-cache",
    Connection: "keep-alive",
    "X-Accel-Buffering": "no",
  });

  const abortController = new AbortController();

  // ── Heartbeat every 15s ──
  const heartbeatTimer = setInterval(() => {
    if (res.destroyed) {
      clearInterval(heartbeatTimer);
      return;
    }
    try {
      res.write("event: heartbeat\ndata: {}\n\n");
    } catch {
      clearInterval(heartbeatTimer);
    }
  }, 15000);

  // ── Daily limit check (обходится в dev через isUsageBlocked) ──
  const usage = await getUsage(userId);
  const tier = resolveTier(req.tier);
  const limit = isDevBypassEnabled() ? DEV_DAILY_LIMIT : (TIER_LIMITS[tier] ?? TIER_LIMITS.free);
  if (isUsageBlocked(usage.count, limit)) {
    emit(res, "error", {
      code: "DAILY_LIMIT_REACHED",
      message: `Достигнут дневной лимит (${limit} запросов). Обновите тариф.`,
    });
    clearInterval(heartbeatTimer);
    res.end();
    return;
  }

  const pruned = pruneMessages(messages);
  const systemMessage =
    (pruned.find((m: { role: string }) => m.role === "system")
      ?.content as string) || undefined;
  const toolSet = convertTools(toolDefs);

  // Наблюдаемость tool-calling-цепи (см. отчёт по разрыву «Мозг↔Руки»).
  // Краткая инфо-строка: дошли ли инструменты до generateText и сколько их.
  // Подробный дамп схемы — на уровне debug, чтобы не шуметь в проде.
  const toolCount = Array.isArray(toolDefs) ? toolDefs.length : 0;
  logger.info(
    {
      toolsReceived: toolCount,
      toolsConverted: Object.keys(toolSet).length,
      hasSystemPrompt: !!systemMessage,
      systemPromptLength: systemMessage ? systemMessage.length : 0,
    },
    "stream request parsed",
  );
  if (toolCount === 0) {
    logger.warn(
      "stream request received 0 tools — модель не сможет вызывать инструменты",
    );
  }
  logger.debug(
    {
      convertedToolNames: Object.keys(toolSet),
      sampleIncomingSchema:
        Array.isArray(toolDefs) && toolDefs[0]
          ? JSON.stringify(toolDefs[0]).slice(0, 300)
          : null,
    },
    "stream tools detail",
  );
  const requestId = `${userId}_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
  let chatHistory = pruned.filter((m: { role: string }) => m.role !== "system");
  let selectedProvider = "unknown";

  registerRequest(requestId);

  emit(res, "status", { phase: "connecting", provider: "", requestId });

  req.on("close", () => {
    abortController.abort();
    cleanupRequest(requestId);
  });

  for (let step = 0; step < 15; step++) {
    if (res.destroyed) break;

    const entries = getFilteredProviderChain(
      resolveTier(req.tier) as UserTier,
      TaskType.COMPLEX,
    );
    const tried: string[] = [];
    // NOTE: используется generateText (не-стрим), а не streamText.
    // Причина: OpenAI-совместимый прокси провайдера зависает/обрывает соединение
    // (ECONNRESET) при стриминге tool_calls — функция tool-call генерируется
    // побайтно и keep-alive обрывается на длинных стримах. Не-стрим режим с tools
    // работает стабильно и быстро (2-4с на шаг). См. диагностику в отчёте.
    // Для пользователя разница: "thinking" приходит одним блоком после готовности
    // ответа, а не побуквенно — приемлемая цена за надёжность.
    let started = false;
    let streamContent = "";
    let streamToolCalls: Array<{
      id: string;
      type: "function";
      function: { name: string; arguments: string };
    }> = [];
    let streamUsage: {
      input_tokens: number;
      output_tokens: number;
      total_tokens: number;
    } | null = null;
    let success = false;

    for (const entry of entries) {
      if (started) break;
      tried.push(entry.name);

      try {
        emit(res, "status", {
          phase: "connecting",
          provider: entry.name,
        });

        const result = await generateText({
          model: entry.model as LanguageModel,
          system: systemMessage,
          messages: chatHistory as any,
          tools: toolSet,
          abortSignal: abortController.signal,
        });

        started = true;
        selectedProvider = entry.name;

        // Наблюдаемость: что вернула модель. finishReason="tool-calls" +
        // непустой toolCalls = цепь работает. "stop" без toolCalls при
        // tools>0 = модель проигнорировала инструменты (проверить prompt).
        logger.info(
          {
            provider: entry.name,
            finishReason: result.finishReason,
            toolCallsCount: result.toolCalls ? result.toolCalls.length : 0,
            toolCallsNames: result.toolCalls
              ? result.toolCalls.map((tc) => tc.toolName)
              : [],
            textLength: (result.text || "").length,
          },
          "generateText result",
        );
        logger.debug(
          { textPreview: (result.text || "").slice(0, 300) },
          "generateText text",
        );

        // Текст ответа — эмитим как "thinking" одним блоком (не побуквенно).
        streamContent = result.text || "";
        if (streamContent) {
          emit(res, "thinking", { text: streamContent });
        }

        // Tool calls из не-стрим результата.
        streamToolCalls = (result.toolCalls ?? []).map((tc) => ({
          id: tc.toolCallId,
          type: "function" as const,
          function: {
            name: tc.toolName,
            arguments: JSON.stringify(tc.input),
          },
        }));

        // Usage.
        streamUsage = {
          input_tokens: result.usage?.inputTokens ?? 0,
          output_tokens: result.usage?.outputTokens ?? 0,
          total_tokens: result.usage?.totalTokens ?? 0,
        };

        success = true;
        break;
      } catch (error: unknown) {
        if (res.destroyed) break;
        const msg = error instanceof Error ? error.message : String(error);
        const safeMsg = redactSecrets(msg);
        logger.warn(
          { provider: entry.name, error: safeMsg },
          "stream provider failed",
        );

        if (tried.length < entries.length) {
          emit(res, "model_fallback", {
            from: entry.name,
            to: entries[tried.length]?.name,
            reason: "rate_limit_or_error",
          });
          continue;
        }

        if (started) {
          if (!res.destroyed) {
            emit(res, "error", {
              code: "PROVIDER_FAILED_MID_STREAM",
              message: "Сейчас перегружено, попробуйте ещё раз",
            });
            clearInterval(heartbeatTimer);
            res.end();
          }
          cleanupRequest(requestId);
          return;
        }
      }
    }

    if (!success || !streamUsage || res.destroyed) {
      if (!res.destroyed && !started) {
        emit(res, "error", {
          code: "ALL_PROVIDERS_EXHAUSTED",
          message: "Все провайдеры недоступны",
        });
        clearInterval(heartbeatTimer);
        res.end();
      }
      cleanupRequest(requestId);
      return;
    }

    // ── No tool calls → done ──
    if (streamToolCalls.length === 0) {
      await incrementUsage(userId, {
        provider: selectedProvider,
        model: model || "unknown",
        inputTokens: streamUsage.input_tokens,
        outputTokens: streamUsage.output_tokens,
      });

      logger.info({ userId, tokens: streamUsage.total_tokens }, "stream done");

      if (!res.destroyed) {
        emit(res, "done", {
          choices: [
            {
              index: 0,
              message: {
                role: "assistant",
                content: streamContent || null,
              },
              finish_reason: "stop",
            },
          ],
          usage: streamUsage,
        });
        clearInterval(heartbeatTimer);
        res.end();
      }
      cleanupRequest(requestId);
      return;
    }

    // ── Emit tool_calls ──
    for (const tc of streamToolCalls) {
      emit(res, "tool_call", tc);
    }

    // ── Add assistant message to history (AI SDK v7 CoreMessage format) ──
    const assistantContent: Array<{ type: "text"; text: string } | { type: "tool-call"; toolCallId: string; toolName: string; input: unknown }> = [];
    if (streamContent) {
      assistantContent.push({ type: "text", text: streamContent });
    }
    for (const tc of streamToolCalls) {
      assistantContent.push({
        type: "tool-call",
        toolCallId: tc.id,
        toolName: tc.function.name,
        input: JSON.parse(tc.function.arguments),
      });
    }
    chatHistory.push({
      role: "assistant" as const,
      content: assistantContent,
    });

    // ── Wait for tool results ──
    const results = await Promise.all(
      streamToolCalls.map(async (tc) => {
        try {
          const result = await registerToolCall(
            requestId,
            tc.id,
            tc.function.name,
            tc.function.arguments,
            120000,
          );
          return { id: tc.id, result };
        } catch (err: unknown) {
          const msg = err instanceof Error ? err.message : String(err);
          return { id: tc.id, error: msg };
        }
      }),
    );

    // ── Check for errors ──
    let hasError = false;
    for (const r of results) {
      if ("error" in r) {
        hasError = true;
        if (r.error === "LOOP_DETECTED") {
          emit(res, "error", {
            code: "LOOP_DETECTED",
            message: "Обнаружено зацикливание. Попробуйте иначе.",
          });
        } else if (r.error === "TOOL_TIMEOUT") {
          emit(res, "error", {
            code: "TOOL_TIMEOUT",
            message: "Клиент не вернул результат вовремя.",
          });
        }
        break;
      }
    }

    if (hasError) {
      clearInterval(heartbeatTimer);
      res.end();
      cleanupRequest(requestId);
      return;
    }

    // ── Add tool results to history and continue loop (AI SDK v7 CoreMessage format) ──
    for (const r of results) {
      const safeResult = "result" in r ? r.result : { error: r.error };
      const tcDef = streamToolCalls.find((tc) => tc.id === r.id);
      chatHistory.push({
        role: "tool" as const,
        content: [
          {
            type: "tool-result",
            toolCallId: r.id,
            toolName: tcDef?.function.name || "unknown",
            output: { type: "json", value: safeResult },
          },
        ],
      });
    }
  }

  // ── Max steps reached ──
  if (!res.destroyed) {
    emit(res, "error", {
      code: "MAX_STEPS",
      message: "Шагов больше 15 — задача сложная, упростите",
    });
    clearInterval(heartbeatTimer);
    res.end();
  }
  cleanupRequest(requestId);
});

// ── POST /tool-result — receive tool execution result from client ──

router.post("/tool-result", async (req: Request, res: Response) => {
  const { requestId, toolCallId, result } = req.body;

  if (!requestId || !toolCallId) {
    res.status(400).json({
      error: "requestId and toolCallId are required",
      code: "INVALID_TOOL_RESULT",
    });
    return;
  }

  const ok = resolveToolCall(requestId, toolCallId, result);
  if (ok) {
    res.json({ status: "ok" });
  } else {
    res.status(404).json({
      error: "Tool call not found or already resolved",
      code: "TOOL_NOT_FOUND",
    });
  }
});

// ── POST /cancel (legacy — kept for Wiro compatibility) ──

router.post("/cancel", async (_req: Request, res: Response) => {
  res.json({ status: "cancelled" });
});

export default router;
