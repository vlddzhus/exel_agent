import { createOpenAI } from "@ai-sdk/openai";
import { createAnthropic } from "@ai-sdk/anthropic";
import { LanguageModel, tool, jsonSchema } from "ai";
import type { ToolSet } from "ai";

export interface ProviderEntry {
  name: string;
  priority: number;
  model: LanguageModel;
  defaultModel: string;
  available: boolean;
}

/**
 * Имя модели по умолчанию для OpenAI-совместимого провайдера.
 * Приоритет: MODEL_NAME → OPENMODEL_MODEL → "claude-haiku-4-5".
 *
 * claude-haiku-4-5 — быстрый дешёвый аналог "flash", поддерживает tool-calling.
 * Идём через агрегатор https://vip.j3gb.com/v1 (OpenAI-совместимый протокол).
 * Примечание: DeepSeek на провайдере j3gb.com отсутствует (проверено /v1/models).
 */
const DEFAULT_FLASH_MODEL =
  process.env.MODEL_NAME ||
  process.env.OPENMODEL_MODEL ||
  "claude-haiku-4-5";

function init(): ProviderEntry[] {
  const entries: ProviderEntry[] = [];

  if (process.env.OPENMODEL_API_KEY) {
    const baseURL = (
      process.env.OPENMODEL_BASE_URL || "https://api.openmodel.ai/v1"
    ).replace(/\/+$/, "");
    const openmodel = createOpenAI({
      apiKey: process.env.OPENMODEL_API_KEY,
      baseURL,
    });
    const modelName = process.env.OPENMODEL_MODEL || DEFAULT_FLASH_MODEL;
    entries.push({
      name: "openmodel",
      priority: 0,
      // .chat() → эндпоинт /v1/chat/completions (Chat Completions API).
      // Обязательно для OpenAI-совместимых прокси (j3gb, Groq): новый
      // /v1/responses, который AI SDK v4 использует по умолчанию через
      // openai(model), провайдеры-агрегаторы часто не поддерживают.
      model: openmodel.chat(modelName),
      defaultModel: modelName,
      available: true,
    });
  }

  if (process.env.OPENAI_API_KEY) {
    const opts: { apiKey: string; baseURL?: string } = {
      apiKey: process.env.OPENAI_API_KEY,
    };
    if (process.env.OPENAI_BASE_URL) opts.baseURL = process.env.OPENAI_BASE_URL;
    const openai = createOpenAI(opts);
    const modelName = process.env.DEFAULT_MODEL || "gpt-4o";
    entries.push({
      name: "openai",
      priority: 5,
      model: openai.chat(modelName),
      defaultModel: modelName,
      available: true,
    });
  }

  if (process.env.GROQ_API_KEY) {
    const groq = createOpenAI({
      apiKey: process.env.GROQ_API_KEY,
      baseURL: "https://api.groq.com/openai/v1",
    });
    const modelName = process.env.GROQ_MODEL || "llama-3.3-70b-versatile";
    entries.push({
      name: "groq",
      priority: 0,
      model: groq.chat(modelName),
      defaultModel: modelName,
      available: true,
    });
  }

  if (process.env.ANTHROPIC_API_KEY) {
    const anthropic = createAnthropic({
      apiKey: process.env.ANTHROPIC_API_KEY,
    });
    const modelName = process.env.ANTHROPIC_MODEL || "claude-sonnet-4-20250514";
    entries.push({
      name: "anthropic",
      priority: 20,
      model: anthropic(modelName),
      defaultModel: modelName,
      available: true,
    });
  }

  entries.sort((a, b) => a.priority - b.priority);
  return entries;
}

const providers: ProviderEntry[] = init();

export function getModelChain(): string[] {
  return providers
    .filter((p) => p.available)
    .map((p) => `${p.name} (priority ${p.priority})`);
}

export function getProviderEntries(): ProviderEntry[] {
  return providers.filter((p) => p.available);
}

export function convertTools(toolDefs: unknown[]): ToolSet {
  const tools: ToolSet = {};
  if (!Array.isArray(toolDefs)) return tools;
  for (const t of toolDefs) {
    const td = t as {
      type: string;
      function: {
        name: string;
        description: string;
        parameters: Record<string, unknown>;
      };
    };
    if (td?.function?.name) {
      tools[td.function.name] = tool({
        description: td.function.description || "",
        inputSchema: jsonSchema(td.function.parameters || {}),
      });
    }
  }
  return tools;
}
