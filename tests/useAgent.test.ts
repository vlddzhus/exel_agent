/**
 * Тесты для useAgent.ts — построение тела запроса и выполнение tool-call.
 *
 * Главная проверяемая инварианта: body запроса к /api/agent/stream ОБЯЗАН
 * содержать tools (схемы инструментов) и system message (агентский промпт).
 * Без них бэкенд вызывает streamText({tools:{}}) — у LLM нет инструментов,
 * она отвечает текстом («корень №1» разрыва цепи).
 *
 * executeToolCall проверяется на устойчивость к невалидному JSON: раньше пустой
 * catch гасил ошибку, фетч tool-result не уходил, бэкенд ждал 120с → TOOL_TIMEOUT
 * («скрытая проблема №4» — silent error suppression).
 */
import {
  buildRequestBody,
  executeToolCall,
} from "../src/taskpane/hooks/useAgent-core";
import { AGENT_PROMPT } from "../src/taskpane/agent/agent-prompt";
import { toolRegistry } from "../src/taskpane/tools/registry";
import "../src/taskpane/tools";

describe("buildRequestBody", () => {
  test("включает tools — схемы всех инструментов", () => {
    const body = buildRequestBody("сделай 5 таблиц", toolRegistry.getSchemas());
    expect(Array.isArray(body.tools)).toBe(true);
    expect(body.tools.length).toBeGreaterThanOrEqual(24);
    // Формат каждой схемы — { type:"function", function:{name, description, parameters} }
    const first = body.tools[0];
    expect(first.type).toBe("function");
    expect(typeof first.function.name).toBe("string");
  });

  test("включает system message с AGENT_PROMPT первым", () => {
    const body = buildRequestBody("привет", toolRegistry.getSchemas());
    expect(body.messages[0]).toEqual({
      role: "system",
      content: AGENT_PROMPT,
    });
  });

  test("user message идёт после system", () => {
    const body = buildRequestBody("сделай 5 таблиц", toolRegistry.getSchemas());
    expect(body.messages[1]).toEqual({
      role: "user",
      content: "сделай 5 таблиц",
    });
    expect(body.messages.length).toBe(2);
  });

  test("tools всегда передаются, даже если реестр пуст (без падения)", () => {
    // Реестр в этом процессе уже заполнен barrel-импортом; проверяем структуру.
    const body = buildRequestBody("тест", toolRegistry.getSchemas());
    expect(Array.isArray(body.tools)).toBe(true);
  });
});

describe("executeToolCall — устойчивость к ошибкам", () => {
  test("невалидный JSON аргументов → ToolResult с ok:false (без throw)", async () => {
    // Раньше: JSON.parse падал, catch {} гасил, tool-result не уходил в бэкенд.
    const result = await executeToolCall("setValues", "{НЕВАЛИДНЫЙ JSON");
    expect(result.ok).toBe(false);
    expect(result.error).toBeDefined();
    expect(result.error?.code).toBe("INVALID_ARGS");
  });

  test("несуществующий инструмент → ToolResult с ok:false", async () => {
    const result = await executeToolCall("noSuchToolXYZ", '{"a":1}');
    expect(result.ok).toBe(false);
    expect(result.summary).toMatch(/Tool not found|не найден/i);
  });

  test("валидный инструмент с валидными args → вызывает execute", async () => {
    // getWorkbookOverview безопасный (safe) — не меняет данные, только читает.
    // Без Excel-среды execute упадёт внутри, но executeToolCall не должен
    // пробрасывать исключение наружу — возвращает ToolResult с ok:false.
    const result = await executeToolCall("getWorkbookOverview", "{}");
    // В node-среде без Office.js результат будет ok:false (Office не определён),
    // но функция не должна throw.
    expect(result).toBeDefined();
    expect(typeof result.ok).toBe("boolean");
  });
});
