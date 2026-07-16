/**
 * Тесты для agent-prompt.ts — агентский системный промпт.
 *
 * Промпт ДОЛЖЕН предписывать модели действовать через инструменты Excel,
 * а не генерировать таблицы текстом в чат. Это исправляет «корень №3»
 * разрыва цепи: старый CHAT_PROMPT буквально запрещал вызывать инструменты.
 *
 * См. docs/06-AGENT-LOOP.md §1 (Hybrid loop) и docs/03-TOOLS-SPEC.md.
 */
import { AGENT_PROMPT } from "../src/taskpane/agent/agent-prompt";

describe("AGENT_PROMPT", () => {
  test("экспортируется непустая строка", () => {
    expect(typeof AGENT_PROMPT).toBe("string");
    expect(AGENT_PROMPT.length).toBeGreaterThan(500);
  });

  test("назначает роль автономного Excel-агента", () => {
    expect(AGENT_PROMPT.toLowerCase()).toMatch(/excel|таблиц|agent|агент/);
  });

  test("предписывает использовать инструменты для действий с таблицей", () => {
    // Ключевое требование: при просьбе изменить таблицу — вызывать инструмент,
    // а не писать результат текстом в чат.
    expect(AGENT_PROMPT.toLowerCase()).toMatch(/инструмент|tool/);
  });

  test("упоминает критические инструменты по имени", () => {
    // Модель должна знать имена canonical-инструментов, чтобы вызывать их.
    for (const name of ["setValues", "appendRows", "getWorkbookOverview"]) {
      expect(AGENT_PROMPT).toContain(name);
    }
  });

  test("запрещает отказ от действий («нет доступа к файлу»)", () => {
    // Старая проблема: агент говорил «у меня нет зрения на файл».
    // Промпт должен явно опровергать этот паттерн.
    expect(AGENT_PROMPT.toLowerCase()).toMatch(/не (говори|отказывай|отказ)|never say.*(no access|нет доступ)/i);
  });

  test("отвечает на языке пользователя", () => {
    expect(AGENT_PROMPT.toLowerCase()).toMatch(/язык пользователя|language of the user|тот же язык|russian/);
  });
});
