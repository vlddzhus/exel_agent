/**
 * Тест barrel-регистрации: tools/index.ts должен загружать все модули инструментов,
 * чтобы их top-level `toolRegistry.registerDefinition(...)` / `.register(...)`
 * срабатывали. Без этого в React-надстройке реестр остаётся ПУСТЫМ, и useAgent
 * отправляет в бэкенд пустой массив tools → модель физически не имеет инструментов.
 *
 * См. диагностику «корень №2» (barrel-регистрация) в отчёте по гибридному циклу.
 */
import { toolRegistry } from "../../src/taskpane/tools/registry";

// Импорт barrel запускает top-level side-effect регистрацию всех инструментов.
import "../../src/taskpane/tools";

describe("tools/index barrel — side-effect регистрация", () => {
  test("реестр содержит не менее 24 инструментов", () => {
    // docs/03-TOOLS-SPEC.md описывает 24 базовых инструмента; фактически
    // в реестре больше (legacy formula/table/knowledge-инструменты).
    // Проверяем нижнюю границу — если она не пройдена, регистрация сломана.
    const count = toolRegistry.getToolNames().length;
    expect(count).toBeGreaterThanOrEqual(24);
  });

  test("присутствуют критические write-инструменты", () => {
    const names = new Set(toolRegistry.getToolNames());
    // Инструменты, без которых задача «сделай 5 таблиц с названиями кофе»
    // физически не может быть выполнена агентом.
    for (const required of ["setValues", "appendRows", "setFormula"]) {
      expect(names.has(required)).toBe(true);
    }
  });

  test("присутствуют read-инструменты для «осмотра» листа", () => {
    const names = new Set(toolRegistry.getToolNames());
    for (const required of ["getWorkbookOverview", "getRange"]) {
      expect(names.has(required)).toBe(true);
    }
  });

  test("getSchemas возвращает валидные function-схемы для отправки в LLM", () => {
    const schemas = toolRegistry.getSchemas();
    expect(schemas.length).toBeGreaterThanOrEqual(24);
    for (const s of schemas) {
      expect(s.type).toBe("function");
      expect(typeof s.function.name).toBe("string");
      expect(s.function.name.length).toBeGreaterThan(0);
      expect(typeof s.function.description).toBe("string");
    }
  });
});
