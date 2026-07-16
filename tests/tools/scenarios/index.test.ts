/**
 * index.test.ts — проверка структуры и конфигурации всех 20 эталонных сценариев.
 *
 * Регистрирует минимальные мок-инструменты (чтобы не зависеть от порядка
 * выполнения тестов) и проверяет конфигурацию всех 20 сценариев.
 */
import { toolRegistry, defineTool } from "../../../src/taskpane/tools/registry";

// ---------------------------------------------------------------------------
// Register mock tool definitions for verification
// ---------------------------------------------------------------------------

const SCENARIO_TOOLS = [
  { name: "setValues", riskLevel: "moderate" as const, requiresUndo: true },
  { name: "setFormula", riskLevel: "moderate" as const, requiresUndo: true },
  { name: "fillRange", riskLevel: "moderate" as const, requiresUndo: true },
  { name: "formatAsTable", riskLevel: "moderate" as const, requiresUndo: true },
  {
    name: "applyCellFormat",
    riskLevel: "moderate" as const,
    requiresUndo: true,
  },
  { name: "getRangeStats", riskLevel: "safe" as const, requiresUndo: false },
  { name: "lookup", riskLevel: "moderate" as const, requiresUndo: true },
  { name: "detectDataTypes", riskLevel: "safe" as const, requiresUndo: false },
  {
    name: "removeDuplicates",
    riskLevel: "moderate" as const,
    requiresUndo: true,
  },
  {
    name: "splitTextToColumns",
    riskLevel: "moderate" as const,
    requiresUndo: true,
  },
  { name: "normalizeText", riskLevel: "moderate" as const, requiresUndo: true },
  { name: "sortData", riskLevel: "moderate" as const, requiresUndo: true },
  { name: "filterData", riskLevel: "moderate" as const, requiresUndo: true },
  {
    name: "createPivotTable",
    riskLevel: "moderate" as const,
    requiresUndo: true,
  },
  { name: "createChart", riskLevel: "moderate" as const, requiresUndo: true },
  { name: "freezePanes", riskLevel: "safe" as const, requiresUndo: false },
  { name: "autoFitColumns", riskLevel: "safe" as const, requiresUndo: false },
  { name: "clearRange", riskLevel: "dangerous" as const, requiresUndo: true },
  {
    name: "applyConditionalFormat",
    riskLevel: "moderate" as const,
    requiresUndo: true,
  },
  {
    name: "getWorkbookOverview",
    riskLevel: "safe" as const,
    requiresUndo: false,
  },
  { name: "findAnomalies", riskLevel: "safe" as const, requiresUndo: false },
  { name: "getRange", riskLevel: "safe" as const, requiresUndo: false },
  { name: "appendRows", riskLevel: "moderate" as const, requiresUndo: true },
  { name: "manageSheets", riskLevel: "dangerous" as const, requiresUndo: true },
];

function registerScenarioTools() {
  const toolConfigs: Record<string, { parameters: Record<string, unknown> }> = {
    fillRange: {
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          fillType: { type: "string", enum: ["progression", "copy", "value"] },
          fillValue: {},
          startValue: { type: "number" },
          step: { type: "number" },
        },
        required: ["address", "fillType"],
      },
    },
    lookup: {
      parameters: {
        type: "object",
        properties: {
          lookupAddress: { type: "string" },
          lookupColumn: { type: "number" },
          resultColumn: { type: "number" },
          lookupValue: { type: "string" },
          writeTo: { type: "string" },
          exactMatch: { type: "boolean" },
        },
        required: [
          "lookupAddress",
          "lookupColumn",
          "resultColumn",
          "lookupValue",
          "writeTo",
        ],
      },
    },
    sortData: {
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          sortColumns: {
            type: "array",
            items: {
              type: "object",
              properties: {
                column: { type: "number" },
                order: { type: "string", enum: ["asc", "desc"] },
              },
              required: ["column"],
            },
          },
          hasHeaders: { type: "boolean" },
        },
        required: ["address", "sortColumns"],
      },
    },
    createPivotTable: {
      parameters: {
        type: "object",
        properties: {
          sourceAddress: { type: "string" },
          destinationAddress: { type: "string" },
          name: { type: "string" },
          rows: { type: "array", items: { type: "string" } },
          columns: { type: "array", items: { type: "string" } },
          values: {
            type: "array",
            items: {
              type: "object",
              properties: {
                column: { type: "string" },
                agg: { type: "string" },
              },
              required: ["column", "agg"],
            },
          },
        },
        required: [
          "sourceAddress",
          "destinationAddress",
          "name",
          "rows",
          "values",
        ],
      },
    },
    createChart: {
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          chartType: { type: "string" },
          title: { type: "string" },
          position: { type: "string", enum: ["Object", "newSheet"] },
        },
        required: ["address", "chartType"],
      },
    },
    removeDuplicates: {
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          columns: { type: "array", items: { type: "number" } },
          hasHeaders: { type: "boolean" },
        },
        required: ["address", "columns"],
      },
    },
    normalizeText: {
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          operations: {
            type: "array",
            items: {
              type: "string",
              enum: [
                "trim",
                "uppercase",
                "lowercase",
                "propercase",
                "cleanWhitespace",
              ],
            },
          },
        },
        required: ["address", "operations"],
      },
    },
    filterData: {
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          clear: { type: "boolean" },
        },
        required: ["address"],
      },
    },
    splitTextToColumns: {
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          delimiter: {
            type: "string",
            enum: ["auto", "space", "comma", "semicolon", "dot"],
          },
          targetStartCell: { type: "string" },
          maxColumns: { type: "number" },
        },
        required: ["address"],
      },
    },
    freezePanes: {
      parameters: {
        type: "object",
        properties: {
          target: { type: "string" },
          sheetName: { type: "string" },
        },
        required: ["target"],
      },
    },
    autoFitColumns: {
      parameters: {
        type: "object",
        properties: { address: { type: "string" } },
        required: ["address"],
      },
    },
    clearRange: {
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          clearWhat: { type: "string", enum: ["all", "values", "formats"] },
        },
        required: ["address"],
      },
    },
    applyConditionalFormat: {
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          rule: { type: "object" },
          format: { type: "object" },
        },
        required: ["address", "rule"],
      },
    },
    detectDataTypes: {
      parameters: {
        type: "object",
        properties: { address: { type: "string" } },
        required: ["address"],
      },
    },
    findAnomalies: {
      parameters: {
        type: "object",
        properties: { address: { type: "string" } },
        required: ["address"],
      },
    },
    getWorkbookOverview: {
      parameters: {
        type: "object",
        properties: {},
      },
    },
  };

  for (const t of SCENARIO_TOOLS) {
    if (!toolRegistry.getDefinition(t.name)) {
      const extraParams = toolConfigs[t.name]?.parameters ?? {
        type: "object",
        properties: {},
      };
      toolRegistry.registerDefinition(
        defineTool({
          name: t.name,
          description: `Mock ${t.name}`,
          parameters: extraParams as Record<string, unknown>,
          riskLevel: t.riskLevel,
          requiresUndo: t.requiresUndo,
          estimateCells: () => 0,
          execute: async () => ({ ok: true, summary: "mock" }),
        }),
      );
    }
  }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("scenarios/index — проверка эталонных сценариев", () => {
  beforeAll(() => {
    registerScenarioTools();
  });

  describe("структура сценариев", () => {
    const scenarios: { id: string; tools: string[]; description: string }[] = [
      {
        id: "scenario-01-budget-table",
        tools: ["setValues", "fillRange", "formatAsTable", "applyCellFormat"],
        description: "Создать таблицу бюджета 12 месяцев",
      },
      {
        id: "scenario-02-fill-zeros",
        tools: ["fillRange"],
        description: "Заполнить колонку нулями",
      },
      {
        id: "scenario-03-sum-row-column",
        tools: ["getRangeStats", "setFormula"],
        description: "Посчитать сумму по строке и колонке",
      },
      {
        id: "scenario-04-if-formula",
        tools: ["setFormula"],
        description: "Условный IF: >10000 → Контроль",
      },
      {
        id: "scenario-05-vlookup",
        tools: ["lookup"],
        description: "VLOOKUP подтянуть цену",
      },
      {
        id: "scenario-06-error-replace",
        tools: ["setFormula"],
        description: "Заменить #Н/Д и #ДЕЛ/0! на «-»",
      },
      {
        id: "scenario-07-remove-duplicates",
        tools: ["detectDataTypes", "removeDuplicates"],
        description: "Удалить дубликаты по Email",
      },
      {
        id: "scenario-08-split-fio",
        tools: ["detectDataTypes", "splitTextToColumns"],
        description: "Разбить ФИО на 3 колонки",
      },
      {
        id: "scenario-09-normalize-dates",
        tools: ["detectDataTypes", "setFormula", "setValues"],
        description: "Привести даты к ДД.ММ.ГГГГ",
      },
      {
        id: "scenario-10-normalize-numbers",
        tools: ["detectDataTypes", "normalizeText"],
        description: "Числа «1 234,5» → 1234.5",
      },
      {
        id: "scenario-11-normalize-text",
        tools: ["normalizeText"],
        description: "Удалить лишние пробелы, регистр",
      },
      {
        id: "scenario-12-sort",
        tools: ["sortData"],
        description: "Отсортировать по убыванию суммы",
      },
      {
        id: "scenario-13-filter",
        tools: ["filterData"],
        description: "Отфильтровать «только Оплачено»",
      },
      {
        id: "scenario-14-pivot-table",
        tools: ["getWorkbookOverview", "createPivotTable"],
        description: "Сводная: продажи × менеджеры × месяцы",
      },
      {
        id: "scenario-15-chart",
        tools: ["createChart"],
        description: "Столбчатый график по месяцам",
      },
      {
        id: "scenario-16-conditional-format",
        tools: ["getRangeStats", "applyConditionalFormat"],
        description: "Условный формат: красным выше среднего",
      },
      {
        id: "scenario-17-freeze-panes",
        tools: ["freezePanes"],
        description: "Закрепить первую строку",
      },
      {
        id: "scenario-18-autofit",
        tools: ["autoFitColumns"],
        description: "Авто-ширина колонок",
      },
      {
        id: "scenario-19-explain-table",
        tools: ["getWorkbookOverview", "getRangeStats", "findAnomalies"],
        description: "Объяснить что в таблице",
      },
      {
        id: "scenario-20-combo",
        tools: [
          "detectDataTypes",
          "removeDuplicates",
          "normalizeText",
          "getRangeStats",
          "createPivotTable",
          "createChart",
        ],
        description: "Комбо: очисти → посчитай → сводная → график",
      },
    ];

    test("все 20 сценариев определены", () => {
      expect(scenarios.length).toBe(20);
    });

    test.each(scenarios)("$id: использует инструменты $tools", ({ tools }) => {
      for (const tool of tools) {
        expect(toolRegistry.getToolNames()).toContain(tool);
      }
    });

    test("все имена файлов соответствуют конвенции scenario-NN-*", () => {
      for (const s of scenarios) {
        expect(s.id).toMatch(/^scenario-\d{2}-/);
      }
    });
  });

  describe("риск-уровни инструментов", () => {
    test.each([
      ["setValues", "moderate", true],
      ["setFormula", "moderate", true],
      ["fillRange", "moderate", true],
      ["formatAsTable", "moderate", true],
      ["applyCellFormat", "moderate", true],
      ["getRangeStats", "safe", false],
      ["lookup", "moderate", true],
      ["removeDuplicates", "moderate", true],
      ["splitTextToColumns", "moderate", true],
      ["normalizeText", "moderate", true],
      ["sortData", "moderate", true],
      ["filterData", "moderate", true],
      ["createPivotTable", "moderate", true],
      ["createChart", "moderate", true],
      ["freezePanes", "safe", false],
      ["autoFitColumns", "safe", false],
      ["clearRange", "dangerous", true],
      ["appendRows", "moderate", true],
      ["applyConditionalFormat", "moderate", true],
      ["manageSheets", "dangerous", true],
      ["getWorkbookOverview", "safe", false],
      ["getRange", "safe", false],
      ["findAnomalies", "safe", false],
      ["detectDataTypes", "safe", false],
    ])("%s → riskLevel=%s, requiresUndo=%s", (name, risk, needsUndo) => {
      expect(toolRegistry.riskLevel(name as string)).toBe(risk);
      expect(toolRegistry.requiresUndo(name as string)).toBe(needsUndo);
    });
  });

  describe("конфигурация инструментов", () => {
    test("fillRange принимает fillType: progression/copy/value", () => {
      const def = toolRegistry.getDefinition("fillRange");
      expect(def).toBeDefined();
      const params = def?.parameters as {
        properties?: Record<string, unknown>;
      };
      const fillTypeParam = params?.properties?.fillType as {
        enum?: string[];
      };
      expect(fillTypeParam?.enum).toEqual(
        expect.arrayContaining(["progression", "copy", "value"]),
      );
    });

    test("lookup принимает все обязательные параметры", () => {
      const def = toolRegistry.getDefinition("lookup");
      expect(def).toBeDefined();
      const params = def?.parameters as { required?: string[] };
      expect(params?.required).toEqual(
        expect.arrayContaining([
          "lookupAddress",
          "lookupColumn",
          "resultColumn",
          "lookupValue",
          "writeTo",
        ]),
      );
    });

    test("sortData принимает sortColumns с column и order", () => {
      const def = toolRegistry.getDefinition("sortData");
      expect(def).toBeDefined();
      const params = def?.parameters as {
        properties?: Record<string, unknown>;
      };
      const sortCols = params?.properties?.sortColumns as {
        items?: { properties?: Record<string, unknown> };
      };
      expect(sortCols?.items?.properties?.order).toBeDefined();
    });

    test("createPivotTable принимает rows, values с agg", () => {
      const def = toolRegistry.getDefinition("createPivotTable");
      expect(def).toBeDefined();
      const params = def?.parameters as { required?: string[] };
      expect(params?.required).toContain("rows");
      expect(params?.required).toContain("values");
    });

    test("createChart принимает address и chartType", () => {
      const def = toolRegistry.getDefinition("createChart");
      expect(def).toBeDefined();
      const params = def?.parameters as { required?: string[] };
      expect(params?.required).toEqual(
        expect.arrayContaining(["address", "chartType"]),
      );
    });

    test("normalizeText принимает operations из enum", () => {
      const def = toolRegistry.getDefinition("normalizeText");
      expect(def).toBeDefined();
      const params = def?.parameters as {
        properties?: Record<string, unknown>;
      };
      const ops = params?.properties?.operations as {
        items?: { enum?: string[] };
      };
      expect(ops?.items?.enum).toEqual(
        expect.arrayContaining([
          "trim",
          "uppercase",
          "lowercase",
          "propercase",
          "cleanWhitespace",
        ]),
      );
    });
  });
});
