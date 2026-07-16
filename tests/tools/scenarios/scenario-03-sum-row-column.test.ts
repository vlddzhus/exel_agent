/**
 * Эталонный сценарий №3: «Посчитать сумму по строке и колонке»
 *
 * Инструменты: R3 getRangeStats, W2 setFormula
 *
 * Сценарий:
 *   1. Прочитать статистику диапазона через getRangeStats
 *   2. Записать SUM-формулу по колонке через setFormula
 *   3. Записать SUM-формулу по строке через setFormula
 *
 * Данные:
 *   A1:D4 — таблица с числами
 *   | Продукт | Янв | Фев | Мар |
 *   | A       | 100 | 200 | 300 |
 *   | B       | 150 | 250 | 350 |
 *   | C       | 200 | 300 | 400 |
 */
import { toolRegistry } from "../../../src/taskpane/tools/registry";

// ---------------------------------------------------------------------------
// Mocks
// ---------------------------------------------------------------------------

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

// Mock withPerformanceGuard to pass through (setFormula uses it)
jest.mock("../../../src/taskpane/tools/_shared/performance", () => {
  const actual = jest.requireActual(
    "../../../src/taskpane/tools/_shared/performance",
  );
  return {
    ...actual,
    withPerformanceGuard: jest
      .fn()
      .mockImplementation(
        async (callback: (ctx: unknown) => Promise<unknown>) => {
          return Excel.run(async (context) => {
            return callback(context);
          });
        },
      ),
  };
});

import "../../../src/taskpane/tools/read";
import "../../../src/taskpane/tools/write";

import {
  createScenarioState,
  setupExcelMock,
  cleanupExcelMock,
  parseAddressDimensions,
  createMockRange,
  type ScenarioState,
} from "./_mock";

// ---------------------------------------------------------------------------
// Extended mock with pre-set values for getRangeStats
// ---------------------------------------------------------------------------

function setupStatsMock(state: ScenarioState) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });

  const statsRange = createMockRange("TestSheet!A1:D4", state, {
    rows: 4,
    cols: 4,
  });
  // Pre-set values for getRangeStats to read
  const testData = [
    ["Продукт", "Янв", "Фев", "Мар"],
    ["A", 100, 200, 300],
    ["B", 150, 250, 350],
    ["C", 200, 300, 400],
  ];
  // Override the values property
  Object.defineProperty(statsRange, "values", {
    get: () => testData,
    set: () => {},
    configurable: true,
  });
  statsRange.rowCount = 4;
  statsRange.columnCount = 4;

  // General range for setFormula calls
  const formulaRange = createMockRange("TestSheet!E2:E4", state, {
    rows: 3,
    cols: 1,
  });

  const rowSumRange = createMockRange("TestSheet!B5:D5", state, {
    rows: 1,
    cols: 3,
  });

  const sheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockImplementation((addr: string) => {
      if (addr.includes("E")) return formulaRange;
      if (addr.includes("D5") || addr.includes("B5")) return rowSumRange;
      // For getRangeStats, return the range with pre-set values
      const r = createMockRange(addr, state);
      // Copy values from testData if reading A1:D4
      if (addr === "A1:D4" || addr.includes("TestSheet!A1:D4")) {
        Object.defineProperty(r, "values", {
          get: () => testData,
          set: () => {},
          configurable: true,
        });
        r.rowCount = 4;
        r.columnCount = 4;
      }
      return r;
    }),
    getUsedRangeOrNullObject: jest.fn().mockReturnValue(
      Object.assign({}, createMockRange("TestSheet!A1:D10", state), {
        isNullObject: false,
      }),
    ),
  };

  (globalThis as { Excel?: unknown }).Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    CalculationMode: { manual: "manual", automatic: "automatic" },
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        const ctx = {
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue(sheet),
              getItem: jest.fn().mockReturnValue(sheet),
            },
          },
          sync: syncMock,
          runtime: { enableEvents: false },
          application: { calculationMode: "" },
        };
        return fn(ctx);
      }),
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("Сценарий 3: Посчитать сумму по строке и колонке", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });

  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: getRangeStats — прочитать статистику", async () => {
    setupStatsMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("getRangeStats", {
        address: "A1:D4",
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
    const data = r.data as {
      columns: { header: string; sum: number }[];
      rowCount: number;
    };
    expect(data.rowCount).toBe(4);
    expect(data.columns).toHaveLength(4);

    // Column 1 (Янв): 100+150+200 = 450
    expect(data.columns[1].header).toBe("Янв");
    expect(data.columns[1].sum).toBe(450);
  });

  test("Шаг 2: setFormula — SUM по колонке (итог по строкам)", async () => {
    setupStatsMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "E2:E4",
        formulas: [["=SUM(B2:D2)"], ["=SUM(B3:D3)"], ["=SUM(B4:D4)"]],
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(3);
  });

  test("Шаг 3: setFormula — SUM по строке (итог по колонке)", async () => {
    setupStatsMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "B5:D5",
        formulas: [["=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)"]],
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(3);
  });

  test("ПОЛНЫЙ ФЛОУ: stats → колонка SUM → строка SUM", async () => {
    setupStatsMock(state);

    // Step 1: getRangeStats
    const step1 = JSON.parse(
      await toolRegistry.execute("getRangeStats", {
        address: "A1:D4",
      }),
    );
    expect(step1.ok).toBe(true);
    const stats = step1.data as {
      columns: { header: string; sum: number; avg: number }[];
    };
    expect(stats.columns[1].sum).toBe(450); // Янв
    expect(stats.columns[2].sum).toBe(750); // Фев
    expect(stats.columns[3].sum).toBe(1050); // Мар

    // Step 2: write SUM formulas per row
    const step2 = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "E2:E4",
        formulas: [["=SUM(B2:D2)"], ["=SUM(B3:D3)"], ["=SUM(B4:D4)"]],
      }),
    );
    expect(step2.ok).toBe(true);
    expect(state.writtenFormulas).toHaveLength(3);
    expect(state.writtenFormulas[0]?.[0]).toContain("SUM");

    // Step 3: write SUM for column totals
    const step3 = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "B5:D5",
        formulas: [["=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)"]],
      }),
    );
    expect(step3.ok).toBe(true);
  });

  test("Ошибка: setFormula с невалидной формулой (несбалансированные скобки)", async () => {
    setupStatsMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "A1",
        formulas: [["=SUM(("]], // unbalanced parens
      }),
    );

    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("FORMULA_INVALID");
  });

  test("Ошибка: getRangeStats без address", async () => {
    const r = JSON.parse(await toolRegistry.execute("getRangeStats", {}));
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("getRangeStats возвращает корректную структуру данных", async () => {
    setupStatsMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("getRangeStats", {
        address: "A1:D4",
      }),
    );

    expect(r.ok).toBe(true);
    const data = r.data as {
      address: string;
      columns: unknown[];
    };
    expect(data.columns).toHaveLength(4);
    // First column should be text (Продукт) — no numeric values
    expect((data.columns[0] as { numericCount: number }).numericCount).toBe(0);
    // Numeric columns (1-3) should have 3 numeric values each
    expect((data.columns[1] as { numericCount: number }).numericCount).toBe(3);
  });
});
