/**
 * Эталонный сценарий №5: «VLOOKUP подтянуть цену»
 *
 * Инструменты: T6 lookup, W1 setValues, F4 formatAsTable
 *
 * Сценарий:
 *   1. Создать две таблицы:
 *      - Таблица товаров A1:B5 (Код, Товар)
 *      - Таблица прайса D1:E5 (Код, Цена)
 *   2. Через lookup найти цену для каждого товара
 *   3. Записать результат в колонку C
 *
 * Структура:
 *   A (Код)  B (Товар)   C (Цена→)    D (Код)  E (Цена)
 *   T001     Ноутбук     ?(lookup)    T001     75000
 *   T002     Мышь        ?(lookup)    T002     1500
 *   T003     Клавиатура  ?(lookup)    T003     3500
 *   T004     Монитор     ?(lookup)    T004     25000
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

// Mock withPerformanceGuard for setValues
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

import "../../../src/taskpane/tools/transform";
import "../../../src/taskpane/tools/write";
import "../../../src/taskpane/tools/format";

import {
  createScenarioState,
  setupExcelMock,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

// ---------------------------------------------------------------------------
// Lookup-aware mock (needs pre-set price table data)
// ---------------------------------------------------------------------------

const PRICE_TABLE = [
  ["Код", "Цена"],
  ["T001", 75000],
  ["T002", 1500],
  ["T003", 3500],
  ["T004", 25000],
];

function setupLookupMock(state: ScenarioState) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });

  const sheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockImplementation((addr: string) => {
      const r = createMockRange(addr, state);
      // If this is the price lookup range, set the values
      if (addr.includes("D1:E5") || addr.includes("D:E")) {
        Object.defineProperty(r, "values", {
          get: () => PRICE_TABLE,
          set: () => {},
          configurable: true,
        });
        r.rowCount = PRICE_TABLE.length;
        r.columnCount = 2;
      }
      return r;
    }),
    getUsedRangeOrNullObject: jest.fn().mockReturnValue(
      Object.assign({}, createMockRange("TestSheet!A1:E10", state), {
        isNullObject: false,
      }),
    ),
    tables: {
      add: jest.fn().mockReturnValue({
        name: "Table",
        style: "TableStyleLight1",
        load: jest.fn(),
      }),
      load: jest.fn(),
    },
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

describe("Сценарий 5: VLOOKUP подтянуть цену", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });

  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: setValues — создать таблицу-прайс", async () => {
    setupLookupMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "D1:E5",
        values: PRICE_TABLE,
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(10);
  });

  test("Шаг 2: lookup — найти цену T001 → 75000", async () => {
    setupLookupMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("lookup", {
        lookupAddress: "D1:E5",
        lookupColumn: 0,
        resultColumn: 1,
        lookupValue: "T001",
        writeTo: "C2",
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
    if (r.data) {
      expect((r.data as { result: unknown }).result).toBe(75000);
    }
  });

  test("Шаг 3: lookup — найти цену T003 → 3500", async () => {
    setupLookupMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("lookup", {
        lookupAddress: "D1:E5",
        lookupColumn: 0,
        resultColumn: 1,
        lookupValue: "T003",
        writeTo: "C4",
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
    if (r.data) {
      expect((r.data as { result: unknown }).result).toBe(3500);
    }
  });

  test("ПОЛНЫЙ ФЛОУ: создать обе таблицы → lookup для всех товаров", async () => {
    setupLookupMock(state);

    // Step 1: product table
    const step1 = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "A1:B5",
        values: [
          ["Код", "Товар"],
          ["T001", "Ноутбук"],
          ["T002", "Мышь"],
          ["T003", "Клавиатура"],
          ["T004", "Монитор"],
        ],
      }),
    );
    expect(step1.ok).toBe(true);

    // Step 2: price table
    const step2 = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "D1:E5",
        values: PRICE_TABLE,
      }),
    );
    expect(step2.ok).toBe(true);

    // Step 3-6: lookup each product price
    const products = [
      { code: "T001", expectedPrice: 75000, cell: "C2" },
      { code: "T002", expectedPrice: 1500, cell: "C3" },
      { code: "T003", expectedPrice: 3500, cell: "C4" },
      { code: "T004", expectedPrice: 25000, cell: "C5" },
    ];

    for (const p of products) {
      const r = JSON.parse(
        await toolRegistry.execute("lookup", {
          lookupAddress: "D1:E5",
          lookupColumn: 0,
          resultColumn: 1,
          lookupValue: p.code,
          writeTo: p.cell,
        }),
      );
      expect(r.ok).toBe(true);
      expect((r.data as { result: unknown }).result).toBe(p.expectedPrice);
    }

    expect(mockCreateBackup).toHaveBeenCalledTimes(6); // 2× setValues + 4× lookup
  });

  test("lookup — значение не найдено (T999)", async () => {
    setupLookupMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("lookup", {
        lookupAddress: "D1:E5",
        lookupColumn: 0,
        resultColumn: 1,
        lookupValue: "T999",
        writeTo: "C6",
      }),
    );

    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("NOT_FOUND");
  });

  test("lookup без lookupAddress → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("lookup", {
        lookupColumn: 0,
        resultColumn: 1,
        lookupValue: "T001",
        writeTo: "C2",
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("lookup без lookupValue → MISSING_VALUE", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("lookup", {
        lookupAddress: "D1:E5",
        lookupColumn: 0,
        resultColumn: 1,
        writeTo: "C2",
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_VALUE");
  });

  test("lookup — регистрозависимость ('t001' vs 'T001')", async () => {
    setupLookupMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("lookup", {
        lookupAddress: "D1:E5",
        lookupColumn: 0,
        resultColumn: 1,
        lookupValue: "t001", // lowercase
        writeTo: "C2",
      }),
    );

    // lookup does exact string match — "t001" !== "T001"
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("NOT_FOUND");
  });

  test("ПОЛНЫЙ ФЛОУ + форматирование таблиц", async () => {
    setupLookupMock(state);

    // Write product table
    const step1 = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "A1:B5",
        values: [
          ["Код", "Товар"],
          ["T001", "Ноутбук"],
          ["T002", "Мышь"],
          ["T003", "Клавиатура"],
          ["T004", "Монитор"],
        ],
      }),
    );
    expect(step1.ok).toBe(true);

    // Write price table
    const step2 = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "D1:E5",
        values: PRICE_TABLE,
      }),
    );
    expect(step2.ok).toBe(true);

    // Format both as tables
    const step3 = JSON.parse(
      await toolRegistry.execute("formatAsTable", {
        address: "A1:B5",
        hasHeaders: true,
        tableName: "Товары",
      }),
    );
    expect(step3.ok).toBe(true);

    const step4 = JSON.parse(
      await toolRegistry.execute("formatAsTable", {
        address: "D1:E5",
        hasHeaders: true,
        tableName: "Прайс",
      }),
    );
    expect(step4.ok).toBe(true);

    // Lookup all prices
    for (const { code, cell } of [
      { code: "T001", cell: "C2" },
      { code: "T002", cell: "C3" },
      { code: "T003", cell: "C4" },
      { code: "T004", cell: "C5" },
    ]) {
      const r = JSON.parse(
        await toolRegistry.execute("lookup", {
          lookupAddress: "D1:E5",
          lookupColumn: 0,
          resultColumn: 1,
          lookupValue: code,
          writeTo: cell,
        }),
      );
      expect(r.ok).toBe(true);
    }
  });
});
