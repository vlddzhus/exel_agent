/**
 * Сценарий №7: «Удалить дубликаты по Email»
 *
 * Инструменты: R4 detectDataTypes, T3 removeDuplicates
 *
 * Сценарий:
 *   1. Определить тип колонки Email через detectDataTypes
 *   2. Удалить дубликаты по колонке Email через removeDuplicates
 */
import { toolRegistry } from "../../../src/taskpane/tools/registry";

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

jest.mock("../../../src/taskpane/tools/_shared/performance", () => {
  const actual = jest.requireActual(
    "../../../src/taskpane/tools/_shared/performance",
  );
  return {
    ...actual,
    withPerformanceGuard: jest
      .fn()
      .mockImplementation(async (cb: (ctx: unknown) => Promise<unknown>) => {
        return Excel.run(async (context) => cb(context));
      }),
  };
});

import "../../../src/taskpane/tools/read";
import "../../../src/taskpane/tools/transform";
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

const EMAIL_DATA = [
  ["Имя", "Email"],
  ["Иван", "ivan@mail.ru"],
  ["Мария", "maria@yandex.ru"],
  ["Иван", "ivan@mail.ru"],
  ["Ольга", "olga@gmail.com"],
  ["Мария", "maria@yandex.ru"],
  ["Петр", "petr@mail.ru"],
];

function setupDupesMock(state: ScenarioState) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });
  const range = createMockRange("TestSheet!A1:B7", state, { rows: 7, cols: 2 });
  Object.defineProperty(range, "values", {
    get: () => EMAIL_DATA,
    set: () => {},
    configurable: true,
  });
  range.rowCount = 7;
  range.columnCount = 2;

  const sheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockImplementation((addr: string) => {
      const r = createMockRange(addr, state);
      Object.defineProperty(r, "values", {
        get: () => EMAIL_DATA,
        set: () => {},
        configurable: true,
      });
      r.rowCount = 7;
      r.columnCount = 2;
      return r;
    }),
    getUsedRangeOrNullObject: jest
      .fn()
      .mockReturnValue(
        Object.assign({}, createMockRange("TestSheet!A1:B7", state), {
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
        return fn({
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue(sheet),
              getItem: jest.fn().mockReturnValue(sheet),
            },
          },
          sync: syncMock,
          runtime: { enableEvents: false },
          application: { calculationMode: "" },
        });
      }),
  };
}

describe("Сценарий 7: Удалить дубликаты по Email", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: detectDataTypes — определить колонку Email", async () => {
    setupDupesMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("detectDataTypes", { address: "A1:B7" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
  });

  test("Шаг 2: removeDuplicates по колонке Email (индекс 1)", async () => {
    setupDupesMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("removeDuplicates", {
        address: "A1:B7",
        columns: [1],
        hasHeaders: true,
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("removeDuplicates без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("removeDuplicates", { columns: [0] }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("removeDuplicates с пустым columns → EMPTY_COLUMNS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("removeDuplicates", {
        address: "A1:B7",
        columns: [],
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("EMPTY_COLUMNS");
  });
});
