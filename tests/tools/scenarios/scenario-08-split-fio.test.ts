/**
 * Сценарий №8: «Разбить ФИО на 3 колонки»
 *
 * Инструменты: R4 detectDataTypes, T4 splitTextToColumns
 *
 * Сценарий:
 *   1. Проанализировать колонку с ФИО
 *   2. Разбить на Фамилия, Имя, Отчество через splitTextToColumns
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

import "../../../src/taskpane/tools/transform";
import "../../../src/taskpane/tools/read";
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

const FIO_DATA = [
  ["ФИО"],
  ["Иванов Иван Иванович"],
  ["Петрова Мария Сергеевна"],
  ["Сидоров Петр Алексеевич"],
  ["Кузнецова Ольга Владимировна"],
];

function setupFioMock(state: ScenarioState) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });
  const range = createMockRange("TestSheet!A1:A5", state, { rows: 5, cols: 1 });
  Object.defineProperty(range, "values", {
    get: () => FIO_DATA,
    set: () => {},
    configurable: true,
  });
  range.rowCount = 5;
  range.columnCount = 1;

  const sheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockImplementation((addr: string) => {
      const r = createMockRange(addr, state);
      Object.defineProperty(r, "values", {
        get: () => FIO_DATA,
        set: () => {},
        configurable: true,
      });
      r.rowCount = 5;
      if (typeof r.columnCount === "number") r.columnCount = 1;
      return r;
    }),
    getUsedRangeOrNullObject: jest
      .fn()
      .mockReturnValue(
        Object.assign({}, createMockRange("TestSheet!A1:E10", state), {
          isNullObject: false,
        }),
      ),
  };

  (globalThis as { Excel?: unknown }).Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    CalculationMode: { manual: "manual", automatic: "automatic" },
    TextQualifier: { textQualifierDoubleQuote: 1 },
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

describe("Сценарий 8: Разбить ФИО на 3 колонки", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: splitTextToColumns с delimiter=space (ФИО)", async () => {
    setupFioMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("splitTextToColumns", {
        address: "A1:A5",
        delimiter: "space",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("Шаг 2: splitTextToColumns с targetStartCell", async () => {
    setupFioMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("splitTextToColumns", {
        address: "A2:A5",
        delimiter: "space",
        targetStartCell: "B2",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("splitTextToColumns с невалидным delimiter → INVALID_DELIMITER", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("splitTextToColumns", {
        address: "A1:A5",
        delimiter: "tab",
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("INVALID_DELIMITER");
  });

  test("splitTextToColumns без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("splitTextToColumns", { delimiter: "space" }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });
});
