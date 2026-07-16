/**
 * Сценарий №11: «Удалить лишние пробелы, регистр»
 *
 * Инструменты: T5 normalizeText
 *
 * Сценарий:
 *   1. Текст с лишними пробелами → cleanWhitespace
 *   2. Верхний регистр → uppercase
 *   3. Каждое слово с заглавной → propercase
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
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

const DIRTY_TEXT = [
  ["  Иван  ", "иванов"],
  ["  МАРИЯ  ", "петрова"],
  ["Ольга  ", "  СИДОРОВА"],
  ["  петр  ", "кузнецов  "],
];

function setupTextMock(state: ScenarioState) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });

  (globalThis as { Excel?: unknown }).Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    CalculationMode: { manual: "manual", automatic: "automatic" },
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        return fn({
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue({
                name: "TestSheet",
                load: jest.fn(),
                getRange: jest.fn().mockImplementation((addr: string) => {
                  const r = createMockRange(addr, state);
                  Object.defineProperty(r, "values", {
                    get: () => DIRTY_TEXT,
                    set: (v: unknown[][]) => {
                      state.writtenValues = v;
                    },
                    configurable: true,
                  });
                  r.rowCount = 4;
                  r.columnCount = 2;
                  return r;
                }),
                getUsedRangeOrNullObject: jest.fn(),
              }),
              getItem: jest.fn(),
            },
          },
          sync: syncMock,
          runtime: { enableEvents: false },
          application: { calculationMode: "" },
        });
      }),
  };
}

describe("Сценарий 11: Удалить лишние пробелы, регистр", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: cleanWhitespace — убрать лишние пробелы", async () => {
    setupTextMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "A1:A4",
        operations: ["cleanWhitespace"],
      }),
    );
    expect(r.ok).toBe(true);
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("Шаг 2: uppercase", async () => {
    setupTextMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "B1:B4",
        operations: ["uppercase"],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("Шаг 3: propercase (каждое слово с заглавной)", async () => {
    setupTextMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "A1:A4",
        operations: ["propercase"],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("Шаг 4: lowercase", async () => {
    setupTextMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "B1:B4",
        operations: ["lowercase"],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("ПОЛНЫЙ ФЛОУ: trim + propercase для ФИО", async () => {
    setupTextMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "A1:B4",
        operations: ["trim", "propercase"],
      }),
    );
    expect(r.ok).toBe(true);
  });
});
