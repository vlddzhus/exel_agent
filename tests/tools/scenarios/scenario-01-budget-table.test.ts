/**
 * Эталонный сценарий №1: «Создать таблицу бюджета 12 месяцев»
 *
 * Инструменты: W1 setValues, W3 fillRange, F4 formatAsTable, F1 applyCellFormat
 *
 * Сценарий:
 *   1. Заполнить колонку A названиями месяцев (fillRange, progression 1..12)
 *   2. Записать бюджетные статьи и цифры (setValues)
 *   3. Превратить диапазон в таблицу (formatAsTable)
 *   4. Применить жирный шрифт к заголовкам (applyCellFormat)
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

// Mirror existing test pattern: mock withPerformanceGuard to pass through
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

import "../../../src/taskpane/tools/write";
import "../../../src/taskpane/tools/format";

import {
  createScenarioState,
  setupExcelMock,
  cleanupExcelMock,
  type ScenarioState,
} from "./_mock";

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("Сценарий 1: Создать таблицу бюджета 12 месяцев", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });

  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: fillRange — заполнить колонку A месяцами (прогрессия 1..12)", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("fillRange", {
        address: "A1:A12",
        fillType: "progression",
        startValue: 1,
        step: 1,
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(12);
    expect(r.summary).toContain("12");
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("Шаг 2: setValues — записать бюджетные статьи и цифры", async () => {
    setupExcelMock(state);

    const budgetData = [
      ["Категория", "Янв", "Фев", "Мар"],
      ["Аренда", 50000, 50000, 50000],
      ["Зарплата", 300000, 300000, 300000],
      ["Маркетинг", 15000, 20000, 18000],
      ["Софт", 8000, 8000, 8000],
    ];

    const r = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "A1:D5",
        values: budgetData,
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(20);
    expect(r.data).toMatchObject({ rows: 5, cols: 4 });
  });

  test("Шаг 3: formatAsTable — превратить диапазон в таблицу", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("formatAsTable", {
        address: "A1:D13",
        hasHeaders: true,
        tableName: "Бюджет",
        style: "TableStyleMedium2",
      }),
    );

    expect(r.ok).toBe(true);
    expect(state.tableCreated).toBe(true);
    expect(r.data).toBeDefined();
    if (r.data) {
      expect((r.data as { name?: string }).name).toBeDefined();
    }
  });

  test("Шаг 4: applyCellFormat — жирный шрифт для заголовков", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("applyCellFormat", {
        address: "A1:D1",
        format: {
          bold: true,
          fontSize: 12,
          fillColor: "#4472C4",
          fontColor: "#FFFFFF",
          horizontalAlignment: "center",
        },
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.summary).toContain("A1:D1");
  });

  test("ПОЛНЫЙ ФЛОУ: все 4 шага последовательно", async () => {
    setupExcelMock(state);

    // Step 1: fill months (progression 1..12 → Jan..Dec)
    const step1 = JSON.parse(
      await toolRegistry.execute("fillRange", {
        address: "A1:A12",
        fillType: "progression",
        startValue: 1,
        step: 1,
      }),
    );
    expect(step1.ok).toBe(true);

    // Step 2: write budget data (headers + 4 categories × 3 months)
    const budgetData = [
      ["Категория", "Янв", "Фев", "Мар"],
      ["Аренда", 50000, 50000, 50000],
      ["Зарплата", 300000, 300000, 300000],
      ["Маркетинг", 15000, 20000, 18000],
      ["Софт", 8000, 8000, 8000],
    ];
    const step2 = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "B1:E5",
        values: budgetData,
      }),
    );
    expect(step2.ok).toBe(true);

    // Step 3: format as Excel table
    const step3 = JSON.parse(
      await toolRegistry.execute("formatAsTable", {
        address: "A1:E12",
        hasHeaders: true,
        tableName: "Бюджет2024",
      }),
    );
    expect(step3.ok).toBe(true);
    expect(state.tableCreated).toBe(true);

    // Step 4: bold header row with blue fill
    const step4 = JSON.parse(
      await toolRegistry.execute("applyCellFormat", {
        address: "A1:E1",
        format: { bold: true, fillColor: "#4472C4", fontColor: "#FFFFFF" },
      }),
    );
    expect(step4.ok).toBe(true);

    // Verify undo was called for all write operations
    expect(mockCreateBackup).toHaveBeenCalledTimes(3); // fillRange, setValues, formatAsTable
  });

  test("Ошибка: formatAsTable без address", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("formatAsTable", {
        hasHeaders: true,
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("Ошибка: setValues с пустыми values", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "A1",
        values: [],
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("EMPTY_VALUES");
  });

  test("fillRange возвращает корректный cellsAffected для 12×1", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("fillRange", {
        address: "A1:A12",
        fillType: "progression",
        startValue: 1,
        step: 1,
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(12);
    expect(r.data).toMatchObject({
      rows: 12,
      cols: 1,
      fillType: "progression",
    });
  });
});
