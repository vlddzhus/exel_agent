/**
 * Сценарий №6: «Заменить #Н/Д и #ДЕЛ/0! на «-»»
 *
 * Инструменты: W2 setFormula
 *
 * Сценарий:
 *   1. Есть таблица с #Н/Д и #ДЕЛ/0! ошибками (результат VLOOKUP/деления)
 *   2. Через IFERROR/ЕСЛИОШИБКА обернуть существующие формулы
 *   3. FormulaGuardian пропускает IFERROR
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

import "../../../src/taskpane/tools/write";
import {
  createScenarioState,
  setupExcelMock,
  cleanupExcelMock,
  type ScenarioState,
} from "./_mock";

describe("Сценарий 6: Заменить #Н/Д и #ДЕЛ/0! на «-»", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });

  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: IFERROR оборачивает VLOOKUP", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "C2",
        formulas: [['=IFERROR(VLOOKUP(A2,$E$2:$F$10,2,FALSE),"-")']],
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(1);
  });

  test("Шаг 2: IFERROR оборачивает деление", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "D2",
        formulas: [['=IFERROR(B2/C2,"-")']],
      }),
    );

    expect(r.ok).toBe(true);
  });

  test("Шаг 3: массовая замена IFERROR на диапазон", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "C2:C10",
        formulas: Array(9).fill([
          '=IFERROR(VLOOKUP(A2,$E$2:$F$10,2,FALSE),"-")',
        ]),
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(9);
  });

  test("Шаг 4: русская ЕСЛИОШИБКА", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "C2",
        formulas: [['=ЕСЛИОШИБКА(ВПР(A2;$E$2:$F$10;2;ЛОЖЬ);"-")']],
      }),
    );

    expect(r.ok).toBe(true);
  });

  test("ПОЛНЫЙ ФЛОУ: данные → IFERROR для всех VLOOKUP", async () => {
    setupExcelMock(state);

    const step1 = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "A1:B5",
        values: [
          ["Код", "Сумма"],
          ["T001", 10000],
          ["T002", 25000],
          ["T003", 15000],
          ["T004", 30000],
        ],
      }),
    );
    expect(step1.ok).toBe(true);

    const step2 = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "C2:C5",
        formulas: Array(4).fill(['=IFERROR(B2*0.13,"-")']),
      }),
    );
    expect(step2.ok).toBe(true);

    for (const row of state.writtenFormulas) {
      for (const f of row) {
        expect(f).toMatch(/^=IFERROR\(/);
      }
    }
  });

  test("setFormula без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        formulas: [['=IFERROR(A1,"-")']],
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });
});
