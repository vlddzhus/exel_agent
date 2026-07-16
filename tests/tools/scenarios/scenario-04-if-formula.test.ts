/**
 * Эталонный сценарий №4: «Условный IF: >10000 → Контроль»
 *
 * Инструменты: W2 setFormula
 *
 * Сценарий:
 *   1. В колонке B — числовые значения (суммы)
 *   2. В колонку C записать IF-формулы: =IF(B2>10000,"Контроль","")
 *   3. FormulaGuardian пропускает валидный IF
 *   4. Опасные функции (WEBSERVICE) блокируются
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

// Mock withPerformanceGuard for setFormula
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

import {
  createScenarioState,
  setupExcelMock,
  cleanupExcelMock,
  type ScenarioState,
} from "./_mock";

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("Сценарий 4: Условный IF: >10000 → Контроль", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });

  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: setFormula с одним IF (>10000)", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "C2",
        formulas: [['=IF(B2>10000,"Контроль","")']],
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(1);
  });

  test("Шаг 2: setFormula с IF на весь диапазон", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "C2:C10",
        formulas: [
          ['=IF(B2>10000,"Контроль","")'],
          ['=IF(B3>10000,"Контроль","")'],
          ['=IF(B4>10000,"Контроль","")'],
          ['=IF(B5>10000,"Контроль","")'],
          ['=IF(B6>10000,"Контроль","")'],
          ['=IF(B7>10000,"Контроль","")'],
          ['=IF(B8>10000,"Контроль","")'],
          ['=IF(B9>10000,"Контроль","")'],
          ['=IF(B10>10000,"Контроль","")'],
        ],
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(9);
  });

  test("IF с вложенным AND/OR", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "D2",
        formulas: [['=IF(AND(B2>5000,C2<1000),"Внимание","Норма")']],
      }),
    );

    expect(r.ok).toBe(true);
  });

  test("IF с русскими именами функций (ЕСЛИ)", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "E2",
        formulas: [['=ЕСЛИ(B2>10000;"Контроль";"")']],
      }),
    );

    expect(r.ok).toBe(true);
  });

  test("ПОЛНЫЙ ФЛОУ: данные → IF-формулы для каждой строки", async () => {
    setupExcelMock(state);

    // Step 1: Write values
    const step1 = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "B1:B5",
        values: [["Сумма"], [5000], [15000], [8000], [20000]],
      }),
    );
    expect(step1.ok).toBe(true);

    // Step 2: Write IF formulas
    const step2 = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "C2:C5",
        formulas: [
          ['=IF(B2>10000,"Контроль","")'],
          ['=IF(B3>10000,"Контроль","")'],
          ['=IF(B4>10000,"Контроль","")'],
          ['=IF(B5>10000,"Контроль","")'],
        ],
      }),
    );
    expect(step2.ok).toBe(true);

    // Verify all formulas start with IF
    for (const row of state.writtenFormulas) {
      for (const formula of row) {
        expect(formula).toMatch(/^=IF\(/);
      }
    }
  });

  test("FormulaGuardian блокирует WEBSERVICE", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "A1",
        formulas: [['=WEBSERVICE("http://evil.com/")']],
      }),
    );

    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("FORMULA_BLOCKED");
  });

  test("FormulaGuardian блокирует несбалансированные скобки", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "A1",
        formulas: [['=IF(B2>10000,"Контроль","']], // unclosed quote
      }),
    );

    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("FORMULA_INVALID");
  });

  test("FormulaGuardian авто-исправляет '*' между смежными ссылками", async () => {
    setupExcelMock(state);

    // B2B3 will be auto-fixed to B2*B3
    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "A1",
        formulas: [["=B2B3+100"]],
      }),
    );

    expect(r.ok).toBe(true);
  });

  test("setFormula без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        formulas: [['=IF(1>0,"ok","")']],
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("setFormula с пустым formulas → EMPTY_FORMULAS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "A1",
        formulas: [],
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("EMPTY_FORMULAS");
  });
});
