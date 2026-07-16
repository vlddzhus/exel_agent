// Smoke-тест: проверяет тестовое окружение для будущего тестирования tools/ в Фазе 1.
//
// Примечание: для unit-тестов инструментов мы используем тонкий ручной mock Office.js
// (быстрее и изолированнее, чем office-addin-mock, который тянет весь office-addin-manifest).
// office-addin-mock применяется позже для более полных integration-сценариев.
//
// Это шаблон: любой инструмент в Фазе 1 (см. docs/03-TOOLS-SPEC.md §0.1) будет
// тестироваться через тот же механизм — настройка mock-Excel, вызов tool,
// ассерты на side-effects и return value.

// ─── Тип минимального mock-контекста Excel ──────────────────────────────────
interface MockRange {
  address: string;
  values: unknown[][];
  numberFormat: unknown[][];
  load: (props?: string | string[]) => void;
}

interface MockWorksheet {
  getUsedRangeOrNullObject: () => MockRange;
  getRange: (address: string) => MockRange;
}

interface MockExcelContext {
  workbook: {
    worksheets: {
      items: MockWorksheet[];
      getActiveWorksheet: () => MockWorksheet;
    };
  };
  sync: () => Promise<void>;
}

interface MockExcel {
  run: <T = unknown>(
    fn: (ctx: MockExcelContext) => Promise<T>,
  ) => Promise<T>;
}

// ─── Фабрика mock-объекта ───────────────────────────────────────────────────
function createMockRange(address = "A1", values: unknown[][] = [[""]]): MockRange {
  return {
    address,
    values,
    numberFormat: values.map((row) => row.map(() => "General")),
    load: () => {
      /* no-op: mock уже имеет значения в памяти */
    },
  };
}

function createMockExcel(): MockExcel {
  const activeSheet: MockWorksheet = {
    getUsedRangeOrNullObject: () => createMockRange("A1:C3", [["a", "b", "c"]]),
    getRange: (address: string) => createMockRange(address),
  };

  return {
    run: async <T>(fn: (ctx: MockExcelContext) => Promise<T>): Promise<T> => {
      const ctx: MockExcelContext = {
        workbook: {
          worksheets: { items: [activeSheet], getActiveWorksheet: () => activeSheet },
        },
        sync: async () => {
          /* no-op */
        },
      };
      return fn(ctx);
    },
  };
}

// ─── Tests ──────────────────────────────────────────────────────────────────
describe("тестовое окружение для tools/ (smoke)", () => {
  beforeEach(() => {
    // Глобально выставляем mock-Excel — так же, как он доступен в реальном add-in.
    (globalThis as { Excel?: MockExcel }).Excel = createMockExcel();
  });

  test("mock-Excel.run исполняется и sync не падает", async () => {
    const Excel = (globalThis as { Excel: MockExcel }).Excel;
    const result = await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const range = ws.getUsedRangeOrNullObject();
      range.load("address, values");
      await ctx.sync();
      return { address: range.address, rowCount: range.values.length };
    });

    expect(result.address).toBe("A1:C3");
    expect(result.rowCount).toBe(1); // значения из mock-фабрики: 1 строка
  });

  test("вспомогательные функции formula-guardian доступны в тестах", async () => {
    const { columnToLetter, letterToColumn } = await import(
      "../src/taskpane/tools/formula-guardian"
    );
    // columnToLetter — 0-based (см. комментарий в formula-guardian.ts: 0→A)
    expect(columnToLetter(0)).toBe("A");
    expect(columnToLetter(26)).toBe("AA");
    expect(columnToLetter(701)).toBe("ZZ");
    // letterToColumn — обратное преобразование
    expect(letterToColumn("A")).toBe(0);
    expect(letterToColumn("AA")).toBe(26);
    expect(letterToColumn("ZZ")).toBe(701);
  });

  test("FormulaGuardian.validate доступен и работает", async () => {
    const { validateFormula } = await import(
      "../src/taskpane/tools/formula-guardian"
    );
    // Валидная формула проходит
    const ok = validateFormula("=A1*B1+C1");
    expect(ok.valid).toBe(true);
    // Несбалансированные скобки — нет
    const bad = validateFormula("=SUM(A1:A5");
    expect(bad.valid).toBe(false);
  });
});
