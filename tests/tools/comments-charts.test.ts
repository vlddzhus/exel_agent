/**
 * Тесты для comments-charts.ts — manageComments и formatChart.
 * Паттерн как в format-extended.test.ts.
 */
import { toolRegistry } from "../../src/taskpane/tools/registry";

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

import "../../src/taskpane/tools/comments-charts";

/**
 * Mock: range.comments (коллекция) + sheet.charts (коллекция).
 */
function setupExcelMock(opts?: { comments?: Array<{ content: string; authorName?: string; creationDate?: Date }>; charts?: Array<{ name: string }> }) {
  const syncMock = jest.fn().mockResolvedValue(undefined);

  // ── Comments ──
  // Оборачиваем входные комментарии в полноценные объекты с API-методами.
  const commentStore = (opts?.comments ?? []).map((c) => ({
    content: c.content,
    authorName: c.authorName ?? "Test",
    creationDate: c.creationDate ?? new Date(),
    load: jest.fn(),
    delete: jest.fn(),
  }));

  const commentsCollection: any = {
    add: jest.fn().mockImplementation((content: string) => {
      const c = {
        content,
        authorName: "Test",
        creationDate: new Date(),
        load: jest.fn(),
        delete: jest.fn(),
      };
      commentStore.push(c);
      return c;
    }),
    load: jest.fn().mockImplementation(function (this: any, _sel: string) {
      this.items = commentStore.map((c) => ({
        content: c.content,
        authorName: c.authorName,
        creationDate: c.creationDate ?? new Date(),
      }));
      this.count = commentStore.length;
    }),
    getItemAt: jest.fn().mockImplementation((i: number) => commentStore[i]),
    clear: jest.fn().mockImplementation(() => {
      commentStore.length = 0;
    }),
  };

  const makeRange = (addr: string): any => ({
    address: addr,
    rowCount: 1,
    columnCount: 1,
    load: jest.fn(),
    comments: commentsCollection,
    getCell: jest.fn(),
    getResizedRange: jest.fn(),
  });

  // ── Charts ──
  const chartStore = (opts?.charts ?? []).map((c) => c.name);

  const makeChart = (name: string): any => ({
    name,
    load: jest.fn(),
    title: {
      visible: false,
      text: "",
    },
    legend: {
      visible: true,
      position: "Right",
    },
    dataLabels: {
      showValue: false,
      showPercentage: false,
      showCategoryName: false,
    },
    colorScheme: "",
    chartStyle: 0,
    axes: {
      categoryAxis: { title: { text: "" } },
      valueAxis: { title: { text: "" } },
    },
  });

  const chartsCollection: any = {
    load: jest.fn().mockImplementation(function (this: any, _sel: string) {
      this.items = chartStore.map((n) => ({ name: n }));
      this.count = chartStore.length;
    }),
    getItem: jest.fn().mockImplementation((name: string) => {
      if (!chartStore.includes(name)) throw new Error(`Chart not found: ${name}`);
      return makeChart(name);
    }),
    getItemAt: jest.fn().mockImplementation((i: number) => makeChart(chartStore[i] ?? `Chart${i}`)),
  };

  const mockSheet: any = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockImplementation((addr: string) => makeRange(addr)),
    charts: chartsCollection,
  };

  (globalThis as { Excel?: unknown }).Excel = {
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        const ctx = {
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue(mockSheet),
              getItem: jest.fn().mockReturnValue(mockSheet),
            },
          },
          sync: syncMock,
        };
        return fn(ctx);
      }),
  };

  return { syncMock, mockSheet, commentsCollection, chartsCollection, commentStore, chartStore };
}

// ===========================================================================
// manageComments
// ===========================================================================

describe("manageComments", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("add — добавляет комментарий и создаёт undo", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageComments", {
        action: "add",
        address: "A1",
        content: "Важная заметка",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.content).toBe("Важная заметка");
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("add без content → MISSING_CONTENT", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageComments", { action: "add", address: "A1" }),
    );
    expect(r.error?.code).toBe("MISSING_CONTENT");
  });

  test("get — возвращает существующие комментарии", async () => {
    setupExcelMock({ comments: [{ content: "Test1" }, { content: "Test2" }] });
    const r = JSON.parse(
      await toolRegistry.execute("manageComments", { action: "get", address: "A1" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.count).toBe(2);
  });

  test("get — без комментариев count=0", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageComments", { action: "get", address: "A1" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.count).toBe(0);
  });

  test("delete — удаляет первый комментарий", async () => {
    setupExcelMock({ comments: [{ content: "X" }] });
    const r = JSON.parse(
      await toolRegistry.execute("manageComments", { action: "delete", address: "A1" }),
    );
    expect(r.ok).toBe(true);
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("clear — очищает все комментарии", async () => {
    const { commentStore } = setupExcelMock({
      comments: [{ content: "A" }, { content: "B" }],
    });
    const r = JSON.parse(
      await toolRegistry.execute("manageComments", { action: "clear", address: "A1:D10" }),
    );
    expect(r.ok).toBe(true);
    expect(commentStore.length).toBe(0);
  });

  test("без address → MISSING_ADDRESS", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageComments", { action: "get" }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("невалидный action → INVALID_ACTION", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageComments", { action: "edit", address: "A1" }),
    );
    expect(r.error?.code).toBe("INVALID_ACTION");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("manageComments")).toBe("moderate");
    expect(toolRegistry.requiresUndo("manageComments")).toBe(true);
  });
});

// ===========================================================================
// formatChart
// ===========================================================================

describe("formatChart", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("по имени — устанавливает заголовок и легенду", async () => {
    setupExcelMock({ charts: [{ name: "Chart1" }] });
    const r = JSON.parse(
      await toolRegistry.execute("formatChart", {
        chartName: "Chart1",
        title: "Продажи по кварталам",
        legendPosition: "bottom",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.applied).toEqual(expect.arrayContaining([expect.stringContaining("title")]));
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("по индексу — устанавливает подписи данных", async () => {
    setupExcelMock({ charts: [{ name: "Chart1" }, { name: "Chart2" }] });
    const r = JSON.parse(
      await toolRegistry.execute("formatChart", {
        chartIndex: 0,
        dataLabelsShow: "value",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.applied).toContain("dataLabels=value");
  });

  test("colorScheme + chartStyle", async () => {
    setupExcelMock({ charts: [{ name: "Chart1" }] });
    const r = JSON.parse(
      await toolRegistry.execute("formatChart", {
        chartName: "Chart1",
        colorScheme: "#2B579A",
        chartStyle: 10,
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.applied).toContain("colorScheme=#2B579A");
    expect(r.data.applied).toContain("chartStyle=10");
  });

  test("подписи осей", async () => {
    setupExcelMock({ charts: [{ name: "Chart1" }] });
    const r = JSON.parse(
      await toolRegistry.execute("formatChart", {
        chartName: "Chart1",
        axisTitleCategory: "Месяцы",
        axisTitleValue: "Сумма, ₽",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.applied).toContain("axisTitleCategory");
    expect(r.data.applied).toContain("axisTitleValue");
  });

  test("индекс вне диапазона → CHART_INDEX_OUT_OF_RANGE", async () => {
    setupExcelMock({ charts: [{ name: "Chart1" }] });
    const r = JSON.parse(
      await toolRegistry.execute("formatChart", { chartIndex: 5 }),
    );
    expect(r.error?.code).toBe("CHART_INDEX_OUT_OF_RANGE");
  });

  test("несуществующее имя → CHART_NOT_FOUND", async () => {
    setupExcelMock({ charts: [{ name: "Chart1" }] });
    const r = JSON.parse(
      await toolRegistry.execute("formatChart", {
        chartName: "Nope",
        title: "X",
      }),
    );
    expect(r.error?.code).toBe("CHART_NOT_FOUND");
  });

  test("без chartName/chartIndex → MISSING_CHART_REF", async () => {
    setupExcelMock();
    const r = JSON.parse(await toolRegistry.execute("formatChart", { title: "X" }));
    expect(r.error?.code).toBe("MISSING_CHART_REF");
  });

  test("без параметров форматирования → MISSING_PARAMS", async () => {
    setupExcelMock({ charts: [{ name: "Chart1" }] });
    const r = JSON.parse(await toolRegistry.execute("formatChart", { chartName: "Chart1" }));
    expect(r.error?.code).toBe("MISSING_PARAMS");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("formatChart")).toBe("moderate");
    expect(toolRegistry.requiresUndo("formatChart")).toBe(true);
  });
});
