/**
 * Тесты для sheet-view.ts — manageSheetView и managePageSetup.
 * Паттерн как в format-extended.test.ts.
 */
import { toolRegistry } from "../../src/taskpane/tools/registry";

// manageSheetView/managePageSetup не используют undoManager, но мок нужен,
// т.к. barrel-импорт мог бы его загрузить.
jest.mock("../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: jest.fn().mockResolvedValue(undefined),
  },
}));

import "../../src/taskpane/tools/sheet-view";

function setupExcelMock() {
  const syncMock = jest.fn().mockResolvedValue(undefined);

  const pageLayout: any = {
    orientation: "Portrait",
    paperSize: 9,
    centerHorizontally: false,
    centerVertically: false,
    margins: { top: 54, bottom: 54, left: 54, right: 54, header: 27, footer: 27 },
  };
  const pageSetup: any = {
    printArea: "",
    fitToPage: false,
    fitToWidth: 0,
    fitToHeight: 0,
    printTitleRows: "",
  };

  const mockSheet: any = {
    name: "TestSheet",
    load: jest.fn(),
    showGridlines: true,
    showHeadings: true,
    showZeros: true,
    zoom: { scale: 100 },
    pageLayout,
    pageSetup,
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

  return { syncMock, mockSheet, pageLayout, pageSetup };
}

// ===========================================================================
// manageSheetView
// ===========================================================================

describe("manageSheetView", () => {
  beforeEach(() => setupExcelMock());
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("скрыть сетку", async () => {
    const { mockSheet } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageSheetView", { showGridlines: false }),
    );
    expect(r.ok).toBe(true);
    expect(mockSheet.showGridlines).toBe(false);
  });

  test("установить масштаб", async () => {
    const { mockSheet } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageSheetView", { zoom: 150 }),
    );
    expect(r.ok).toBe(true);
    expect(mockSheet.zoom.scale).toBe(150);
  });

  test("скрыть заголовки и нули вместе", async () => {
    const { mockSheet } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageSheetView", {
        showHeadings: false,
        showZeros: false,
      }),
    );
    expect(r.ok).toBe(true);
    expect(mockSheet.showHeadings).toBe(false);
    expect(mockSheet.showZeros).toBe(false);
  });

  test("невалидный zoom (500) → INVALID_ZOOM", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageSheetView", { zoom: 500 }),
    );
    expect(r.error?.code).toBe("INVALID_ZOOM");
  });

  test("без параметров → MISSING_PARAMS", async () => {
    setupExcelMock();
    const r = JSON.parse(await toolRegistry.execute("manageSheetView", {}));
    expect(r.error?.code).toBe("MISSING_PARAMS");
  });

  test("риск safe, не требует undo", () => {
    expect(toolRegistry.riskLevel("manageSheetView")).toBe("safe");
    expect(toolRegistry.requiresUndo("manageSheetView")).toBe(false);
  });
});

// ===========================================================================
// managePageSetup
// ===========================================================================

describe("managePageSetup", () => {
  beforeEach(() => setupExcelMock());
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("landscape ориентация + A3", async () => {
    const { pageLayout } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("managePageSetup", {
        orientation: "landscape",
        paperSize: "a3",
      }),
    );
    expect(r.ok).toBe(true);
    expect(pageLayout.orientation).toBe("Landscape");
    expect(pageLayout.paperSize).toBe(8);
  });

  test("поля и область печати", async () => {
    const { pageLayout, pageSetup } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("managePageSetup", {
        margins: { top: 30, bottom: 30, left: 20, right: 20 },
        printArea: "A1:F50",
      }),
    );
    expect(r.ok).toBe(true);
    expect(pageLayout.margins.top).toBe(30);
    expect(pageSetup.printArea).toBe("A1:F50");
  });

  test("вписать в 1 страницу по ширине", async () => {
    const { pageSetup } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("managePageSetup", { fitToWidth: 1 }),
    );
    expect(r.ok).toBe(true);
    expect(pageSetup.fitToPage).toBe(true);
    expect(pageSetup.fitToWidth).toBe(1);
  });

  test("повторяемые строки заголовка", async () => {
    const { pageSetup } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("managePageSetup", { printTitleRows: "1:1" }),
    );
    expect(r.ok).toBe(true);
    expect(pageSetup.printTitleRows).toBe("1:1");
  });

  test("центрирование по горизонтали", async () => {
    const { pageLayout } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("managePageSetup", { centerHorizontally: true }),
    );
    expect(r.ok).toBe(true);
    expect(pageLayout.centerHorizontally).toBe(true);
  });

  test("без параметров → MISSING_PARAMS", async () => {
    setupExcelMock();
    const r = JSON.parse(await toolRegistry.execute("managePageSetup", {}));
    expect(r.error?.code).toBe("MISSING_PARAMS");
  });

  test("риск moderate, не требует undo", () => {
    expect(toolRegistry.riskLevel("managePageSetup")).toBe("moderate");
    expect(toolRegistry.requiresUndo("managePageSetup")).toBe(false);
  });
});
