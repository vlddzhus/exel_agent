/**
 * sheet-view.ts — Вид листа и настройка печати (категория View/PageSetup).
 *
 * Итерация «Расширение инструментов». Новые инструменты:
 *   SV1 manageSheetView — сетка/заголовки/нули/масштаб листа.
 *   SV2 managePageSetup — параметры печати (ориентация, поля, размер бумаги,
 *                         область печати, вписать в N страниц, колонтитулы).
 *
 * Безопасность: новый файл, существующие инструменты не трогаются.
 * Регистрация через defineTool + toolRegistry.registerDefinition (единый API).
 *
 * Office.js:
 *   - Worksheet.showGridlines / showHeadings / showZeros — boolean.
 *   - Worksheet.zoom = { scale: number } (10-400 проценты).
 *   - Worksheet.pageLayout: orientation, paperSize, margins {top/bottom/left/right/header/footer}.
 *   - Worksheet.pageSetup: printArea, fitToPage, fitToWidth, fitToHeight,
 *     printTitleRows, printTitleColumns.
 */
import { defineTool, toolRegistry, type ToolResult } from "./registry";

// ============================================================================
// SV1. manageSheetView — сетка/заголовки/масштаб
// ============================================================================

export const manageSheetViewTool = defineTool({
  name: "manageSheetView",
  description: `Настраивает вид листа: сетка, заголовки (A,B,C / 1,2,3), нулевые значения, масштаб.
Параметры (все опциональны — указывайте только то, что нужно изменить):
  - showGridlines: показать/скрыть линии сетки (для чистых отчётов/дашбордов — false)
  - showHeadings: показать/скрыть заголовки строк/столбцов
  - showZeros: показывать ли нули в ячейках (false = пустые вместо 0)
  - zoom: масштаб в процентах (10-400, стандартно 100)
Используй для "скрой сетку", "сделай чистый вид", "увеличь масштаб", "спрячь нули".`,
  parameters: {
    type: "object",
    properties: {
      showGridlines: { type: "boolean", description: "Показать/скрыть линии сетки" },
      showHeadings: { type: "boolean", description: "Показать/скрыть заголовки A,B,C / 1,2,3" },
      showZeros: { type: "boolean", description: "Показывать нули в ячейках" },
      zoom: {
        type: "number",
        description: "Масштаб в процентах (10-400)",
      },
      sheetName: { type: "string", description: "Имя листа (опционально)" },
    },
    required: [],
  },
  riskLevel: "safe",
  requiresUndo: false,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    // Хотя бы один параметр должен быть указан
    const hasAny =
      args.showGridlines !== undefined ||
      args.showHeadings !== undefined ||
      args.showZeros !== undefined ||
      args.zoom !== undefined;
    if (!hasAny) {
      return {
        ok: false,
        summary: "Не указано ни одного параметра",
        error: {
          code: "MISSING_PARAMS",
          message: "Укажите хотя бы один из showGridlines/showHeadings/showZeros/zoom",
          retryable: false,
        },
      };
    }

    let zoom = 100;
    if (args.zoom !== undefined) {
      zoom = Number(args.zoom);
      if (!Number.isFinite(zoom) || zoom < 10 || zoom > 400) {
        return {
          ok: false,
          summary: "Некорректный zoom",
          error: {
            code: "INVALID_ZOOM",
            message: "zoom должен быть числом в диапазоне 10-400",
            retryable: false,
          },
        };
      }
    }

    const sheetName = typeof args.sheetName === "string" ? args.sheetName : undefined;

    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");

      if (args.showGridlines !== undefined)
        (sheet as any).showGridlines = args.showGridlines;
      if (args.showHeadings !== undefined)
        (sheet as any).showHeadings = args.showHeadings;
      if (args.showZeros !== undefined)
        (sheet as any).showZeros = args.showZeros;
      if (args.zoom !== undefined) {
        (sheet as any).zoom = { scale: zoom };
      }
      await context.sync();

      return {
        ok: true,
        summary: `Вид листа "${(sheet as any).name}" обновлён`,
        data: {
          sheetName: (sheet as any).name,
          showGridlines: args.showGridlines,
          showHeadings: args.showHeadings,
          showZeros: args.showZeros,
          zoom: args.zoom !== undefined ? zoom : undefined,
        },
      };
    });
  },
});

toolRegistry.registerDefinition(manageSheetViewTool);

// ============================================================================
// SV2. managePageSetup — настройка печати
// ============================================================================

/**
 * Карта общих названий ориентации → Office.js enum значения.
 */
const ORIENTATION_MAP: Record<string, string> = {
  portrait: "Portrait",
  landscape: "Landscape",
};

/**
 * Карта размеров бумаги (общие имена → Office.js PaperType).
 * Полный список см. в Office.js Excel.PaperType.
 */
const PAPER_SIZE_MAP: Record<string, number> = {
  letter: 1,
  letterSmall: 18,
  tabloid: 3,
  ledger: 4,
  legal: 5,
  statement: 6,
  executive: 7,
  a3: 8,
  a4: 9,
  a4Small: 10,
  a5: 11,
  b4: 12,
  b5: 13,
  folio: 14,
};

export const managePageSetupTool = defineTool({
  name: "managePageSetup",
  description: `Настраивает параметры печати листа.
Параметры (все опциональны):
  - orientation: "portrait" (книжная) или "landscape" (альбомная)
  - paperSize: "a4" / "a3" / "letter" / "legal" (по умолчанию A4)
  - margins: объект { top, bottom, left, right, header, footer } — в пунктах (стандартно ~0.75 inch = 54pt)
  - printArea: диапазон области печати "A1:F50"
  - fitToWidth / fitToHeight: вписать в N страниц по ширине/высоте (fitToPage автоматически включается)
  - printTitleRows: "1:1" — повторять строку 1 на каждой странице (для заголовков)
  - centerHorizontally / centerVertically: центрировать при печати
Используй для "сделай альбомную", "настрой печать", "вписать в одну страницу", "поля 1 см".`,
  parameters: {
    type: "object",
    properties: {
      sheetName: { type: "string", description: "Имя листа (опционально)" },
      orientation: {
        type: "string",
        enum: ["portrait", "landscape"],
        description: "Ориентация: portrait (книжная) или landscape (альбомная)",
      },
      paperSize: {
        type: "string",
        description: 'Размер бумаги: "a4", "a3", "letter", "legal"',
      },
      margins: {
        type: "object",
        description: "Поля в пунктах: { top, bottom, left, right, header, footer }",
        properties: {
          top: { type: "number" },
          bottom: { type: "number" },
          left: { type: "number" },
          right: { type: "number" },
          header: { type: "number" },
          footer: { type: "number" },
        },
      },
      printArea: { type: "string", description: 'Область печати: "A1:F50"' },
      fitToWidth: {
        type: "number",
        description: "Вписать в N страниц по ширине (0 = не ограничено)",
      },
      fitToHeight: {
        type: "number",
        description: "Вписать в N страниц по высоте (0 = не ограничено)",
      },
      printTitleRows: {
        type: "string",
        description: 'Повторяемые строки: "1:1" (заголовок на каждой странице)',
      },
      centerHorizontally: { type: "boolean", description: "Центр по горизонтали" },
      centerVertically: { type: "boolean", description: "Центр по вертикали" },
    },
    required: [],
  },
  riskLevel: "moderate",
  requiresUndo: false,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const applied: string[] = [];

    return Excel.run(async (context) => {
      const sheetName = typeof args.sheetName === "string" ? args.sheetName : undefined;
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const pageLayout = (sheet as any).pageLayout;
      const pageSetup = (sheet as any).pageSetup;

      // ── Ориентация ──
      if (args.orientation) {
        const o = ORIENTATION_MAP[String(args.orientation)];
        if (o) {
          pageLayout.orientation = o;
          applied.push(`orientation=${args.orientation}`);
        }
      }

      // ── Размер бумаги ──
      if (args.paperSize) {
        const size = PAPER_SIZE_MAP[String(args.paperSize).toLowerCase()];
        if (size !== undefined) {
          pageLayout.paperSize = size;
          applied.push(`paperSize=${args.paperSize}`);
        }
      }

      // ── Поля ──
      if (args.margins && typeof args.margins === "object") {
        const m = args.margins as Record<string, number>;
        const margins = pageLayout.margins;
        if (margins) {
          for (const key of ["top", "bottom", "left", "right", "header", "footer"]) {
            if (typeof m[key] === "number" && Number.isFinite(m[key])) {
              margins[key] = m[key];
            }
          }
          applied.push("margins");
        }
      }

      // ── Область печати ──
      if (typeof args.printArea === "string" && args.printArea) {
        pageSetup.printArea = args.printArea;
        applied.push(`printArea=${args.printArea}`);
      }

      // ── Вписать в страницы ──
      const fitW = args.fitToWidth;
      const fitH = args.fitToHeight;
      if (fitW !== undefined || fitH !== undefined) {
        pageSetup.fitToPage = true;
        if (typeof fitW === "number") pageSetup.fitToWidth = fitW;
        if (typeof fitH === "number") pageSetup.fitToHeight = fitH;
        applied.push(`fitTo(W=${fitW ?? "auto"},H=${fitH ?? "auto"})`);
      }

      // ── Повторяемые строки ──
      if (typeof args.printTitleRows === "string" && args.printTitleRows) {
        pageSetup.printTitleRows = args.printTitleRows;
        applied.push(`printTitleRows=${args.printTitleRows}`);
      }

      // ── Центрирование ──
      if (args.centerHorizontally !== undefined) {
        pageLayout.centerHorizontally = args.centerHorizontally;
        applied.push(`centerH=${args.centerHorizontally}`);
      }
      if (args.centerVertically !== undefined) {
        pageLayout.centerVertically = args.centerVertically;
        applied.push(`centerV=${args.centerVertically}`);
      }

      if (applied.length === 0) {
        return {
          ok: false,
          summary: "Не указано ни одного параметра page setup",
          error: {
            code: "MISSING_PARAMS",
            message: "Укажите хотя бы один параметр (orientation/paperSize/margins/printArea/fitTo*/printTitleRows/center*)",
            retryable: false,
          },
        };
      }

      await context.sync();

      return {
        ok: true,
        summary: `Настройки печати листа "${(sheet as any).name}" обновлены: ${applied.join(", ")}`,
        data: {
          sheetName: (sheet as any).name,
          applied,
        },
      };
    });
  },
});

toolRegistry.registerDefinition(managePageSetupTool);
