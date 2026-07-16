/**
 * comments-charts.ts — Комментарии и пост-форматирование графиков.
 *
 * Итерация «Расширение инструментов». Новые инструменты:
 *   CM1 manageComments — add/get/delete/clear комментарии к ячейкам.
 *   CF1 formatChart    — тонкая настройка существующего графика (title,
 *                        legend, dataLabels, оси, цветовая схема, стиль).
 *
 * Безопасность: новый файл, существующие инструменты не трогаются.
 * Регистрация через defineTool + toolRegistry.registerDefinition (единый API).
 *
 * Office.js:
 *   - Range.comments — RangeCommentCollection. .add(content), .getItemAt(i),
 *     .items, .clear(). Comment.content / .author / .creationDate.
 *   - Worksheet.charts — ChartCollection. .getItem(name), .getItemAt(i),
 *     .add(type, range). Chart: title, legend, axes, series, dataLabels,
 *     colorScheme (chart.colorScheme = "#hex"), chartStyle (1-48).
 */
import { defineTool, toolRegistry, type ToolResult } from "./registry";
import { getRangeSafe } from "./address-helper";
import { undoManager } from "./backup";

type CommentAction = "add" | "get" | "delete" | "clear";

// ============================================================================
// CM1. manageComments
// ============================================================================

export const manageCommentsTool = defineTool({
  name: "manageComments",
  description: `Управление комментариями (примечаниями) к ячейкам.
Действия:
  - add: добавить комментарий к ячейке/диапазону. Параметр content (текст).
  - get: прочитать комментарии в диапазоне.
  - delete: удалить конкретный комментарий (content/resolved фильтры не требуются — удаляет первый).
  - clear: очистить все комментарии в диапазоне.
Используй для "добавь комментарий к A1", "поясни расчёт", "оставь заметку в ячейке".`,
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["add", "get", "delete", "clear"],
        description: "Действие: add / get / delete / clear",
      },
      address: {
        type: "string",
        description: 'Адрес ячейки/диапазона: "A1" или "A1:D10"',
      },
      content: {
        type: "string",
        description: "Текст комментария (для add)",
      },
      sheetName: { type: "string", description: "Имя листа (опционально)" },
    },
    required: ["action", "address"],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 1,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const action = String(args.action ?? "") as CommentAction;
    if (!action || !["add", "get", "delete", "clear"].includes(action)) {
      return {
        ok: false,
        summary: "Некорректный action",
        error: {
          code: "INVALID_ACTION",
          message: "action должен быть add/get/delete/clear",
          retryable: false,
        },
      };
    }
    const address = String(args.address ?? "");
    if (!address) {
      return {
        ok: false,
        summary: "address обязателен",
        error: {
          code: "MISSING_ADDRESS",
          message: "address обязателен",
          retryable: false,
        },
      };
    }
    const sheetName = typeof args.sheetName === "string" ? args.sheetName : undefined;

    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = getRangeSafe(context, address);

      // ── ADD ──
      if (action === "add") {
        const content = String(args.content ?? "");
        if (!content) {
          return {
            ok: false,
            summary: "content обязателен для add",
            error: {
              code: "MISSING_CONTENT",
              message: "content (текст комментария) обязателен",
              retryable: false,
            },
          };
        }

        await undoManager.createBackup(address, "manageComments:add", {
          description: `Комментарий в ${address}: ${content.slice(0, 50)}`,
        });

        const comment = (range as any).comments.add(content);
        (comment as any).load("content, authorName");
        await context.sync();
        return {
          ok: true,
          summary: `Добавлен комментарий в ${address}`,
          data: {
            action: "add",
            address,
            content: (comment as any).content,
            author: (comment as any).authorName,
          },
        };
      }

      // ── GET ──
      if (action === "get") {
        const comments = (range as any).comments;
        comments.load("items/content, items/authorName, items/creationDate");
        await context.sync();
        const items = ((comments as any).items ?? []) as any[];
        if (items.length === 0) {
          return {
            ok: true,
            summary: `В ${address} нет комментариев`,
            data: { action: "get", address, count: 0, comments: [] },
          };
        }
        const list = items.map((c) => ({
          content: c.content,
          author: c.authorName,
          created: c.creationDate,
        }));
        return {
          ok: true,
          summary: `Найдено ${list.length} комментариев в ${address}`,
          data: { action: "get", address, count: list.length, comments: list },
        };
      }

      // ── DELETE (первый комментарий) ──
      if (action === "delete") {
        await undoManager.createBackup(address, "manageComments:delete", {
          description: `Удаление комментария из ${address}`,
        });
        try {
          const comments = (range as any).comments;
          comments.load("items/count");
          await context.sync();
          const count = (comments as any).count ?? ((comments as any).items?.length ?? 0);
          if (count === 0) {
            return {
              ok: false,
              summary: `В ${address} нет комментариев`,
              error: {
                code: "NO_COMMENTS",
                message: `Нет комментариев для удаления в ${address}`,
                retryable: false,
              },
            };
          }
          const first = comments.getItemAt(0);
          first.delete();
          await context.sync();
          return {
            ok: true,
            summary: `Удалён комментарий из ${address}`,
            data: { action: "delete", address },
          };
        } catch (e: any) {
          return {
            ok: false,
            summary: `Не удалось удалить комментарий из ${address}`,
            error: {
              code: "DELETE_FAILED",
              message: e?.message ?? "Комментарий не найден",
              retryable: false,
            },
          };
        }
      }

      // ── CLEAR (все комментарии в диапазоне) ──
      await undoManager.createBackup(address, "manageComments:clear", {
        description: `Очистка комментариев в ${address}`,
      });
      (range as any).comments.clear();
      await context.sync();
      return {
        ok: true,
        summary: `Очищены все комментарии в ${address}`,
        data: { action: "clear", address },
      };
    });
  },
});

toolRegistry.registerDefinition(manageCommentsTool);

// ============================================================================
// CF1. formatChart — пост-форматирование существующего графика
// ============================================================================

const LEGEND_POSITION_MAP: Record<string, string> = {
  top: "Top",
  bottom: "Bottom",
  left: "Left",
  right: "Right",
  topRight: "TopRight",
  corner: "Corner",
  none: "None",
};

const DATA_LABEL_MAP: Record<string, boolean> = {
  none: false,
  value: true,
};

/**
 * Map общих имён dataLabels в флаг showValue и showPercentage.
 */
function resolveDataLabels(show: string): { showValue: boolean; showPercentage: boolean; showCategoryName: boolean } {
  switch (show) {
    case "value":
      return { showValue: true, showPercentage: false, showCategoryName: false };
    case "percent":
      return { showValue: false, showPercentage: true, showCategoryName: false };
    case "category":
      return { showValue: false, showPercentage: false, showCategoryName: true };
    case "none":
    default:
      return { showValue: false, showPercentage: false, showCategoryName: false };
  }
}

export const formatChartTool = defineTool({
  name: "formatChart",
  description: `Тонкая настройка существующего графика (после createChart).
График ищется по chartName (имя) или chartIndex (порядковый номер на листе, 0-based).
Параметры (все опциональны — указывайте то, что нужно изменить):
  - title: заголовок графика (текст)
  - titleVisible: показать/скрыть заголовок (true/false)
  - legendVisible: показать/скрыть легенду
  - legendPosition: "top"/"bottom"/"left"/"right"/"topRight"/"corner"/"none"
  - dataLabelsShow: "none"/"value"/"percent"/"category" (подписи точек)
  - colorScheme: hex-цвет основной палитры ("#2B579A")
  - chartStyle: индекс встроенного стиля (1-48)
  - axisTitleCategory: подпись оси X (категорий)
  - axisTitleValue: подпись оси Y (значений)
  - sheetName: имя листа (опционально)
Используй для "добавь заголовок к графику", "покажи значения на столбцах", "поменяй цвет графика".`,
  parameters: {
    type: "object",
    properties: {
      chartName: { type: "string", description: "Имя графика (альтернатива chartIndex)" },
      chartIndex: {
        type: "number",
        description: "Порядковый номер графика на листе (0-based, альтернатива chartName)",
      },
      title: { type: "string", description: "Текст заголовка" },
      titleVisible: { type: "boolean", description: "Показать/скрыть заголовок" },
      legendVisible: { type: "boolean", description: "Показать/скрыть легенду" },
      legendPosition: {
        type: "string",
        enum: ["top", "bottom", "left", "right", "topRight", "corner", "none"],
      },
      dataLabelsShow: {
        type: "string",
        enum: ["none", "value", "percent", "category"],
      },
      colorScheme: { type: "string", description: "Основной цвет: '#2B579A'" },
      chartStyle: {
        type: "number",
        description: "Индекс стиля 1-48",
      },
      axisTitleCategory: { type: "string", description: "Подпись оси X" },
      axisTitleValue: { type: "string", description: "Подпись оси Y" },
      sheetName: { type: "string", description: "Имя листа (опционально)" },
    },
    required: [],
  },
  riskLevel: "moderate",
  requiresUndo: true,
  estimateCells: () => 0,

  async execute(args: Record<string, unknown>): Promise<ToolResult> {
    const chartName = typeof args.chartName === "string" ? args.chartName : undefined;
    const chartIndex = typeof args.chartIndex === "number" ? args.chartIndex : undefined;

    if (chartName === undefined && chartIndex === undefined) {
      return {
        ok: false,
        summary: "Нужно указать chartName или chartIndex",
        error: {
          code: "MISSING_CHART_REF",
          message: "Укажите chartName (имя) или chartIndex (0-based номер)",
          retryable: false,
        },
      };
    }
    const sheetName = typeof args.sheetName === "string" ? args.sheetName : undefined;

    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const charts = (sheet as any).charts;

      await undoManager.createBackup(
        sheetName ?? "active",
        "formatChart",
        { description: `Форматирование графика ${chartName ?? `[${chartIndex}]`}` },
      );

      // Находим график по имени или индексу
      let chart: any;
      try {
        if (chartName) {
          chart = charts.getItem(chartName);
        } else {
          charts.load("items/count, items/name");
          await context.sync();
          const idx = chartIndex ?? 0;
          const count = charts.count ?? (charts.items?.length ?? 0);
          if (idx < 0 || idx >= count) {
            return {
              ok: false,
              summary: `График с индексом ${idx} не найден (всего ${count})`,
              error: {
                code: "CHART_INDEX_OUT_OF_RANGE",
                message: `На листе ${count} графиков. Индекс должен быть 0..${count - 1}`,
                retryable: false,
              },
            };
          }
          chart = charts.getItemAt(idx);
        }
        chart.load("name");
        await context.sync();
      } catch (e: any) {
        return {
          ok: false,
          summary: `График "${chartName ?? chartIndex}" не найден`,
          error: {
            code: "CHART_NOT_FOUND",
            message: e?.message ?? "График не существует на листе",
            retryable: false,
          },
        };
      }

      const applied: string[] = [];

      // ── Заголовок ──
      if (args.titleVisible !== undefined || args.title !== undefined) {
        const title = chart.title;
        if (args.titleVisible !== undefined) title.visible = args.titleVisible;
        if (typeof args.title === "string") {
          title.visible = true;
          title.text = args.title;
        }
        applied.push(`title=${args.title ?? (args.titleVisible ? "visible" : "hidden")}`);
      }

      // ── Легенда ──
      if (args.legendVisible !== undefined || args.legendPosition) {
        const legend = chart.legend;
        if (args.legendVisible !== undefined) legend.visible = args.legendVisible;
        if (args.legendPosition) {
          const legendPos = String(args.legendPosition);
          const pos = LEGEND_POSITION_MAP[legendPos];
          if (pos) {
            legend.visible = pos === "None" ? false : true;
            legend.position = pos;
          }
        }
        applied.push(`legend=${args.legendPosition ?? (args.legendVisible ? "visible" : "hidden")}`);
      }

      // ── Подписи данных ──
      if (args.dataLabelsShow) {
        const labels = chart.dataLabels;
        const dl = resolveDataLabels(String(args.dataLabelsShow));
        labels.showValue = dl.showValue;
        labels.showPercentage = dl.showPercentage;
        labels.showCategoryName = dl.showCategoryName;
        applied.push(`dataLabels=${args.dataLabelsShow}`);
      }

      // ── Цветовая схема ──
      if (typeof args.colorScheme === "string" && args.colorScheme) {
        chart.colorScheme = args.colorScheme;
        applied.push(`colorScheme=${args.colorScheme}`);
      }

      // ── Встроенный стиль (1-48) ──
      if (typeof args.chartStyle === "number") {
        const style = Math.max(1, Math.min(48, Math.floor(args.chartStyle)));
        chart.chartStyle = style;
        applied.push(`chartStyle=${style}`);
      }

      // ── Подписи осей ──
      if (typeof args.axisTitleCategory === "string") {
        const axis = chart.axes.categoryAxis;
        axis.title.text = args.axisTitleCategory;
        applied.push(`axisTitleCategory`);
      }
      if (typeof args.axisTitleValue === "string") {
        const axis = chart.axes.valueAxis;
        axis.title.text = args.axisTitleValue;
        applied.push(`axisTitleValue`);
      }

      if (applied.length === 0) {
        return {
          ok: false,
          summary: "Не указано ни одного параметра форматирования",
          error: {
            code: "MISSING_PARAMS",
            message: "Укажите хотя бы один параметр (title/legend/dataLabels/colorScheme/chartStyle/axisTitle*)",
            retryable: false,
          },
        };
      }

      await context.sync();

      return {
        ok: true,
        summary: `График "${chart.name}" обновлён: ${applied.join(", ")}`,
        data: {
          chartName: chart.name,
          applied,
        },
      };
    });
  },
});

toolRegistry.registerDefinition(formatChartTool);
